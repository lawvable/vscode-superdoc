import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as https from 'https';
import * as http from 'http';

// Debug logging - set to false to disable verbose logs
const DEBUG_ENABLED = true;

function debug(message: string) {
  if (DEBUG_ENABLED) {
    console.log('[SuperDoc - Provider]', message);
  }
}

// Command folder for Claude API
const SUPERDOC_FOLDER = '.superdoc';

export class SuperDocEditorProvider implements vscode.CustomEditorProvider<SuperDocDocument> {
  public static readonly viewType = 'superdoc.docxEditor';

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SuperDocDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<SuperDocDocument> {
    debug(`Opening DOCX file: ${uri.fsPath}`);
    const fileUri = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const data = await vscode.workspace.fs.readFile(fileUri);
    return new SuperDocDocument(uri, data);
  }

  async resolveCustomEditor(
    document: SuperDocDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    debug(`Resolving custom editor for: ${document.uri.fsPath}`);

    webviewPanel.webview.options = {
      enableScripts: true,
      localResourceRoots: [
        vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview'),
      ],
    };

    webviewPanel.webview.html = this.getWebviewContent(webviewPanel.webview);
    this.setupMessageHandler(webviewPanel.webview, document);

    // Send document data when webview is ready
    const readyListener = webviewPanel.webview.onDidReceiveMessage((message) => {
      if (message.type === 'ready') {
        debug(`Sending document to webview, size: ${document.data.length}`);
        webviewPanel.webview.postMessage({
          type: 'update',
          content: { data: Array.from(document.data) },
        });
      }
    });

    // Watch for external file changes
    const fileWatcher = this.setupFileWatcher(document, webviewPanel.webview);

    // Watch for command file (Claude API)
    const commandWatcher = this.setupCommandWatcher(document, webviewPanel.webview);

    const cmdFilePath = this.getCommandFilePath(document.uri);

    webviewPanel.onDidDispose(() => {
      readyListener.dispose();
      fileWatcher.dispose();
      if (commandWatcher) {
        commandWatcher.close();
      }
      // Clean up per-document state
      const state = this._commandStates.get(cmdFilePath);
      if (state?.writingTimer) clearTimeout(state.writingTimer);
      this._commandStates.delete(cmdFilePath);
    });
  }

  private setupFileWatcher(document: SuperDocDocument, webview: vscode.Webview): { dispose: () => void } {
    const fileDir = vscode.Uri.joinPath(document.uri, '..');
    const fileName = path.basename(document.uri.fsPath);
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(fileDir, fileName)
    );

    let reloadTimer: ReturnType<typeof setTimeout> | null = null;

    watcher.onDidChange(async (uri) => {
      // Ignore our own saves (within 1 second)
      if (Date.now() - document.lastSaveTime < 1000) {
        debug('Ignoring file change - recent save');
        return;
      }

      // Debounce rapid external changes (e.g., editors that write temp + rename)
      if (reloadTimer) clearTimeout(reloadTimer);
      reloadTimer = setTimeout(async () => {
        reloadTimer = null;
        debug(`External file change detected: ${uri.fsPath}`);
        await document.reloadFromDisk();
        debug(`Sending reload to webview, size: ${document.data.length} bytes`);
        webview.postMessage({
          type: 'reload',
          content: { data: Array.from(document.data) },
        });
      }, 200);
    });

    return {
      dispose: () => {
        if (reloadTimer) clearTimeout(reloadTimer);
        watcher.dispose();
      }
    };
  }

  /**
   * Get the command file path for a document: .superdoc/{docname}.json
   */
  private getCommandFilePath(documentUri: vscode.Uri): string {
    const docDir = path.dirname(documentUri.fsPath);
    const docName = path.basename(documentUri.fsPath, '.docx');
    return path.join(docDir, SUPERDOC_FOLDER, `${docName}.json`);
  }

  /**
   * Setup watcher for command file (Claude API)
   * Each document watches its own file: .superdoc/{docname}.json
   *
   * fs.watch() may fire multiple events per write (known Node.js issue #7420).
   * Three guards prevent duplicate execution:
   *   1. lastProcessedId — skips same command before response is written
   *   2. writingResponse — skips events during the 200ms response write window
   *   3. !data.command  — skips events after response (file has no 'command' field)
   */
  private setupCommandWatcher(document: SuperDocDocument, webview: vscode.Webview): fs.FSWatcher | null {
    const cmdFilePath = this.getCommandFilePath(document.uri);
    const superdocDir = path.dirname(cmdFilePath);
    const cmdFileName = path.basename(cmdFilePath);
    let processing = false;

    // Initialize per-document command state
    this._commandStates.set(cmdFilePath, {
      pendingFile: null,
      writingResponse: false,
      lastProcessedId: null,
      writingTimer: null,
    });
    const state = this._commandStates.get(cmdFilePath)!;

    debug(`Setting up command watcher: ${cmdFilePath}`);

    // Ensure .superdoc folder exists
    if (!fs.existsSync(superdocDir)) {
      fs.mkdirSync(superdocDir, { recursive: true });
      debug(`Created folder: ${superdocDir}`);
    }

    const processIfExists = async () => {
      if (processing || state.writingResponse || !fs.existsSync(cmdFilePath)) return;
      processing = true;

      try {
        const content = fs.readFileSync(cmdFilePath, 'utf-8');
        const data = JSON.parse(content);

        // Only process if it's a command (has 'command' field), not a response
        if (!data.command) return;

        // Skip if this command was already processed (idempotency)
        const commandId = data.id || `${data.command}:${JSON.stringify(data.args || {})}`;
        if (commandId === state.lastProcessedId) {
          debug(`Skipping duplicate command: ${data.command} (id: ${commandId})`);
          return;
        }

        state.lastProcessedId = commandId;
        await this.processCommandFile(cmdFilePath, data, webview, document);
      } catch {
        // Ignore parse errors or missing file
      } finally {
        processing = false;
      }
    };

    // Check if command file already exists
    processIfExists();

    try {
      const fsWatcher = fs.watch(superdocDir, (eventType, filename) => {
        if (filename === cmdFileName) {
          processIfExists();
        }
      });

      return fsWatcher;
    } catch (error) {
      debug(`Failed to setup command watcher: ${error}`);
      return null;
    }
  }

  /**
   * Process a command and send to webview
   */
  private async processCommandFile(
    cmdFilePath: string,
    cmd: { command: string; args?: Record<string, unknown>; id?: string },
    webview: vscode.Webview,
    document: SuperDocDocument
  ): Promise<void> {
    debug(`Processing command: ${cmd.command}`);

    // Store the file path for response writing (per-document)
    const state = this._commandStates.get(cmdFilePath);
    if (state) state.pendingFile = cmdFilePath;

    let args = cmd.args || {};

    // Special handling for insertImage - convert URL/path to base64
    if (cmd.command === 'insertImage' && args.src) {
      try {
        args = { ...args };
        const src = args.src as string;

        if (!src.startsWith('data:')) {
          debug(`Converting image source to base64: ${src.substring(0, 100)}...`);
          const docDir = path.dirname(document.uri.fsPath);
          args.src = await this.convertImageToBase64(src, docDir);
          debug(`Image converted, base64 length: ${(args.src as string).length}`);
        }
      } catch (error) {
        // Write error response directly
        this.writeCommandResponse(cmdFilePath, {
          success: false,
          error: `Failed to load image: ${error}`
        });
        return;
      }
    }

    webview.postMessage({
      type: 'executeCommand',
      command: cmd.command,
      args,
      id: cmd.id,
    });
  }

  /**
   * Convert image URL or file path to base64 data URI
   */
  private async convertImageToBase64(src: string, docDir: string): Promise<string> {
    // Check if it's a URL
    if (src.startsWith('http://') || src.startsWith('https://')) {
      return this.fetchImageAsBase64(src);
    }

    // Otherwise treat as file path
    let filePath = src;
    if (!path.isAbsolute(src)) {
      filePath = path.join(docDir, src);
    }

    if (!fs.existsSync(filePath)) {
      throw new Error(`Image file not found: ${filePath}`);
    }

    const buffer = fs.readFileSync(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const mimeType = this.getMimeType(ext);
    return `data:${mimeType};base64,${buffer.toString('base64')}`;
  }

  /**
   * Fetch image from URL and convert to base64
   */
  private fetchImageAsBase64(url: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const protocol = url.startsWith('https://') ? https : http;

      protocol.get(url, { headers: { 'User-Agent': 'VSCode-SuperDoc/1.0' } }, (response) => {
        // Handle redirects
        if (response.statusCode === 301 || response.statusCode === 302) {
          const redirectUrl = response.headers.location;
          if (redirectUrl) {
            this.fetchImageAsBase64(redirectUrl).then(resolve).catch(reject);
            return;
          }
        }

        if (response.statusCode !== 200) {
          reject(new Error(`HTTP ${response.statusCode}: Failed to fetch image`));
          return;
        }

        const contentType = response.headers['content-type'] || 'image/png';
        const chunks: Buffer[] = [];

        response.on('data', (chunk) => chunks.push(chunk));
        response.on('end', () => {
          const buffer = Buffer.concat(chunks);
          resolve(`data:${contentType};base64,${buffer.toString('base64')}`);
        });
        response.on('error', reject);
      }).on('error', reject);
    });
  }

  /**
   * Get MIME type from file extension
   */
  private getMimeType(ext: string): string {
    const mimeTypes: Record<string, string> = {
      '.png': 'image/png',
      '.jpg': 'image/jpeg',
      '.jpeg': 'image/jpeg',
      '.gif': 'image/gif',
      '.webp': 'image/webp',
      '.svg': 'image/svg+xml',
      '.bmp': 'image/bmp',
      '.ico': 'image/x-icon'
    };
    return mimeTypes[ext] || 'image/png';
  }

  // Per-document command state (keyed by command file path)
  private _commandStates = new Map<string, {
    pendingFile: string | null;
    writingResponse: boolean;
    lastProcessedId: string | null;
    writingTimer: ReturnType<typeof setTimeout> | null;
  }>();

  /**
   * Write response by overwriting the command file
   */
  private writeCommandResponse(cmdFilePath: string, result: { success: boolean; result?: unknown; error?: string }): void {
    const state = this._commandStates.get(cmdFilePath);
    if (!state?.pendingFile) {
      debug('No pending command file to write response to');
      return;
    }

    debug(`Writing response: success=${result.success}`);
    state.writingResponse = true;
    fs.writeFileSync(state.pendingFile, JSON.stringify(result, null, 2));
    state.pendingFile = null;
    // Reset dedup so the next identical command (intentional repeat) is processed
    state.lastProcessedId = null;
    // Keep suppression active briefly to cover the debounced watcher callback
    if (state.writingTimer) clearTimeout(state.writingTimer);
    state.writingTimer = setTimeout(() => { state.writingResponse = false; state.writingTimer = null; }, 200);
  }

  private getWebviewContent(webview: vscode.Webview): string {
    const scriptUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview', 'main.js')
    );
    const styleUri = webview.asWebviewUri(
      vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview', 'style.css')
    );
    const nonce = getNonce();

    return `<!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="Content-Security-Policy" content="default-src 'none'; font-src ${webview.cspSource} https: data:; img-src ${webview.cspSource} https: data: blob:; script-src 'nonce-${nonce}' ${webview.cspSource}; style-src ${webview.cspSource} 'unsafe-inline';">
        <link href="${styleUri}" rel="stylesheet">
        <title>SuperDoc Editor</title>
      </head>
      <body>
        <div id="superdoc-toolbar"></div>
        <div id="search-bar" style="display: none;">
          <button id="search-expand" class="search-icon-btn search-chevron" title="Toggle Replace"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M6 4l4 4-4 4"/></svg></button>
          <div class="search-fields">
            <div class="search-row">
              <div class="search-input-wrap">
                <input type="text" id="search-input" placeholder="Find" />
                <button id="search-case" class="search-toggle-btn" title="Match Case">Aa</button>
              </div>
              <span id="search-count">No results</span>
              <button id="search-prev" class="search-icon-btn" title="Previous Match"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="none" stroke="currentColor" stroke-width="1.5" d="M8 12V4m0 0L4 8m4-4l4 4"/></svg></button>
              <button id="search-next" class="search-icon-btn" title="Next Match"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="none" stroke="currentColor" stroke-width="1.5" d="M8 4v8m0 0l4-4m-4 4L4 8"/></svg></button>
              <button id="search-close" class="search-icon-btn" title="Close"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M8 8.7L3.3 13.4 2.6 12.7 7.3 8 2.6 3.3 3.3 2.6 8 7.3l4.7-4.7.7.7L8.7 8l4.7 4.7-.7.7z"/></svg></button>
            </div>
            <div class="search-row search-replace-row" style="display: none;">
              <div class="search-input-wrap">
                <input type="text" id="replace-input" placeholder="Replace" />
              </div>
              <button id="replace-one" class="search-icon-btn" title="Replace"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="none" stroke="currentColor" stroke-width="1.5" d="M3 8h7m0 0L7 5m3 3L7 11"/></svg></button>
              <button id="replace-all" class="search-icon-btn" title="Replace All"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="none" stroke="currentColor" stroke-width="1.5" d="M3 5h7m0 0L7 2m3 3L7 8M3 11h7m0 0L7 8m3 3l-3 3"/></svg></button>
            </div>
          </div>
        </div>
        <div id="superdoc-scroll-wrapper">
          <div id="superdoc"></div>
        </div>
        <script type="module" nonce="${nonce}" src="${scriptUri}"></script>
      </body>
      </html>`;
  }

  private setupMessageHandler(webview: vscode.Webview, document: SuperDocDocument): void {
    const cmdFilePath = this.getCommandFilePath(document.uri);

    webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'update': {
          debug(`Received update from webview, size: ${message.content?.length}`);
          document.update(new Uint8Array(message.content));
          await document.save();
          debug(`Document saved: ${document.uri.fsPath}`);
          break;
        }
        case 'debug':
          debug(message.message);
          break;
        case 'commandResult': {
          // Overwrite command file with response
          this.writeCommandResponse(cmdFilePath, {
            success: message.success,
            result: message.result,
            error: message.error
          });
          break;
        }
      }
    });
  }

  // Required by CustomEditorProvider interface
  async saveCustomDocument(document: SuperDocDocument): Promise<void> {
    await document.save();
  }

  async saveCustomDocumentAs(document: SuperDocDocument, destination: vscode.Uri): Promise<void> {
    await vscode.workspace.fs.writeFile(destination, document.data);
  }

  async revertCustomDocument(document: SuperDocDocument): Promise<void> {
    await document.reloadFromDisk();
  }

  async backupCustomDocument(
    document: SuperDocDocument,
    context: vscode.CustomDocumentBackupContext
  ): Promise<vscode.CustomDocumentBackup> {
    await vscode.workspace.fs.writeFile(context.destination, document.data);
    return {
      id: context.destination.toString(),
      delete: async () => {
        try { await vscode.workspace.fs.delete(context.destination); } catch {}
      },
    };
  }
}

class SuperDocDocument implements vscode.CustomDocument {
  private _data: Uint8Array;
  private _lastSaveTime = 0;

  constructor(public readonly uri: vscode.Uri, initialData: Uint8Array) {
    this._data = initialData;
  }

  get data(): Uint8Array {
    return this._data;
  }

  get lastSaveTime(): number {
    return this._lastSaveTime;
  }

  update(newData: Uint8Array): void {
    this._data = newData;
  }

  async save(): Promise<void> {
    this._lastSaveTime = Date.now();
    await vscode.workspace.fs.writeFile(this.uri, this._data);
  }

  async reloadFromDisk(): Promise<void> {
    this._data = await vscode.workspace.fs.readFile(this.uri);
  }

  dispose(): void {}
}

function getNonce(): string {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  return Array.from({ length: 32 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
}
