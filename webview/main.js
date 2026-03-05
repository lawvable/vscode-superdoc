// SuperDoc integration for VS Code webview
/* global document, window, setTimeout, clearTimeout, acquireVsCodeApi, File */

import superdocCss from 'superdoc/style.css';
import { SuperDoc } from 'superdoc';

// =============================================================================
// Configuration & Setup
// =============================================================================

const DEBUG_ENABLED = false;
const AUTO_SAVE_DELAY = 1000;

// Suppress noisy SuperDoc internal logs
const originalConsoleLog = console.log;
console.log = (...args) => {
    if (typeof args[0] === 'string' && args[0].includes('[sd-table-borders]')) return;
    originalConsoleLog.apply(console, args);
};

const vscode = acquireVsCodeApi();

function debug(message) {
    if (DEBUG_ENABLED && vscode) {
        vscode.postMessage({ type: 'debug', message: `[Webview] - ${message}` });
    }
}

// Inject CSS when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectStyles);
} else {
    injectStyles();
}

function injectStyles() {
    if (superdocCss) {
        const style = document.createElement('style');
        style.textContent = superdocCss;
        document.head.appendChild(style);
    }
}

let editor = null;
let originalEditorEmit = null; // Set in setupEditorListeners; used by cmdAddComment
let saveTimeout = null;
let isInitialLoad = true;
let isExporting = false;
let modeObserver = null; // MutationObserver for document mode dropdown highlighting

// Factory for accept/reject tracked change toolbar handlers
function makeTrackedChangeHandler(label, selectionCmd, allCmd) {
    return ({ option }) => {
        if (!option) return; // Called when dropdown opens (no option selected yet)
        const activeEditor = editor?.activeEditor;
        if (!activeEditor) return;
        if (option.key === 'selection') activeEditor.commands[selectionCmd]();
        else if (option.key === 'all') activeEditor.commands[allCmd]();
        debug(`${label} changes: ${option.key}`);
    };
}

// Initialize editor with file data
function initializeEditor(fileArrayBuffer) {
    debug('Initializing editor with file buffer');

    try {
        // Convert ArrayBuffer to File object
        const file = new File([fileArrayBuffer], 'document.docx', {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        debug(`File created: ${file.name}, ${file.size} bytes`);

        // Clean up previous editor instance
        if (modeObserver) {
            modeObserver.disconnect();
            modeObserver = null;
        }
        if (editor) {
            debug('Destroying previous editor...');
            try {
                if (editor.destroy) {
                    editor.destroy();
                }
            } catch (e) {
                debug(`Error destroying previous editor: ${e.message}`);
            }
            editor = null;
            debug('Destroyed previous editor');
        }

        // Reset state for new editor
        isInitialLoad = true;

        // Check if DOM elements exist
        const superdocElement = document.getElementById('superdoc');
        const toolbarElement = document.getElementById('superdoc-toolbar');

        // Clear existing content from containers
        if (superdocElement) {
            superdocElement.innerHTML = '';
        }
        if (toolbarElement) {
            toolbarElement.innerHTML = '';
        }
        
        if (!superdocElement || !toolbarElement) {
            throw new Error('Required DOM elements not found (#superdoc or #superdoc-toolbar)');
        }
        
        debug('DOM elements found, creating SuperDoc...');
        
        try {
            editor = new SuperDoc({
                selector: '#superdoc',
                toolbar: '#superdoc-toolbar',
                document: file,
                documentMode: 'editing',  // Default to normal editing (commands switch to suggesting mode for tracked changes)
                role: 'editor',
                permissionResolver: ({ defaultDecision, permission }) => {
                    // Allow all tracked change accept/reject actions regardless of author
                    if (permission === 'RESOLVE_OWN' || permission === 'RESOLVE_OTHER' ||
                        permission === 'REJECT_OWN' || permission === 'REJECT_OTHER') {
                        return true;
                    }
                    return defaultDecision;
                },
                pagination: true,
                rulers: true,
                user: {
                    name: 'Claude',
                    email: 'claude@anthropic.com'
                },
                modules: {
                    toolbar: {
                        excludeItems: ['acceptTrackedChangeBySelection', 'rejectTrackedChangeOnSelection'],
                        customButtons: [
                            {
                                type: 'button',
                                name: 'findReplace',
                                tooltip: 'Find & Replace',
                                icon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512"><path d="M416 208c0 45.9-14.9 88.3-40 122.7L502.6 457.4c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L330.7 376c-34.4 25.2-76.8 40-122.7 40C93.1 416 0 322.9 0 208S93.1 0 208 0S416 93.1 416 208zM208 352a144 144 0 1 0 0-288 144 144 0 1 0 0 288z"/></svg>',
                                group: 'right',
                                command: () => {
                                    openSearchBar();
                                }
                            },
                            {
                                type: 'dropdown',
                                name: 'acceptChanges',
                                tooltip: 'Accept Changes',
                                icon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512"><path d="M128 0c17.7 0 32 14.3 32 32l0 32 128 0 0-32c0-17.7 14.3-32 32-32s32 14.3 32 32l0 32 48 0c26.5 0 48 21.5 48 48l0 48L0 160l0-48C0 85.5 21.5 64 48 64l48 0 0-32c0-17.7 14.3-32 32-32zM0 192l448 0 0 272c0 26.5-21.5 48-48 48L48 512c-26.5 0-48-21.5-48-48L0 192zM329 305c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-95 95-47-47c-9.4-9.4-24.6-9.4-33.9 0s-9.4 24.6 0 33.9l64 64c9.4 9.4 24.6 9.4 33.9 0L329 305z"/></svg>',
                                hasCaret: true,
                                group: 'left',
                                options: [
                                    { label: 'Accept under selection', key: 'selection' },
                                    { label: 'Accept all changes', key: 'all' }
                                ],
                                command: makeTrackedChangeHandler('Accept', 'acceptTrackedChangeBySelection', 'acceptAllTrackedChanges')
                            },
                            {
                                type: 'dropdown',
                                name: 'rejectChanges',
                                tooltip: 'Reject Changes',
                                icon: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512"><path d="M128 0c17.7 0 32 14.3 32 32l0 32 128 0 0-32c0-17.7 14.3-32 32-32s32 14.3 32 32l0 32 48 0c26.5 0 48 21.5 48 48l0 48L0 160l0-48C0 85.5 21.5 64 48 64l48 0 0-32c0-17.7 14.3-32 32-32zM0 192l448 0 0 272c0 26.5-21.5 48-48 48L48 512c-26.5 0-48-21.5-48-48L0 192zM305 305c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-47 47-47-47c-9.4-9.4-24.6-9.4-33.9 0s-9.4 24.6 0 33.9l47 47-47 47c-9.4 9.4-9.4 24.6 0 33.9s24.6 9.4 33.9 0l47-47 47 47c9.4 9.4 24.6 9.4 33.9 0s9.4-24.6 0-33.9l-47-47 47-47z"/></svg>',
                                hasCaret: true,
                                group: 'left',
                                options: [
                                    { label: 'Reject under selection', key: 'selection' },
                                    { label: 'Reject all changes', key: 'all' }
                                ],
                                command: makeTrackedChangeHandler('Reject', 'rejectTrackedChangeOnSelection', 'rejectAllTrackedChanges')
                            }
                        ]
                    }
                },
                onReady: () => {
                    debug('SuperDoc is ready (editing mode)');
                    isInitialLoad = false;
                    setupEditorListeners();
                },
                onEditorCreate: () => {
                    debug('Editor created');
                },
                onError: (error) => {
                    debug(`SuperDoc error: ${error.message || error}`);
                }
            });
            
            debug('SuperDoc init complete');
            
        } catch (constructorError) {
            debug(`SuperDoc constructor failed: ${constructorError.message}`);
        }

    } catch (error) {
        debug(`Failed to initialize SuperDoc: ${error.message}`);
    }
}

// Setup editor listeners for content changes (debounced save on update)
function setupEditorListeners() {
    if (!editor?.activeEditor) {
        debug('No editor or activeEditor available');
        return;
    }

    debug('Setting up editor update listener');

    // Intercept commentsUpdate 'selected' events to prevent an infinite loop.
    //
    // SuperDoc's comments plugin emits 'selected' events when the cursor is
    // inside a comment/tracked-change range. SuperDoc's internal handler
    // (onEditorCommentsUpdate) calls setActiveComment() in response, which
    // dispatches a new transaction, which triggers apply() again, which emits
    // another 'selected' event — creating an infinite oscillation between
    // selected(commentId) and selected(null) at ~1Hz.
    //
    // Fix: suppress 'selected' events at the emit level so they never reach
    // the handler. Other event types (add, trackedChange, etc.) pass through.
    originalEditorEmit = editor.activeEditor.emit.bind(editor.activeEditor);
    editor.activeEditor.emit = function(event, ...args) {
        if (event === 'commentsUpdate' && args[0]?.type === 'selected') {
            return;
        }
        return originalEditorEmit(event, ...args);
    };

    editor.activeEditor.on('update', () => {
        if (isInitialLoad || isExporting) return;
        scheduleAutoSave();
        // Refresh search results if search bar is open (handles undo/redo)
        refreshSearch();
    });

    debug('Editor update listener ready');

    // Highlight the active document mode option when the dropdown opens
    // NPopover teleports dropdown content to document.body, so we must observe body
    modeObserver = new MutationObserver(() => {
        const options = document.querySelectorAll('[data-item="btn-documentMode-option"]');
        if (options.length === 0) return;
        const currentMode = editor?.config?.documentMode || 'editing';
        const modeLabels = { editing: 'Editing', suggesting: 'Suggesting', viewing: 'Viewing' };
        const activeLabel = modeLabels[currentMode];
        options.forEach(opt => {
            const label = opt.querySelector('.document-mode-type');
            opt.classList.toggle('document-mode-active', label?.textContent?.trim() === activeLabel);
        });
    });
    modeObserver.observe(document.body, { childList: true, subtree: true });
}

// Schedule auto-save with debouncing
function scheduleAutoSave() {
    // Clear existing timeout
    if (saveTimeout) {
        clearTimeout(saveTimeout);
    }

    // Schedule new save
    saveTimeout = setTimeout(() => {
        saveDocument();
    }, AUTO_SAVE_DELAY);
}

// Save document back to VS Code
async function saveDocument() {
    if (!editor) {
        debug('No editor available for saving');
        return;
    }

    if (isExporting) return;
    isExporting = true;

    try {
        debug('Starting document save...');

        const blob = await editor.export({ format: 'docx', triggerDownload: false });
        if (!blob) {
            debug('Failed to export - no blob returned');
            return;
        }

        debug(`Exported blob size: ${blob.size} bytes`);

        const arrayBuffer = await blob.arrayBuffer();
        const contentArray = Array.from(new Uint8Array(arrayBuffer));

        vscode.postMessage({
            type: 'update',
            content: contentArray
        });

        debug(`Document sent to VS Code (${contentArray.length} bytes)`);
    } catch (error) {
        debug(`Error saving document: ${error.message}`);
    } finally {
        // Delay reset to catch any post-export microtask/reactivity updates
        setTimeout(() => { isExporting = false; }, 100);
    }
}

// Handle messages from VS Code
let lastExecutedCommandId = null;
let commandExecuting = false;

window.addEventListener('message', async event => {
    const message = event.data;
    if (!message?.type) {
        return;
    }

    debug(`Received message: ${message.type}`);

    switch (message.type) {
        case 'update':
        case 'reload':
            if (message.content?.data) {
                const fileBuffer = new Uint8Array(message.content.data).buffer;
                debug(`${message.type}: ${fileBuffer.byteLength} bytes`);
                initializeEditor(fileBuffer);
            }
            break;

        case 'executeCommand': {
            // Prevent duplicate/concurrent command execution
            const cmdId = message.id || `${message.command}:${JSON.stringify(message.args || {})}`;
            if (cmdId === lastExecutedCommandId || commandExecuting) {
                debug(`Skipping duplicate command: ${message.command}`);
                break;
            }
            commandExecuting = true;
            lastExecutedCommandId = cmdId;
            try {
                debug(`Executing command: ${message.command}`);
                const result = await executeCommand(message.command, message.args || {});
                vscode.postMessage({ type: 'commandResult', ...result });
            } finally {
                commandExecuting = false;
                // Reset dedup so the next identical command (intentional repeat) is processed
                lastExecutedCommandId = null;
            }
            break;
        }
    }
});

// =============================================================================
// Command Execution API
// =============================================================================

const COMMANDS = {
    getText: cmdGetText,
    getNodes: cmdGetNodes,
    replaceText: cmdReplaceText,
    insertContent: cmdInsertContent,
    formatText: cmdFormatText,
    insertImage: cmdInsertImage,
    deleteNode: cmdDeleteNode,
    insertTable: cmdInsertTable,
    addComment: cmdAddComment,
    insertTableOfContents: cmdInsertTableOfContents,
    deleteTableOfContents: cmdDeleteTableOfContents,
    undo: cmdUndo,
    redo: cmdRedo,
    acceptAllChanges: cmdAcceptAllChanges,
    rejectAllChanges: cmdRejectAllChanges,
    focusHeader: cmdFocusHeader,
    focusFooter: cmdFocusFooter,
    exitHeaderFooter: cmdExitHeaderFooter
};

async function executeCommand(command, args) {
    const handler = COMMANDS[command];
    if (!handler) {
        return { success: false, error: `Unknown command: ${command}` };
    }
    try {
        return await handler(args);
    } catch (error) {
        debug(`Command error: ${error.message}`);
        return { success: false, error: error.message };
    }
}

// =============================================================================
// Editor Helpers
// =============================================================================

function getActiveEditor() {
    // First check if there's a presentationEditor with getActiveEditor method
    // This handles header/footer mode correctly
    const baseEditor = editor?.activeEditor;
    if (!baseEditor) return null;

    const presentationEditor = baseEditor.presentationEditor || baseEditor._presentationEditor;
    if (presentationEditor && typeof presentationEditor.getActiveEditor === 'function') {
        return presentationEditor.getActiveEditor();
    }

    return baseEditor;
}

function requireActiveEditor() {
    const activeEditor = getActiveEditor();
    if (!activeEditor) {
        return { error: { success: false, error: 'No active editor' } };
    }
    return { activeEditor };
}

function getFormattedText(activeEditor, { resolveTrackedChanges = false } = {}) {
    if (resolveTrackedChanges) {
        return getTextWithTrackedChangesResolved(activeEditor.state.doc);
    }
    try {
        return activeEditor.getText({ blockSeparator: '\n\n' });
    } catch {
        return activeEditor.state.doc.textContent;
    }
}

/**
 * Walk the document tree and return plain text as if all tracked changes
 * were accepted: text with `trackDelete` marks is excluded, text with
 * `trackInsert` marks is kept (mark stripped conceptually).
 */
function getTextWithTrackedChangesResolved(doc) {
    const parts = [];
    let lastBlockPos = -1;

    doc.descendants((node, pos, parent) => {
        // Add block separator between top-level blocks
        if (node.isBlock && node.isTextblock) {
            if (lastBlockPos !== -1) {
                parts.push('\n\n');
            }
            lastBlockPos = pos;
        }

        if (node.isText) {
            const isDeleted = node.marks.some(m => m.type.name === 'trackDelete');
            if (!isDeleted) {
                parts.push(node.text);
            }
        }
    });

    return parts.join('');
}

function setDocumentMode(mode) {
    editor.setDocumentMode(mode);
    if (mode === 'suggesting') {
        getActiveEditor()?.commands.enableTrackChanges?.();
    }
}

function setAuthorIfProvided(author) {
    if (author?.name) {
        const user = {
            name: author.name,
            email: author.email || `${author.name.toLowerCase().replace(/\s+/g, '.')}@user.local`
        };
        // Set on both SuperDoc wrapper and Editor instance (track changes reads from editor.options.user)
        editor.user = user;
        const activeEditor = getActiveEditor();
        if (activeEditor) {
            activeEditor.options.user = user;
        }
        debug(`Author set to: ${user.name}`);
    }
}

function searchText(activeEditor, text) {
    return activeEditor.commands.search(text, { highlight: false }) || [];
}

function findAnchor(activeEditor, anchor) {
    const matches = searchText(activeEditor, anchor);
    return matches[0] || null;
}

function findMatch(activeEditor, search, occurrence) {
    const matches = searchText(activeEditor, search);
    if (matches.length === 0) {
        return { error: `Text not found: "${search}"` };
    }
    const idx = occurrence ? parseInt(occurrence, 10) - 1 : 0;
    if (idx < 0 || idx >= matches.length) {
        return { error: `Occurrence ${occurrence} not found (only ${matches.length} matches)` };
    }
    return { match: matches[idx], matches };
}

function insertAtPosition(activeEditor, position, content) {
    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from: position, to: position });
    activeEditor.commands.insertContent(content);
}

function selectRange(activeEditor, from, to, { focus = true } = {}) {
    if (focus) activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from, to: to ?? from });
}

function applyScope(activeEditor, scope) {
    activeEditor.view.focus();
    if (scope === 'document') {
        activeEditor.commands.selectAll();
    } else if (scope?.from !== undefined && scope?.to !== undefined) {
        activeEditor.commands.setTextSelection({ from: scope.from, to: scope.to });
    }
}

function cmdGetText({ format, resolveTrackedChanges, head, tail } = {}) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const validFormats = ['text', 'html', 'both'];
    const selectedFormat = format || 'both';

    if (!validFormats.includes(selectedFormat)) {
        return { success: false, error: `Invalid format: "${format}". Valid formats: ${validFormats.join(', ')}` };
    }

    if (head != null && tail != null) {
        return { success: false, error: 'Cannot use both "head" and "tail" at the same time' };
    }

    const resolve = resolveTrackedChanges === true;
    const result = {};

    if (selectedFormat === 'text' || selectedFormat === 'both') {
        let text = getFormattedText(activeEditor, { resolveTrackedChanges: resolve });
        text = truncateText(text, { head, tail });
        result.text = text;
    }

    if (selectedFormat === 'html' || selectedFormat === 'both') {
        result.html = activeEditor.getHTML?.() || null;
    }

    const charCount = result.text?.length || result.html?.length || 0;
    debug(`getText(${selectedFormat}, resolve=${resolve}, head=${head ?? '-'}, tail=${tail ?? '-'}): ${charCount} chars`);
    return { success: true, result };
}

function truncateText(text, { head, tail }) {
    const h = head != null ? Math.floor(Number(head)) : null;
    const t = tail != null ? Math.floor(Number(tail)) : null;

    if (h != null && h > 0 && text.length > h) {
        return text.slice(0, h) + `\n\n[... truncated, ${text.length - h} more chars]`;
    }
    if (t != null && t > 0 && text.length > t) {
        return `[... truncated, ${text.length - t} chars skipped]\n\n` + text.slice(-t);
    }
    return text;
}

function cmdGetNodes({ type }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!type) return { success: false, error: 'Node type is required' };

    const validTypes = ['paragraph', 'table', 'tableRow', 'tableCell',
                        'bulletList', 'orderedList', 'listItem', 'image', 'blockquote'];

    if (!validTypes.includes(type)) {
        return { success: false, error: `Invalid type: "${type}". Valid types: ${validTypes.join(', ')}` };
    }

    const nodes = activeEditor.getNodesOfType(type);

    const result = nodes.map((item, index) => {
        const { node, pos } = item;
        const from = pos;
        const to = pos + node.nodeSize;
        const text = node.textContent || '';

        const entry = {
            index,
            type,
            from,
            to,
            text: text.substring(0, 100) + (text.length > 100 ? '...' : ''),
            textLength: text.length,
        };

        // For paragraphs, include numbering marker if present (e.g., "1.", "(A)", "•")
        if (type === 'paragraph') {
            try {
                const dom = activeEditor.view.nodeDOM(pos);
                const marker = dom?.getAttribute?.('data-marker-type');
                if (marker) entry.marker = marker;
            } catch {}
        }

        return entry;
    });

    debug(`getNodes: found ${result.length} ${type} nodes`);
    return { success: true, result: { nodes: result, count: result.length } };
}

async function cmdFormatText({ fontFamily, fontSize, color, highlight, bold, italic, underline, strikethrough, link, lineHeight, indent, spacingBefore, spacingAfter, textAlign, scope }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const hasFormat = fontFamily || fontSize || color || highlight !== undefined ||
                      bold !== undefined || italic !== undefined ||
                      underline !== undefined || strikethrough !== undefined ||
                      link !== undefined ||
                      lineHeight !== undefined || indent !== undefined ||
                      spacingBefore !== undefined || spacingAfter !== undefined ||
                      textAlign !== undefined;
    if (!hasFormat) {
        return { success: false, error: 'At least one format option required' };
    }

    const previousMode = editor.config.documentMode || 'editing';
    editor.setDocumentMode('editing');
    applyScope(activeEditor, scope);

    const applied = [];

    // Font properties
    if (fontFamily) {
        activeEditor.commands.setFontFamily(fontFamily);
        applied.push(`fontFamily: ${fontFamily}`);
    }

    if (fontSize) {
        activeEditor.commands.setFontSize(fontSize);
        applied.push(`fontSize: ${fontSize}`);
    }

    if (color) {
        activeEditor.commands.setColor(color);
        applied.push(`color: ${color}`);
    }

    // Highlight (background color) - string = set color, false = remove
    if (highlight && highlight !== false) {
        activeEditor.commands.setHighlight(highlight);
        applied.push(`highlight: ${highlight}`);
    } else if (highlight === false) {
        activeEditor.commands.unsetHighlight();
        applied.push('highlight: removed');
    }

    // Text formatting (true = set, false = unset)
    if (bold === true) {
        activeEditor.commands.setBold();
        applied.push('bold: true');
    } else if (bold === false) {
        activeEditor.commands.unsetBold();
        applied.push('bold: false');
    }

    if (italic === true) {
        activeEditor.commands.setItalic();
        applied.push('italic: true');
    } else if (italic === false) {
        activeEditor.commands.unsetItalic();
        applied.push('italic: false');
    }

    if (underline === true) {
        activeEditor.commands.setUnderline();
        applied.push('underline: true');
    } else if (underline === false) {
        activeEditor.commands.unsetUnderline();
        applied.push('underline: false');
    }

    if (strikethrough === true) {
        activeEditor.commands.setStrike();
        applied.push('strikethrough: true');
    } else if (strikethrough === false) {
        activeEditor.commands.unsetStrike();
        applied.push('strikethrough: false');
    }

    // Link - string = set link href, false = remove link
    if (link && link !== false) {
        activeEditor.commands.setLink({ href: link });
        applied.push(`link: ${link}`);
    } else if (link === false) {
        activeEditor.commands.unsetLink();
        applied.push('link: removed');
    }

    // Indentation (in points) — uses setTextIndentation which sets paragraphProperties.indent.left
    if (indent !== undefined) {
        const indentPoints = parseFloat(indent);
        if (!Number.isNaN(indentPoints)) {
            if (indentPoints === 0) {
                activeEditor.commands.unsetTextIndentation();
            } else {
                activeEditor.commands.setTextIndentation(indentPoints);
            }
            applied.push(`indent: ${indent}`);
        }
    }

    // Line height — uses setLineHeight which sets paragraphProperties.spacing.line
    if (lineHeight !== undefined) {
        const lh = parseFloat(lineHeight);
        if (!Number.isNaN(lh)) {
            if (lh === 0) {
                activeEditor.commands.unsetLineHeight();
            } else {
                activeEditor.commands.setLineHeight(lh);
            }
            applied.push(`lineHeight: ${lineHeight}`);
        }
    }

    // Spacing before/after — set via paragraphProperties.spacing (values in twips, 1pt = 20 twips)
    if (spacingBefore !== undefined) {
        const pts = parseFloat(spacingBefore);
        if (!Number.isNaN(pts)) {
            activeEditor.commands.updateAttributes('paragraph', {
                'paragraphProperties.spacing.before': pts * 20,
            });
            applied.push(`spacingBefore: ${spacingBefore}`);
        }
    }
    if (spacingAfter !== undefined) {
        const pts = parseFloat(spacingAfter);
        if (!Number.isNaN(pts)) {
            activeEditor.commands.updateAttributes('paragraph', {
                'paragraphProperties.spacing.after': pts * 20,
            });
            applied.push(`spacingAfter: ${spacingAfter}`);
        }
    }

    // Text alignment — left, center, right, justify
    if (textAlign !== undefined) {
        const validAligns = ['left', 'center', 'right', 'justify'];
        if (validAligns.includes(textAlign)) {
            activeEditor.commands.setTextAlign(textAlign);
            applied.push(`textAlign: ${textAlign}`);
        }
    }

    // Restore previous mode
    editor.setDocumentMode(previousMode);

    await saveDocument();
    debug(`formatText: applied ${applied.join(', ')}`);
    return { success: true, result: { applied } };
}

async function cmdReplaceText({ search, replacement, occurrence, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!search) return { success: false, error: 'Search text is required' };

    // Search for matches first (before changing document mode)
    const matches = searchText(activeEditor, search);
    if (matches.length === 0) {
        return { success: false, error: `Text not found: "${search}"` };
    }

    // Validate occurrence if specified
    let toReplace;
    if (occurrence != null) {
        const idx = parseInt(occurrence, 10) - 1;
        if (idx < 0 || idx >= matches.length) {
            return { success: false, error: `Occurrence ${occurrence} not found (only ${matches.length} matches)` };
        }
        toReplace = [matches[idx]];
    } else {
        // Reverse to maintain positions when replacing multiple
        toReplace = [...matches].reverse();
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    debug(`replaceText: found ${matches.length} matches`);

    // Replace using proper positions from search
    // Note: Each replacement is a separate undo step (TipTap limitation with track changes)
    for (const m of toReplace) {
        selectRange(activeEditor, m.from, m.to);
        activeEditor.commands.insertContent(replacement);
    }

    await saveDocument();
    setDocumentMode('editing');
    debug(`replaceText: replaced ${toReplace.length} occurrence(s)`);
    return { success: true, result: { replacedCount: toReplace.length } };
}

async function cmdInsertContent({ content, position, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!content) return { success: false, error: 'Content is required' };

    const anchor = position?.after || position?.before;
    const insertAfter = Boolean(position?.after);
    const textContent = activeEditor.state.doc.textContent.trim();
    const isEmptyDoc = textContent.length === 0;

    // If no anchor provided and document is empty (or position is explicitly "start"/"end"), insert at position 1
    const insertAtStart = !anchor || position?.before === '' || position?.after === '';

    // Validate before changing mode
    if (!insertAtStart && !isEmptyDoc && !anchor) {
        return { success: false, error: 'Position anchor required: use "after" or "before" with existing text, or use empty string to insert at start' };
    }

    // Find anchor before changing mode (if needed)
    let insertPos = 1;
    if (!insertAtStart && !isEmptyDoc && anchor) {
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = insertAfter ? match.to : match.from;
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    insertAtPosition(activeEditor, insertPos, content);
    debug(isEmptyDoc || insertAtStart ? 'insertContent: at start of document' : `insertContent: ${insertAfter ? 'after' : 'before'} "${anchor}"`);

    await saveDocument();
    setDocumentMode('editing');
    return { success: true };
}

async function cmdInsertImage({ src, alt, width, position }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!src) return { success: false, error: 'Image src is required (URL or base64 data URI)' };

    const anchor = position?.after || position?.before;
    if (!anchor) {
        return { success: false, error: 'Position anchor (after/before) is required' };
    }

    // Use edit mode for images (not track changes)
    setDocumentMode('editing');

    const match = findAnchor(activeEditor, anchor);
    if (!match) {
        return { success: false, error: `Anchor text not found: "${anchor}"` };
    }

    const imageAttrs = { src };
    if (alt) imageAttrs.alt = alt;
    if (width) imageAttrs.size = { width };

    try {
        const insertPos = position?.after ? match.to : match.from;
        insertAtPosition(activeEditor, insertPos, { type: 'image', attrs: imageAttrs });
    } catch (e) {
        return { success: false, error: `Failed to insert image: ${e.message}` };
    }

    await saveDocument();
    return { success: true };
}

async function cmdDeleteNode({ type, index }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!type) return { success: false, error: 'Node type is required' };
    if (index === undefined) return { success: false, error: 'Node index is required' };

    const nodes = activeEditor.getNodesOfType(type);
    if (!nodes || nodes.length === 0) {
        return { success: false, error: `No ${type} nodes found in document` };
    }

    const idx = parseInt(index, 10);
    if (idx < 0 || idx >= nodes.length) {
        return { success: false, error: `Index ${index} out of range (${nodes.length} ${type} nodes found)` };
    }

    const { pos, node } = nodes[idx];

    setDocumentMode('suggesting');
    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from: pos, to: pos + node.nodeSize });
    activeEditor.commands.deleteSelection();

    await saveDocument();
    setDocumentMode('editing');
    debug(`deleteNode: deleted ${type} at index ${index}`);
    return { success: true, result: { deletedType: type, deletedIndex: index } };
}

async function cmdInsertTable({ rows, cols, data, position, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    // Infer dimensions from data if provided
    const tableRows = rows || (data ? data.length : 2);
    const tableCols = cols || (data && data[0] ? data[0].length : 2);

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    // Position the cursor if anchor provided
    let insertPos = null;
    if (position?.after || position?.before) {
        const anchor = position.after || position.before;
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = position.after ? match.to : match.from;
        activeEditor.view.focus();
        activeEditor.commands.setTextSelection({ from: insertPos, to: insertPos });
    }

    const result = activeEditor.commands.insertTable({ rows: tableRows, cols: tableCols });
    if (!result) {
        return { success: false, error: 'Failed to insert table' };
    }

    // Populate cells with data if provided
    if (data && Array.isArray(data)) {
        // Find the newly created table
        const tablesAfter = activeEditor.getNodesOfType('table');
        const newTable = tablesAfter.find(t => t.pos >= (insertPos || 0) - 5)
            || tablesAfter[tablesAfter.length - 1];

        if (newTable) {
            // Get cells only from this specific table
            const tableEnd = newTable.pos + newTable.node.nodeSize;
            const tableCells = activeEditor.getNodesOfType('tableCell')
                .filter(cell => cell.pos >= newTable.pos && cell.pos < tableEnd);

            // Build list of {cellIndex, content} pairs to insert
            const insertions = [];
            let cellIndex = 0;
            for (let row = 0; row < data.length && row < tableRows; row++) {
                const rowData = data[row];
                if (!Array.isArray(rowData)) {
                    cellIndex += tableCols;
                    continue;
                }
                for (let col = 0; col < tableCols; col++) {
                    if (col < rowData.length && rowData[col]) {
                        insertions.push({ cellIndex, content: rowData[col] });
                    }
                    cellIndex++;
                }
            }

            // Insert in REVERSE order so positions don't shift for unprocessed cells
            activeEditor.view.focus();
            for (let i = insertions.length - 1; i >= 0; i--) {
                const { cellIndex: idx, content } = insertions[i];
                if (tableCells[idx]) {
                    const cellInsertPos = tableCells[idx].pos + 2;
                    activeEditor.commands.setTextSelection({ from: cellInsertPos, to: cellInsertPos });
                    activeEditor.commands.insertContent(content);
                }
            }
        }
    }

    await saveDocument();
    setDocumentMode('editing');
    debug(`insertTable: ${tableRows}x${tableCols} table created${data ? ' with data' : ''}`);
    return { success: true, result: { rows: tableRows, cols: tableCols } };
}

async function cmdUndo() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const result = activeEditor.commands.undo();
    if (result) {
        await saveDocument();
        debug('undo: success');
    }
    return { success: result };
}

async function cmdRedo() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const result = activeEditor.commands.redo();
    if (result) {
        await saveDocument();
        debug('redo: success');
    }
    return { success: result };
}

async function cmdAcceptAllChanges() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    activeEditor.commands.acceptAllTrackedChanges();
    await saveDocument();
    debug('acceptAllChanges: success');
    return { success: true };
}

async function cmdRejectAllChanges() {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    activeEditor.commands.rejectAllTrackedChanges();
    await saveDocument();
    debug('rejectAllChanges: success');
    return { success: true };
}

// =============================================================================
// Header/Footer Commands
// =============================================================================

function dispatchKeyboardShortcut(key, code, { ctrl = false, alt = false } = {}) {
    const activeEditor = getActiveEditor();
    const editorDom = activeEditor?.view?.dom;
    if (!editorDom) return { error: 'Editor DOM not available' };

    const container = editorDom.closest('.presentation-editor') || document.querySelector('.presentation-editor');
    if (!container) return { error: 'Presentation editor container not found' };

    container.dispatchEvent(new KeyboardEvent('keydown', {
        key, code, ctrlKey: ctrl, altKey: alt, shiftKey: false, bubbles: true, cancelable: true
    }));
    return { success: true };
}

async function cmdFocusHeader() {
    const result = dispatchKeyboardShortcut('h', 'KeyH', { ctrl: true, alt: true });
    if (result.error) return { success: false, error: result.error };
    await new Promise(resolve => setTimeout(resolve, 150));
    debug('focusHeader: entered header mode');
    return { success: true, result: { mode: 'header' } };
}

async function cmdFocusFooter() {
    const result = dispatchKeyboardShortcut('f', 'KeyF', { ctrl: true, alt: true });
    if (result.error) return { success: false, error: result.error };
    await new Promise(resolve => setTimeout(resolve, 150));
    debug('focusFooter: entered footer mode');
    return { success: true, result: { mode: 'footer' } };
}

async function cmdExitHeaderFooter() {
    const result = dispatchKeyboardShortcut('Escape', 'Escape');
    if (result.error) return { success: false, error: result.error };
    await new Promise(resolve => setTimeout(resolve, 100));
    await saveDocument();
    debug('exitHeaderFooter: exited to body mode');
    return { success: true, result: { mode: 'body' } };
}

/**
 * Insert a table of contents with bookmarks and internal links.
 *
 * The LLM identifies heading entries (positions + levels) and passes them in.
 * This command:
 *   1. Inserts bookmarkStart/bookmarkEnd pairs at each heading (no style change)
 *   2. Builds a proper tableOfContents node with link marks pointing to those bookmarks
 *   3. Inserts the TOC at the specified position
 *
 * @param {Array} entries - [{level: 1-6, from: N, to: M}, ...] heading positions
 * @param {Object} [position] - {after: "text"} or {before: "text"} or omit for beginning
 * @param {string} [title] - TOC title (default: "Table of Contents", "" for none)
 * @param {Object} [style] - {fontFamily, fontSize, color} for TOC styling. Bold + black by default.
 * @param {Object} [author] - {name, email} for track changes attribution
 */
async function cmdInsertTableOfContents({ entries, position, title, style, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    if (!entries || !Array.isArray(entries) || entries.length === 0) {
        return { success: false, error: 'entries is required: [{level: 1-6, from: N, to: M}, ...]' };
    }

    // Validate and read text for each entry from the document
    const doc = activeEditor.state.doc;
    const schema = activeEditor.state.schema;
    const resolvedEntries = [];
    const timestamp = Date.now();
    for (let i = 0; i < entries.length; i++) {
        const { level, from, to } = entries[i];
        if (!level || !Number.isInteger(from) || !Number.isInteger(to)) {
            return { success: false, error: `Entry ${i}: level, from, and to are required integers` };
        }
        if (level < 1 || level > 6) {
            return { success: false, error: `Entry ${i}: level must be 1-6, got ${level}` };
        }
        if (from >= to || from < 0 || to > doc.content.size) {
            return { success: false, error: `Entry ${i}: invalid range from=${from} to=${to} (doc size: ${doc.content.size})` };
        }
        let text = doc.textBetween(from, to, ' ').trim();
        if (!text) {
            return { success: false, error: `Entry ${i}: no text found at range ${from}-${to}` };
        }
        // Prepend numbering marker if the paragraph has one (e.g., "1.", "(A)")
        try {
            const dom = activeEditor.view.nodeDOM(from);
            const marker = dom?.getAttribute?.('data-marker-type');
            if (marker) text = `${marker} ${text}`;
        } catch {}
        const bookmarkName = `_Toc_${i}_${timestamp}`;
        resolvedEntries.push({ level, from, to, text, bookmarkName });
    }

    // Resolve TOC insertion position before modifying the document
    const anchor = position?.after || position?.before;
    let insertPos = 1;
    if (anchor) {
        const match = findAnchor(activeEditor, anchor);
        if (!match) {
            return { success: false, error: `Anchor text not found: "${anchor}"` };
        }
        insertPos = position?.after ? match.to : match.from;
    }

    setAuthorIfProvided(author);
    setDocumentMode('suggesting');

    // Build everything in a single transaction for atomic undo
    const tr = activeEditor.state.tr;

    // Step 1: Insert bookmarks inside each heading paragraph (work backwards to preserve positions)
    // Positions from getNodes: from = before <p>, to = after </p>
    // Insert inside: from+1 = start of text content, to-1 = end of text content
    const sortedForBookmarks = [...resolvedEntries].sort((a, b) => b.from - a.from);
    for (const entry of sortedForBookmarks) {
        const bmEnd = schema.nodes.bookmarkEnd.create({ id: entry.bookmarkName });
        tr.insert(entry.to - 1, bmEnd);
        const bmStart = schema.nodes.bookmarkStart.create({ name: entry.bookmarkName, id: entry.bookmarkName });
        tr.insert(entry.from + 1, bmStart);
    }

    // Step 2: Build TOC node with plain text + link marks (styling applied after insertion via commands)
    const tocParagraphs = [];
    const tocTitle = title !== undefined ? title : 'Table of Contents';
    const linkMarkType = schema.marks.link;

    if (tocTitle) {
        tocParagraphs.push(
            schema.nodes.paragraph.create({}, [schema.text(tocTitle)])
        );
    }

    const entryIndentLevels = [];
    for (const entry of resolvedEntries) {
        entryIndentLevels.push(entry.level);
        const marks = [
            linkMarkType.create({ href: null, anchor: entry.bookmarkName, name: entry.bookmarkName }),
        ];
        tocParagraphs.push(
            schema.nodes.paragraph.create({}, [schema.text(entry.text, marks)])
        );
    }
    const tocNode = schema.nodes.tableOfContents.create(
        { instruction: `TOC \\o "1-${Math.max(...resolvedEntries.map(e => e.level))}"` },
        tocParagraphs
    );

    // Step 3: Insert TOC node (position mapped through bookmark insertions)
    const mappedInsertPos = tr.mapping.map(insertPos);
    tr.insert(mappedInsertPos, tocNode);
    activeEditor.view.dispatch(tr);

    // Step 4: Apply styling and indentation via SuperDoc commands
    // Switch to editing mode — formatting must not create track changes
    setDocumentMode('editing');
    // Helper: re-fetch TOC and compute a child paragraph's position (each command changes the doc)
    const getTocChild = (childIdx) => {
        const nodes = activeEditor.getNodesOfType('tableOfContents');
        if (nodes.length === 0) return null;
        const t = nodes[0];
        let pos = t.pos + 1;
        for (let c = 0; c < childIdx; c++) pos += t.node.child(c).nodeSize;
        return { pos, node: t.node.child(childIdx), totalChildren: t.node.childCount };
    };

    const firstChild = getTocChild(0);
    if (firstChild) {
        // Apply base styling to each paragraph individually
        // (TextSelection cannot span the TOC node, only its inline-content children)
        for (let c = 0; c < firstChild.totalChildren; c++) {
            const p = getTocChild(c);
            if (!p) break;
            selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
            if (style?.bold !== false) activeEditor.commands.setBold();
            activeEditor.commands.setColor(style?.color || '#000000');
            if (style?.fontFamily) activeEditor.commands.setFontFamily(style.fontFamily);
            if (style?.fontSize) activeEditor.commands.setFontSize(style.fontSize);
        }

        // Apply larger font size to title
        if (tocTitle && style?.fontSize) {
            const sizeMatch = style.fontSize.match(/^(\d+(?:\.\d+)?)(pt|px)$/);
            if (sizeMatch) {
                const titleSize = `${parseFloat(sizeMatch[1]) + 2}${sizeMatch[2]}`;
                const p = getTocChild(0);
                if (p) {
                    selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
                    activeEditor.commands.setFontSize(titleSize);
                }
            }
        }

        // Apply indentation per entry (level 1 = 0.5", level 2 = 1.0", etc.)
        const titleOffset = tocTitle ? 1 : 0;
        for (let i = 0; i < entryIndentLevels.length; i++) {
            const p = getTocChild(i + titleOffset);
            if (p) {
                selectRange(activeEditor, p.pos + 1, p.pos + p.node.nodeSize - 1);
                activeEditor.commands.setTextIndentation(entryIndentLevels[i] * 36);
            }
        }
    }

    await saveDocument();
    setDocumentMode('editing');

    debug(`insertTableOfContents: inserted TOC with ${resolvedEntries.length} entries and bookmarks`);
    return {
        success: true,
        result: {
            entriesCount: resolvedEntries.length,
            entries: resolvedEntries.map(e => ({ level: e.level, text: e.text, bookmark: e.bookmarkName }))
        }
    };
}

/**
 * Delete the table of contents node from the document.
 * Optionally also removes the bookmarks that were created for TOC entries.
 */
async function cmdDeleteTableOfContents({ removeBookmarks } = {}) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;

    const doc = activeEditor.state.doc;
    const tr = activeEditor.state.tr;

    // Collect all deletions in one pass, apply in a single transaction
    let tocPos = null;
    let tocNodeSize = null;
    const bookmarksToRemove = [];

    doc.descendants((node, pos) => {
        if (node.type.name === 'tableOfContents' && tocPos === null) {
            tocPos = pos;
            tocNodeSize = node.nodeSize;
        }
        if (removeBookmarks &&
            (node.type.name === 'bookmarkStart' || node.type.name === 'bookmarkEnd') &&
            (node.attrs.name?.startsWith('_Toc_') || node.attrs.id?.startsWith('_Toc_'))) {
            bookmarksToRemove.push({ pos, size: node.nodeSize });
        }
    });

    if (tocPos === null) {
        return { success: false, error: 'No table of contents found in document' };
    }

    // Combine TOC + bookmarks, sort backwards, delete in one transaction
    const deletions = [{ pos: tocPos, size: tocNodeSize }, ...bookmarksToRemove];
    deletions.sort((a, b) => b.pos - a.pos);
    for (const del of deletions) {
        tr.delete(del.pos, del.pos + del.size);
    }
    activeEditor.view.dispatch(tr);

    await saveDocument();
    debug('deleteTableOfContents: success');
    return { success: true };
}

async function cmdAddComment({ search, comment, occurrence, author }) {
    const { activeEditor, error } = requireActiveEditor();
    if (error) return error;
    if (!search) return { success: false, error: 'Search text is required' };
    if (!comment) return { success: false, error: 'Comment text is required' };

    const { match, error: searchError } = findMatch(activeEditor, search, occurrence);
    if (searchError) {
        return { success: false, error: searchError };
    }

    // Use quiet selection to avoid stealing focus from other VS Code UI
    selectRange(activeEditor, match.from, match.to, { focus: false });

    const authorName = author?.name || editor.user?.name || 'Claude';
    const authorEmail = author?.email || editor.user?.email || 'claude@anthropic.com';
    const commentId = crypto.randomUUID();

    const result = activeEditor.commands.insertComment({
        commentId,
        commentText: comment,
        creatorName: authorName,
        creatorEmail: authorEmail,
        skipEmit: true
    });

    if (!result) {
        return { success: false, error: 'Failed to add comment' };
    }

    // Register the comment in the store for persistence.
    // Use the original (non-intercepted) emit so 'add' events reach SuperDoc's
    // onEditorCommentsUpdate handler, which adds the comment to the store.
    // Use a dummy activeCommentId to prevent setActiveComment cascade.
    originalEditorEmit('commentsUpdate', {
        type: 'add',
        comment: {
            commentId,
            commentText: comment,
            creatorName: authorName,
            creatorEmail: authorEmail,
            createdTime: Date.now(),
        },
        activeCommentId: '__skip__',
    });

    // Collapse selection away from the comment mark
    activeEditor.commands.setTextSelection({ from: match.to, to: match.to });

    await saveDocument();
    debug(`addComment: added comment on "${match.text}"`);
    return { success: true, result: { commentedText: match.text } };
}

// =============================================================================
// Find & Replace
// =============================================================================

let searchMatches = [];
let currentMatchIndex = -1;
let searchBarOpen = false;
let searchCaseSensitive = false;
let replaceExpanded = false;

function toggleReplace() {
    replaceExpanded = !replaceExpanded;
    const chevron = document.getElementById('search-expand');
    const replaceRow = document.querySelector('.search-replace-row');
    chevron.classList.toggle('expanded', replaceExpanded);
    replaceRow.style.display = replaceExpanded ? 'flex' : 'none';
    if (replaceExpanded) {
        document.getElementById('replace-input').focus();
    }
}

function openSearchBar() {
    const bar = document.getElementById('search-bar');
    bar.style.display = 'flex';
    searchBarOpen = true;
    const input = document.getElementById('search-input');
    input.focus();
    input.select();
}

function closeSearchBar() {
    document.getElementById('search-bar').style.display = 'none';
    searchBarOpen = false;
    replaceExpanded = false;
    document.getElementById('search-expand').classList.remove('expanded');
    document.querySelector('.search-replace-row').style.display = 'none';
    searchMatches = [];
    currentMatchIndex = -1;
    document.getElementById('search-count').textContent = '';
    clearSearchHighlights();
    editor?.activeEditor?.view?.focus();
}

function updateSearchCount() {
    const countEl = document.getElementById('search-count');
    if (searchMatches.length === 0) {
        countEl.textContent = document.getElementById('search-input').value ? 'No results' : '';
    } else {
        countEl.textContent = `${currentMatchIndex + 1} of ${searchMatches.length}`;
    }
}

function buildSearchPattern(query) {
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return new RegExp(escaped, searchCaseSensitive ? 'g' : 'gi');
}

function clearSearchHighlights() {
    try { editor?.activeEditor?.commands.search('', { highlight: true }); } catch {}
}

// Run search, optionally preserving the current match index (e.g. after undo/replace)
function runSearch(preserveIndex = false) {
    const query = document.getElementById('search-input').value;
    if (!query || !editor?.activeEditor) {
        searchMatches = [];
        currentMatchIndex = -1;
        if (!query) clearSearchHighlights();
        updateSearchCount();
        return;
    }
    const prevIndex = currentMatchIndex;
    searchMatches = editor.activeEditor.commands.search(buildSearchPattern(query), { highlight: true }) || [];
    if (searchMatches.length > 0) {
        currentMatchIndex = preserveIndex ? Math.max(0, Math.min(prevIndex, searchMatches.length - 1)) : 0;
    } else {
        currentMatchIndex = -1;
    }
    updateSearchCount();
}

function performSearch() { runSearch(false); }

function refreshSearch() {
    if (searchBarOpen) runSearch(true);
}

function toggleCaseSensitive() {
    searchCaseSensitive = !searchCaseSensitive;
    document.getElementById('search-case').classList.toggle('active', searchCaseSensitive);
    clearSearchHighlights();
    performSearch();
    document.getElementById('search-input').focus();
}

function goToMatch(index) {
    if (searchMatches.length === 0) return;
    currentMatchIndex = index;
    editor.activeEditor.commands.goToSearchResult(searchMatches[index]);
    updateSearchCount();
}

function nextMatch() {
    if (searchMatches.length === 0) return;
    goToMatch((currentMatchIndex + 1) % searchMatches.length);
}

function prevMatch() {
    if (searchMatches.length === 0) return;
    goToMatch((currentMatchIndex - 1 + searchMatches.length) % searchMatches.length);
}

function replaceOne() {
    if (searchMatches.length === 0 || currentMatchIndex < 0) return;
    const match = searchMatches[currentMatchIndex];
    const activeEditor = editor.activeEditor;

    activeEditor.view.focus();
    activeEditor.commands.setTextSelection({ from: match.from, to: match.to });
    activeEditor.commands.insertContent(document.getElementById('replace-input').value);

    runSearch(true);
    if (searchMatches.length > 0) {
        activeEditor.commands.goToSearchResult(searchMatches[currentMatchIndex]);
    }
    scheduleAutoSave();
    document.getElementById('replace-input').focus();
}

function replaceAllMatches() {
    if (searchMatches.length === 0) return;
    const replaceValue = document.getElementById('replace-input').value;
    const activeEditor = editor.activeEditor;

    // Replace in reverse order to preserve positions
    const sorted = [...searchMatches].sort((a, b) => b.from - a.from);
    activeEditor.view.focus();
    for (const match of sorted) {
        activeEditor.commands.setTextSelection({ from: match.from, to: match.to });
        activeEditor.commands.insertContent(replaceValue);
    }

    runSearch(false);
    scheduleAutoSave();
    document.getElementById('replace-input').focus();
}

// Wire up search bar buttons
document.getElementById('search-expand').addEventListener('click', toggleReplace);
document.getElementById('search-input').addEventListener('input', performSearch);
document.getElementById('search-case').addEventListener('click', toggleCaseSensitive);
document.getElementById('search-next').addEventListener('click', nextMatch);
document.getElementById('search-prev').addEventListener('click', prevMatch);
document.getElementById('search-close').addEventListener('click', closeSearchBar);
document.getElementById('replace-one').addEventListener('click', replaceOne);
document.getElementById('replace-all').addEventListener('click', replaceAllMatches);

document.getElementById('search-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.shiftKey) {
        e.preventDefault();
        prevMatch();
    } else if (e.key === 'Enter') {
        e.preventDefault();
        nextMatch();
    } else if (e.key === 'Escape') {
        e.preventDefault();
        closeSearchBar();
    }
});

document.getElementById('replace-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
        e.preventDefault();
        replaceOne();
    } else if (e.key === 'Escape') {
        e.preventDefault();
        closeSearchBar();
    }
});

// Workaround: In VS Code webview (Electron), the click event from a tool button
// mousedown can fire AFTER Vue mounts the CommentDialog and registers its
// v-click-outside listener, causing the dialog to immediately dismiss.
// Suppress click events that are part of the same physical press as a tool mousedown.
let lastToolMousedownTime = 0;
document.addEventListener('mousedown', (e) => {
    if (e.target.closest('[data-id="is-tool"]') || e.target.closest('.superdoc__tools')) {
        lastToolMousedownTime = Date.now();
    }
}, true);
document.addEventListener('click', (e) => {
    if (Date.now() - lastToolMousedownTime < 300) {
        e.stopImmediatePropagation();
        e.preventDefault();
    }
}, true);

// Notify VS Code that the webview is ready
debug('Notifying VS Code that webview is ready');
vscode.postMessage({ type: 'ready' });

// Handle keyboard shortcuts
document.addEventListener('keydown', (event) => {
    // Ctrl/Cmd + F to open search
    if ((event.ctrlKey || event.metaKey) && event.key === 'f') {
        event.preventDefault();
        openSearchBar();
    }
    // Ctrl/Cmd + S to save
    if ((event.ctrlKey || event.metaKey) && event.key === 's') {
        event.preventDefault();
        saveDocument();
    }
});
