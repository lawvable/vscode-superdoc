# **SuperDoc VS Code Extension**

<img src="logo.png" alt="SuperDoc Logo" width="200">

Edit and view DOCX files inside Visual Studio Code with [SuperDoc](https://github.com/superdoc-dev/superdoc).

## **Features**

- **Edit DOCX in VS Code** - Keep your code and documents open side-by-side
- **Live reload** - When an AI agent or external process modifies your file, your document automatically refreshes
- **Auto-save** - Changes are saved as you type

## **Usage**

Once installed, any `.docx` file you open will automatically use SuperDoc. Just open a file and start editing.

## **Install**

To be notified when live on Marketplace, star this repository.

### From source (development)

```bash
npm install
npm run compile
```

Then press `F5` in VS Code to launch a development window.

### Deploy to an existing VS Code installation

After building, copy the compiled files and native dependencies to the installed extension:

```bash
npm run compile

EXT=~/.vscode/extensions/superdoc.superdoc-vscode-extension-0.1.0

cp dist/extension.js "$EXT/dist/"
cp dist/webview/main.js "$EXT/dist/webview/"
cp webview/style.css "$EXT/dist/webview/"

for pkg in @parcel detect-libc is-glob is-extglob node-addon-api picomatch; do
  cp -r "node_modules/$pkg" "$EXT/node_modules/"
done
```

Then reload VS Code (`Cmd+Shift+P` → "Developer: Reload Window").

## **License**

This project is licensed under the GNU Affero General Public License version 3.0 (AGPLv3). See the full license at [gnu.org/licenses/agpl-3.0](https://www.gnu.org/licenses/agpl-3.0.en.html).
