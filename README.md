# LnkOpener

A small Visual Studio Code extension that helps you open Windows `.lnk` (shortcut) files in their target applications.

When you click a `.lnk` in the Explorer or use the command, the extension resolves the shortcut and opens the underlying target with the OS default program (for folders: Explorer, for documents: Word/Excel/other registered apps, etc.).

## Features

- Resolve `.lnk` shortcuts on Windows and open the resolved target externally.
- Support for shortcuts that include program arguments (the extension will spawn executables with their arguments when appropriate).
- Explorer context menu command: `Open .lnk target` (right-click a `.lnk` file in Explorer).
## Requirements & Platform

- This extension only supports Windows because it uses PowerShell and the WScript.Shell COM object to read `.lnk` properties.
- PowerShell must be available on PATH (standard on modern Windows). The extension runs PowerShell commands locally to read the shortcut TargetPath and Arguments.

## Install & Run (development)

1. Install dependencies and build the extension:

```bash
npm install
npm run compile
```

2. Launch the Extension Development Host (press F5 in VS Code). In the dev host:

- Click a `.lnk` in the Explorer to open its target externally.
- Or right-click a `.lnk` and choose `Open .lnk target`.

## Usage

- Clicking a `.lnk` file in Explorer (inside VS Code) will open the resolved target and the `.lnk` editor will be closed automatically.
- Use the command palette (Ctrl+Shift+P) and run `Open .lnk target` to open the selected `.lnk` file.
- Right-click a `.lnk` in Explorer and select `Open .lnk target` from the context menu.

## Limitations & Notes

- Windows-only: The extension shows a warning on non-Windows platforms.
- The extension resolves `TargetPath` and `Arguments` from the shortcut. If the shortcut points to another shortcut or an unusual target, results may vary.

## Troubleshooting

- If opening the target doesn't work:
	- Ensure you are running in Windows and PowerShell is available on PATH.
	- Check that the `.lnk` file is a valid Windows shortcut and not corrupted.
	- If a shortcut points to an executable with spaces, the extension attempts to handle quoted and unquoted paths, but very unusual cases may fail.
