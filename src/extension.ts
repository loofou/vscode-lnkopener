import * as vscode from 'vscode';
import { execFile, spawn } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';

const execFileAsync = promisify(execFile);

async function resolveLnk(targetPath: string): Promise<string> {
	if (process.platform !== 'win32') { throw new Error('.lnk resolution requires Windows'); }

	const psPath = targetPath.replace(/'/g, "''");
	const psCommand = `try { $s=(New-Object -ComObject WScript.Shell).CreateShortcut('${psPath}'); Write-Output $s.TargetPath; Write-Output $s.Arguments } catch { exit 1 }`;
	const { stdout } = await execFileAsync('powershell', ['-NoProfile', '-Command', psCommand], { windowsHide: true });
	const out = (stdout || '').toString().trim();
	if (!out) { throw new Error('Could not resolve target from .lnk'); }
	const lines = out.split(/\r?\n/);
	const target = (lines[0] || '').trim();
	const args = (lines[1] || '').trim();
	if (!target) { throw new Error('Could not resolve target from .lnk'); }
	return args ? `${target} ${args}` : target;
}

async function openLnkUri(uri: vscode.Uri) {
	if (process.platform !== 'win32') {
		vscode.window.showWarningMessage('.lnk handling is only supported on Windows');
		return;
	}

	try {
		const filePath = uri.fsPath;
		const resolved = await resolveLnk(filePath);

		const parts = resolved.match(/(?:"[^"]+"|\S)+/g) || [];
		if (parts.length === 0) { return; }
		const strip = (s: string) => s.replace(/^"|"$/g, '');
		let execPath = strip(parts[0]!);
		let args: string[] = parts.slice(1).map(strip);

		if (!fs.existsSync(execPath)) {
			for (let i = 2; i <= Math.min(parts.length, 6); i++) {
				const candidate = strip(parts.slice(0, i).join(' '));
				if (fs.existsSync(candidate)) {
					execPath = candidate;
					args = parts.slice(i).map(strip);
					break;
				}
			}
		}

		if (fs.existsSync(execPath) && args.length > 0) {
			spawn(execPath, args, { detached: true, stdio: 'ignore' }).unref();
			return;
		}

		const openTarget = fs.existsSync(execPath) ? execPath : resolved;
		if (fs.existsSync(openTarget)) {
			try {
				const stat = fs.statSync(openTarget);
				if (stat.isDirectory()) {
					spawn('explorer.exe', [openTarget], { detached: true, windowsHide: false, stdio: 'ignore' }).unref();
					return;
				}
			} catch { }
		}

		spawn('cmd.exe', ['/c', 'start', '', openTarget], { detached: true, windowsHide: false, stdio: 'ignore' }).unref();
	} catch (err: any) {
		vscode.window.showErrorMessage(`Failed to open .lnk: ${err?.message ?? err}`);
	}
}

export function activate(context: vscode.ExtensionContext) {
	const disposable = vscode.commands.registerCommand('lnkopener.openLnk', async (resource: vscode.Uri | undefined) => {
		const uri = resource ?? vscode.window.activeTextEditor?.document.uri;
		if (!uri) { return; }
		await openLnkUri(uri);
	});
	context.subscriptions.push(disposable);

	const opened = vscode.workspace.onDidOpenTextDocument(async (doc) => {
		try {
			if (doc.uri.fsPath.toLowerCase().endsWith('.lnk')) {
				await openLnkUri(doc.uri);
				const editor = vscode.window.visibleTextEditors.find(e => e.document.uri.toString() === doc.uri.toString());
				if (editor) {
					await vscode.window.showTextDocument(doc, { preview: true });
					await vscode.commands.executeCommand('workbench.action.closeActiveEditor');
				}
			}
		} catch { }
	});
	context.subscriptions.push(opened);

	const changeEmitter = new vscode.EventEmitter<any>();
	context.subscriptions.push(changeEmitter);

	const provider: vscode.CustomEditorProvider<vscode.CustomDocument> = {
		onDidChangeCustomDocument: changeEmitter.event,
		async openCustomDocument(uri: vscode.Uri) {
			const doc: any = { uri, dispose: () => { } };
			return doc as unknown as vscode.CustomDocument;
		},
		async resolveCustomEditor(document: vscode.CustomDocument, webviewPanel: vscode.WebviewPanel) {
			void openLnkUri(document.uri).catch(err => { try { vscode.window.showErrorMessage(`Failed to open .lnk: ${err?.message ?? err}`); } catch { } });
			try { webviewPanel.webview.html = '<!doctype html><meta charset="utf-8"><body></body>'; } catch { }
			try { webviewPanel.reveal(webviewPanel.viewColumn); } catch { }
			setTimeout(() => { try { vscode.commands.executeCommand('workbench.action.closeActiveEditor'); } catch { } }, 120);
		},
		async saveCustomDocument() { },
		async saveCustomDocumentAs() { },
		async revertCustomDocument() { },
		async backupCustomDocument() { return { id: '', delete: () => { } } as vscode.CustomDocumentBackup; }
	};

	try { context.subscriptions.push(vscode.window.registerCustomEditorProvider('lnkopener.lnkPreview', provider, { supportsMultipleEditorsPerDocument: false })); } catch { }
}

export function deactivate() { }
