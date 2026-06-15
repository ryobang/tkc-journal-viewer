import * as vscode from 'vscode';
import * as fs from 'fs/promises';
import * as path from 'path';
import { exec } from 'child_process';
import { parseTkc, ParseResult } from './parser';
import { renderHtml } from './webview';

const VIEW_TYPE = 'tkcJournal.viewer';

class TkcDocument implements vscode.CustomDocument {
  constructor(public readonly uri: vscode.Uri, public result: ParseResult) {}
  dispose(): void { /* nothing */ }
}

class TkcJournalEditorProvider implements vscode.CustomReadonlyEditorProvider<TkcDocument> {
  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(uri: vscode.Uri): Promise<TkcDocument> {
    const buf = await readUri(uri);
    const result = await parseTkc(buf, uri.fsPath);
    return new TkcDocument(uri, result);
  }

  async resolveCustomEditor(
    document: TkcDocument,
    panel: vscode.WebviewPanel,
  ): Promise<void> {
    panel.webview.options = {
      enableScripts: true,
      localResourceRoots: [vscode.Uri.file(path.join(this.context.extensionPath, 'media'))],
    };

    const fileName = path.basename(document.uri.fsPath);
    panel.webview.html = renderHtml({
      fileName,
      result: document.result,
      cspSource: panel.webview.cspSource,
      config: getConfig(),
    });

    this.wireMessages(panel, document, fileName);
  }

  /** Webview からのメッセージを処理する。本体パネル・ポップアップパネル共通。 */
  private wireMessages(
    panel: vscode.WebviewPanel,
    document: TkcDocument,
    fileName: string,
    initial?: { filter?: string; month?: string; fontPx?: number },
  ): void {
    panel.webview.onDidReceiveMessage(
      async (msg: { type: string; filter?: string; month?: string; fontPx?: number }) => {
        if (msg.type === 'openInExcel') {
          await openInExcel(document.uri);
        } else if (msg.type === 'reload') {
          const buf = await readUri(document.uri);
          document.result = await parseTkc(buf, document.uri.fsPath);
          panel.webview.html = renderHtml({
            fileName,
            result: document.result,
            cspSource: panel.webview.cspSource,
            config: getConfig(),
            initial,
          });
        } else if (msg.type === 'popout') {
          await this.openPopout(document, fileName, {
            filter: msg.filter,
            month: msg.month,
            fontPx: msg.fontPx,
          });
        }
      },
    );
  }

  /**
   * 表示中の内容を独立した Webview パネルで開き直し、
   * VSCode のフローティングウィンドウへ移動させる。
   * これにより別モニターへドラッグして配置できる。
   */
  private async openPopout(
    document: TkcDocument,
    fileName: string,
    initial: { filter?: string; month?: string; fontPx?: number },
  ): Promise<void> {
    const panel = vscode.window.createWebviewPanel(
      'tkcJournal.popout',
      `${fileName}（別ウィンドウ）`,
      vscode.ViewColumn.Active,
      {
        enableScripts: true,
        retainContextWhenHidden: true,
        localResourceRoots: [
          vscode.Uri.file(path.join(this.context.extensionPath, 'media')),
        ],
      },
    );

    panel.webview.html = renderHtml({
      fileName,
      result: document.result,
      cspSource: panel.webview.cspSource,
      config: getConfig(),
      initial,
    });

    this.wireMessages(panel, document, fileName, initial);

    // たった今アクティブになったこのパネルを別ウィンドウ(補助ウィンドウ)へ移動。
    // 補助ウィンドウは OS ウィンドウなので別モニターへ移動できる。
    try {
      await vscode.commands.executeCommand('workbench.action.moveEditorToNewWindow');
    } catch (err) {
      vscode.window.showWarningMessage(
        `別ウィンドウへの移動に失敗しました。タブをドラッグして手動で別ウィンドウ化してください。 (${
          err instanceof Error ? err.message : String(err)
        })`,
      );
    }
  }
}

async function readUri(uri: vscode.Uri): Promise<Buffer> {
  if (uri.scheme === 'file') {
    return await fs.readFile(uri.fsPath);
  }
  const data = await vscode.workspace.fs.readFile(uri);
  return Buffer.from(data);
}

function getConfig() {
  const c = vscode.workspace.getConfiguration('tkcJournal');
  return {
    dateFormat: c.get<string>('dateFormat', 'MM/DD'),
    zeroAmount: c.get<string>('zeroAmount', 'blank'),
  };
}

export function activate(context: vscode.ExtensionContext): void {
  const provider = new TkcJournalEditorProvider(context);
  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(VIEW_TYPE, provider, {
      webviewOptions: { retainContextWhenHidden: true },
      supportsMultipleEditorsPerDocument: false,
    }),
  );

  context.subscriptions.push(
    vscode.commands.registerCommand('tkcJournal.openWith', async (uri?: vscode.Uri) => {
      const target = uri ?? vscode.window.activeTextEditor?.document.uri;
      if (!target) {
        vscode.window.showWarningMessage('開く .xlsx ファイルが指定されていません');
        return;
      }
      await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
    }),
  );
}

export function deactivate(): void { /* nothing */ }

async function openInExcel(uri: vscode.Uri): Promise<void> {
  if (uri.scheme !== 'file') {
    vscode.window.showWarningMessage('ローカルファイルのみExcelで開けます');
    return;
  }
  const filePath = uri.fsPath;
  if (process.platform === 'win32') {
    // Use cmd `start` to open with default app (Excel for .xlsx)
    exec(`start "" "${filePath}"`, (err) => {
      if (err) {
        vscode.window.showErrorMessage(`Excelで開けませんでした: ${err.message}`);
      }
    });
  } else if (process.platform === 'darwin') {
    exec(`open "${filePath}"`, (err) => {
      if (err) vscode.window.showErrorMessage(`開けませんでした: ${err.message}`);
    });
  } else {
    exec(`xdg-open "${filePath}"`, (err) => {
      if (err) vscode.window.showErrorMessage(`開けませんでした: ${err.message}`);
    });
  }
}
