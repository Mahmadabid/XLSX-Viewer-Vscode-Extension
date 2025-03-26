import * as vscode from 'vscode';
import * as Excel from 'exceljs';

export function activate(context: vscode.ExtensionContext) {
    const provider = new XLSXEditorProvider(context);
    context.subscriptions.push(
        vscode.window.registerCustomEditorProvider('xlsxViewer.xlsx', provider, {
            webviewOptions: {
                retainContextWhenHidden: true
            },
            supportsMultipleEditorsPerDocument: false
        })
    );
}

class XLSXEditorProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private readonly context: vscode.ExtensionContext) { }

    async openCustomDocument(
        uri: vscode.Uri,
        openContext: vscode.CustomDocumentOpenContext,
        token: vscode.CancellationToken
    ): Promise<vscode.CustomDocument> {
        return { uri, dispose: () => { } };
    }

    async resolveCustomEditor(
        document: vscode.CustomDocument,
        webviewPanel: vscode.WebviewPanel,
        token: vscode.CancellationToken
    ): Promise<void> {
        try {
            const workbook = new Excel.Workbook();
            await workbook.xlsx.readFile(document.uri.fsPath);
            const worksheet = workbook.worksheets[0];
            let tableHtml = '<table id="xlsx-table" border="1" cellspacing="0" cellpadding="5">';

            worksheet.eachRow((row, rowNumber) => {
                tableHtml += '<tr>';
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    let cellValue = cell.value ? cell.value.toString() : '&nbsp;';
                    let style = '';
                    let isDefaultBlack = false;

                    // Font styles
                    if (cell.font) {
                        style += cell.font.bold ? 'font-weight:bold;' : '';
                        style += cell.font.italic ? 'font-style:italic;' : '';
                        if (cell.font.size) {
                            style += `font-size:${cell.font.size}px;`;
                        }

                        if (cell.font.color && typeof cell.font.color.argb === "string") {
                            console.log(`Row ${rowNumber}, Col ${colNumber} Font Color:`, cell.font.color.argb);
                            style += `color: ${convertARGBToRGBA(cell.font.color.argb)};`;
                        } else {
                            console.log(`Row ${rowNumber}, Col ${colNumber} has no explicit font color. Defaulting to black.`);
                            style += `color: rgb(0, 0, 0);`;  // Default black
                            isDefaultBlack = true;  // Mark for toggling
                        }
                    }

                    // Background color
                    if (cell.fill && (cell.fill as any).fgColor && (cell.fill as any).fgColor.argb) {
                        style += `background-color:${convertARGBToRGBA((cell.fill as any).fgColor.argb)};`;
                    }

                    // Add data attribute if default black
                    const dataAttr = isDefaultBlack ? 'data-default-color="true"' : '';
                    tableHtml += `<td ${dataAttr} style="${style}">${cellValue}</td>`;
                });
                tableHtml += '</tr>';
            });
            tableHtml += '</table>';

            webviewPanel.webview.options = { enableScripts: true };
            webviewPanel.webview.html = this.getWebviewContent(tableHtml);
        } catch (error) {
            vscode.window.showErrorMessage(`Error reading XLSX file: ${error}`);
        }
    }

    private getWebviewContent(content: string): string {
        return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>XLSX Viewer</title>
            <style>
                body { font-family: sans-serif; padding: 10px; }
                table { border-collapse: collapse; width: 100%; }
                th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                /* Default background */
                td { background-color: rgb(255, 255, 255); }
                /* Alternate background when toggled */
                .alt-bg td { background-color: rgb(0, 0, 0); }
                #toggleButton {
                    margin-bottom: 10px;
                    padding: 5px 10px;
                    font-size: 14px;
                }
            </style>
        </head>
        <body>
            <button id="toggleButton">Toggle Background Color</button>
            ${content}
            <script>
                const toggleButton = document.getElementById('toggleButton');
                const table = document.getElementById('xlsx-table');

                toggleButton.addEventListener('click', () => {
                    table.classList.toggle('alt-bg');

                    // Change only default black text
                    const defaultBlackCells = document.querySelectorAll('#xlsx-table td[data-default-color="true"]');
                    defaultBlackCells.forEach(cell => {
                        if (table.classList.contains('alt-bg')) {
                            cell.style.color = "rgb(255, 255, 255)"; // Change to white
                        } else {
                            cell.style.color = "rgb(0, 0, 0)"; // Change back to black
                        }
                    });
                });
            </script>
        </body>
        </html>
        `;
    }
}

export function deactivate() { }

/**
 * Converts an Excel ARGB color string ("AARRGGBB") to a CSS rgba() string.
 */
function convertARGBToRGBA(argb: string): string {
    if (argb.length !== 8) {
        return `#${argb}`;
    }
    const alpha = parseInt(argb.substring(0, 2), 16) / 255;
    const red = parseInt(argb.substring(2, 4), 16);
    const green = parseInt(argb.substring(4, 6), 16);
    const blue = parseInt(argb.substring(6, 8), 16);
    return `rgba(${red}, ${green}, ${blue}, ${alpha.toFixed(2)})`;
}
