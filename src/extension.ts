import * as vscode from 'vscode';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

let lastLanguage: 'Arabic' | 'English' | 'None' = 'None';
let statusBarItem: vscode.StatusBarItem;

export function activate(context: vscode.ExtensionContext) {
    console.log('Auto Language  is now active!');

    statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
    context.subscriptions.push(statusBarItem);
    updateStatusBar('Ready', 'None', 'None');

    let disposable = vscode.window.onDidChangeTextEditorSelection(async (event) => {
        const editor = event.textEditor;
        const settings = vscode.workspace.getConfiguration('autolanguage');
        
        if (!settings.get('enabled')) {
            updateStatusBar('Disabled', 'None', 'None');
            return;
        }

        const selection = event.selections[0];
        const document = editor.document;
        const offset = document.offsetAt(selection.active);

        const { char, lang } = detectLanguageAround(document, offset);

        if (lang !== 'None' && lang !== lastLanguage) {
            try {
                if (lang === 'Arabic') {
                    await switchToArabic();
                } else if (lang === 'English') {
                    await switchToEnglish();
                }
                
                lastLanguage = lang;
                
                if (settings.get('showNotifications')) {
                    const flag = lang === 'Arabic' ? '🇸🇦' : '🇺🇸';
                    vscode.window.showInformationMessage(`Language Switched to ${lang} ${flag}`);
                }
            } catch (err) {
                console.error('Failed to switch language:', err);
                vscode.window.showErrorMessage('Auto Language: Failed to switch keyboard layout.');
            }
        }
        
        updateStatusBar('Active', char || ' ', lang);
    });

    context.subscriptions.push(disposable);

    let toggleDisposable = vscode.commands.registerCommand('autolanguage.toggle', () => {
        const settings = vscode.workspace.getConfiguration('autolanguage');
        const currentState = settings.get('enabled');
        settings.update('enabled', !currentState, true);
        vscode.window.showInformationMessage(`Auto Language is now ${!currentState ? 'Enabled' : 'Disabled'}`);
    });

    context.subscriptions.push(toggleDisposable);
}

function detectLanguageAround(document: vscode.TextDocument, offset: number): { char: string | null, lang: 'Arabic' | 'English' | 'None' } {
    const text = document.getText();
    const textLength = text.length;

    // Check character before
    if (offset > 0) {
        const c = text[offset - 1];
        if (c.trim()) {
            if (isArabic(c)) return { char: c, lang: 'Arabic' };
            if (isEnglish(c)) return { char: c, lang: 'English' };
        }
    }

    // Check character after
    if (offset < textLength) {
        const c = text[offset];
        if (c.trim()) {
            if (isArabic(c)) return { char: c, lang: 'Arabic' };
            if (isEnglish(c)) return { char: c, lang: 'English' };
        }
    }

    // Expanded search (up to 10 chars)
    for (let i = 1; i <= 10; i++) {
        // Backward
        const p = offset - i;
        if (p >= 0) {
            const c = text[p];
            if (c.trim()) {
                if (isArabic(c)) return { char: c, lang: 'Arabic' };
                if (isEnglish(c)) return { char: c, lang: 'English' };
            }
        }
        // Forward
        const n = offset + i - 1;
        if (n < textLength) {
            const c = text[n];
            if (c.trim()) {
                if (isArabic(c)) return { char: c, lang: 'Arabic' };
                if (isEnglish(c)) return { char: c, lang: 'English' };
            }
        }
    }

    return { char: null, lang: 'None' };
}

function isArabic(c: string): boolean {
    const code = c.charCodeAt(0);
    return (code >= 0x0600 && code <= 0x06FF) ||
           (code >= 0x0750 && code <= 0x077F) ||
           (code >= 0x08A0 && code <= 0x08FF) ||
           (code >= 0xFB50 && code <= 0xFDFF) ||
           (code >= 0xFE70 && code <= 0xFEFF);
}

function isEnglish(c: string): boolean {
    return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
}

async function switchToArabic() {
    await runPowerShellSwitch('Arabic');
}

async function switchToEnglish() {
    await runPowerShellSwitch('English');
}

async function runPowerShellSwitch(target: 'Arabic' | 'English') {
    const langId = target === 'Arabic' ? '0x01' : '0x09';
    // This script finds a layout with the primary language ID and posts the change request to the foreground window.
    const psScript = `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class KeyboardLayout {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetKeyboardLayoutList(int nBuff, [Out] IntPtr[] lpList);

    public const uint WM_INPUTLANGCHANGEREQUEST = 0x0050;

    public static IntPtr FindLayout(uint targetPrimaryId) {
        IntPtr[] layouts = new IntPtr[16];
        int count = GetKeyboardLayoutList(16, layouts);
        for (int i = 0; i < count; i++) {
            uint langId = (uint)layouts[i].ToInt64() & 0xFFFF;
            uint primaryId = langId & 0x3FF;
            if (primaryId == targetPrimaryId) return layouts[i];
        }
        return IntPtr.Zero;
    }
}
"@
$hwnd = [KeyboardLayout]::GetForegroundWindow()
$layout = [KeyboardLayout]::FindLayout(${langId})
if ($layout -ne [IntPtr]::Zero) {
    [KeyboardLayout]::PostMessage($hwnd, [KeyboardLayout]::WM_INPUTLANGCHANGEREQUEST, [IntPtr]::Zero, $layout)
}
`;
    // Minify the script for exec
    const command = `powershell -Command "${psScript.replace(/\n/g, ' ').replace(/"/g, '\\"')}"`;
    await execAsync(command);
}

function updateStatusBar(status: string, char: string, lang: string) {
    const settings = vscode.workspace.getConfiguration('autolanguage');
    if (!settings.get('showStatusBar')) {
        statusBarItem.hide();
        return;
    }

    statusBarItem.text = `$(globe) AutoLang: ${lang === 'None' ? '...' : lang}`;
    statusBarItem.tooltip = `Status: ${status}\nDetected: ${char}\nLanguage: ${lang}`;
    statusBarItem.show();
}

export function deactivate() {}
