import * as vscode from 'vscode';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

const execAsync = promisify(exec);

let lastLanguage: 'Arabic' | 'English' | 'None' = 'None';
let statusBarItem: vscode.StatusBarItem;

export function activate(context: vscode.ExtensionContext) {
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
                    const icon = lang === 'Arabic' ? '$(symbol-namespace)' : '$(symbol-enum)';
                    const translate = lang === "Arabic" ? "تم تغيير اللغة الى العربية" : 'Language Switched to English'
                    vscode.window.showInformationMessage(`${translate}`);
                }
            } catch (err) {
                const errorMessage = err instanceof Error ? err.message : String(err);
                // عرض رسالة خطأ مفصلة للمستخدم
                if (errorMessage.includes('not installed')) {
                    vscode.window.showErrorMessage(`Auto Language: ${errorMessage} Please add the keyboard layout in Windows Settings.`);
                } else {
                    vscode.window.showErrorMessage(`Auto Language: ${errorMessage}`);
                }
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
    // نطاقات اليونيكود العربية الكاملة:
    // 0x0600-0x06FF: العربية الأساسية
    // 0x0750-0x077F: العربية التكميلية
    // 0x08A0-0x08FF: العربية الموسعة A
    // 0xFB50-0xFDFF: أشكال العرض العربية
    // 0xFE70-0xFEFF: أشكال العرض العربية للنصوص
    // 0x0860-0x086F: العربية الموسعة B (السندية)
    // 0x08A0-0x08FF: العربية الموسعة A
    // 0x1EE00-0x1EEFF: أشكال العرض العربية الرياضية
    return (code >= 0x0600 && code <= 0x06FF) ||
           (code >= 0x0750 && code <= 0x077F) ||
           (code >= 0x08A0 && code <= 0x08FF) ||
           (code >= 0xFB50 && code <= 0xFDFF) ||
           (code >= 0xFE70 && code <= 0xFEFF) ||
           (code >= 0x0860 && code <= 0x086F) ||
           (code >= 0x1EE00 && code <= 0x1EEFF);
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
    // معرفات اللغة الصحيحة في Windows:
    // العربية: Primary Language ID = 0x01 (1)
    // الإنجليزية (الأمريكية): Primary Language ID = 0x09 (9)
    const primaryLangId = target === 'Arabic' ? '0x01' : '0x09';
    
    // This script finds a layout with the primary language ID and posts the change request to the foreground window.
    const psScript = `Add-Type -TypeDefinition @"
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
            uint primaryId = langId & 0x03FF;
            if (primaryId == targetPrimaryId) return layouts[i];
        }
        return IntPtr.Zero;
    }
}
"@
$hwnd = [KeyboardLayout]::GetForegroundWindow()
$layout = [KeyboardLayout]::FindLayout(${primaryLangId})
if ($layout -ne [IntPtr]::Zero) {
    [KeyboardLayout]::PostMessage($hwnd, [KeyboardLayout]::WM_INPUTLANGCHANGEREQUEST, [IntPtr]::Zero, $layout)
} else {
    Write-Error "Keyboard layout for ${target} (ID: ${primaryLangId}) not found. Please ensure the keyboard layout is installed."
}`;
    
    // إنشاء ملف ps1 مؤقت لتجنب مشاكل cmd.exe
    const tempDir = os.tmpdir();
    const tempScriptPath = path.join(tempDir, `autolanguage_switch_${Date.now()}.ps1`);
    
    try {
        // كتابة السكريبت إلى ملف مؤقت
        fs.writeFileSync(tempScriptPath, psScript, 'utf8');
        
        // تنفيذ الملف
        const command = `powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`;
        const { stderr } = await execAsync(command);
        
        if (stderr) {
            // تحقق إذا كان الخطأ بسبب عدم وجود تخطيط لوحة المفاتيح
            if (stderr.includes('not found') || stderr.includes('not installed')) {
                throw new Error(`Keyboard layout for ${target} is not installed on your system.`);
            }
        }
    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        throw new Error(`Failed to switch to ${target}: ${errorMessage}`);
    } finally {
        // حذف الملف المؤقت
        try {
            if (fs.existsSync(tempScriptPath)) {
                fs.unlinkSync(tempScriptPath);
            }
        } catch (cleanupError) {
            // تجاهل خطأ حذف الملف المؤقت
        }
    }
}

function updateStatusBar(status: string, char: string, lang: string) {
    const settings = vscode.workspace.getConfiguration('autolanguage');
    if (!settings.get('showStatusBar')) {
        statusBarItem.hide();
        return;
    }

    // استخدام أيقونات Codicons من VSCode
    let icon: string;
    if (lang === 'Arabic') {
        icon = '$(symbol-namespace)';
    } else if (lang === 'English') {
        icon = '$(symbol-enum)';
    } else {
        icon = '$(globe)';
    }
    
    statusBarItem.text = `${icon} AutoLang: ${lang === 'None' ? '...' : lang}`;
    statusBarItem.tooltip = `Status: ${status}\nDetected: ${char}\nLanguage: ${lang}`;
    statusBarItem.show();
}

export function deactivate() {}
