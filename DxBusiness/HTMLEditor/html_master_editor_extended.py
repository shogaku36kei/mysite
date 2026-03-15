#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HTML Master Editor - VSCodeを超えるHTML特化エディタ
拡張版：.bat/.ps1/.js/.json/.xml/.css/.vbs/.vba/.csv/.py 対応
"""

import sys
import os
import re
import json
import csv
import io
import tempfile
from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter,
    QVBoxLayout, QHBoxLayout, QTextEdit, QPlainTextEdit,
    QTreeWidget, QTreeWidgetItem, QLabel, QPushButton,
    QToolBar, QStatusBar, QFileDialog, QMessageBox,
    QTabWidget, QListWidget, QListWidgetItem, QFrame,
    QCompleter, QDialog, QDialogButtonBox, QCheckBox,
    QLineEdit, QGroupBox, QScrollArea, QSizePolicy,
    QMenu, QInputDialog
)
from PyQt6.QtCore import (
    Qt, QTimer, QThread, pyqtSignal, QRect, QSize,
    QStringListModel, QPoint, QUrl, QRegularExpression
)
from PyQt6.QtGui import (
    QFont, QColor, QPainter, QTextFormat, QTextCursor,
    QSyntaxHighlighter, QTextCharFormat, QPalette,
    QKeySequence, QAction, QIcon, QTextDocument,
    QFontMetrics
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEnginePage

import lxml.etree as etree
from pygments import highlight
from pygments.lexers import HtmlLexer, CssLexer, JavascriptLexer
from pygments.formatters import HtmlFormatter


# ============================================================
# カラーテーマ定義
# ============================================================
THEME = {
    "bg_editor":      "#1E1E2E",
    "bg_panel":       "#181825",
    "bg_line":        "#313244",
    "fg_text":        "#CDD6F4",
    "fg_comment":     "#6C7086",
    "fg_tag":         "#89B4FA",
    "fg_attr":        "#A6E3A1",
    "fg_value":       "#F38BA8",
    "fg_string":      "#FAB387",
    "fg_doctype":     "#CBA6F7",
    "fg_entity":      "#F9E2AF",
    "fg_css_prop":    "#89DCEB",
    "fg_css_val":     "#A6E3A1",
    "fg_js_keyword":  "#CBA6F7",
    "fg_js_func":     "#89B4FA",
    "fg_js_string":   "#A6E3A1",
    "accent":         "#89B4FA",
    "accent2":        "#A6E3A1",
    "error":          "#F38BA8",
    "warning":        "#F9E2AF",
    "current_line":   "#2A2A3E",
    "selection":      "#45475A",
    "border":         "#45475A",
    "gutter_bg":      "#181825",
    "gutter_fg":      "#585B70",
}

# ============================================================
# 言語モード定義
# ============================================================
LANGUAGE_MAP = {
    ".html": "HTML",   ".htm": "HTML",
    ".css":  "CSS",
    ".js":   "JavaScript",
    ".json": "JSON",
    ".xml":  "XML",
    ".py":   "Python",
    ".bat":  "Batch",
    ".ps1":  "PowerShell",
    ".vbs":  "VBScript",
    ".vba":  "VBA",
    ".csv":  "CSV",
}

FILE_FILTER = (
    "すべての対応ファイル (*.html *.htm *.css *.js *.json *.xml "
    "*.py *.bat *.ps1 *.vbs *.vba *.csv);;"
    "HTML Files (*.html *.htm);;"
    "CSS Files (*.css);;"
    "JavaScript Files (*.js);;"
    "JSON Files (*.json);;"
    "XML Files (*.xml);;"
    "Python Files (*.py);;"
    "Batch Files (*.bat);;"
    "PowerShell Files (*.ps1);;"
    "VBScript Files (*.vbs);;"
    "VBA Files (*.vba);;"
    "CSV Files (*.csv);;"
    "All Files (*)"
)


# ============================================================
# ユーティリティ：フォーマット定義ヘルパー
# ============================================================
def _fmt(color, bold=False, italic=False):
    fmt = QTextCharFormat()
    fmt.setForeground(QColor(color))
    if bold:
        fmt.setFontWeight(700)
    if italic:
        fmt.setFontItalic(True)
    return fmt


# ============================================================
# シンタックスハイライター基底クラス
# ============================================================
class BaseHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)
        self.rules = []
        self._build_rules()

    def _build_rules(self):
        pass

    def highlightBlock(self, text):
        self.setFormat(0, len(text), _fmt(THEME["fg_text"]))
        for pattern, fmt in self.rules:
            it = pattern.globalMatch(text)
            while it.hasNext():
                m = it.next()
                self.setFormat(m.capturedStart(), m.capturedLength(), fmt)


# ============================================================
# HTML シンタックスハイライター
# ============================================================
class HtmlSyntaxHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)
        self._build_rules()

    def _build_rules(self):
        self.rules = []
        self.rules.append((
            QRegularExpression(r'<!DOCTYPE[^>]*>',
                QRegularExpression.PatternOption.CaseInsensitiveOption),
            _fmt(THEME["fg_doctype"], bold=True)
        ))
        self.rules.append((
            QRegularExpression(r'<!--[\s\S]*?-->'),
            _fmt(THEME["fg_comment"], italic=True)
        ))
        self.rules.append((
            QRegularExpression(r'</?([a-zA-Z][a-zA-Z0-9\-]*)'),
            _fmt(THEME["fg_tag"], bold=True)
        ))
        self.rules.append((
            QRegularExpression(r'\b([a-zA-Z\-]+)(?=\s*=)'),
            _fmt(THEME["fg_attr"])
        ))
        self.rules.append((
            QRegularExpression(r'"[^"]*"'),
            _fmt(THEME["fg_value"])
        ))
        self.rules.append((
            QRegularExpression(r"'[^']*'"),
            _fmt(THEME["fg_string"])
        ))
        self.rules.append((
            QRegularExpression(r'&[a-zA-Z#][a-zA-Z0-9]*;'),
            _fmt(THEME["fg_entity"])
        ))
        self.rules.append((
            QRegularExpression(r'[<>/]'),
            _fmt(THEME["fg_tag"])
        ))
        self.css_rules = [
            (QRegularExpression(r'[\w\-]+(?=\s*:)'), _fmt(THEME["fg_css_prop"])),
            (QRegularExpression(r':\s*[^;"}]+'),      _fmt(THEME["fg_css_val"])),
        ]
        js_kw = (r'\b(var|let|const|function|return|if|else|for|while|class|new|'
                 r'this|import|export|from|async|await|try|catch|typeof|null|'
                 r'undefined|true|false)\b')
        self.js_rules = [
            (QRegularExpression(js_kw),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(r'\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*\()'),
             _fmt(THEME["fg_js_func"])),
            (QRegularExpression(r'"[^"]*"|\'[^\']*\'|`[^`]*`'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'//[^\n]*'),
             _fmt(THEME["fg_comment"], italic=True)),
        ]

    def highlightBlock(self, text):
        self.setFormat(0, len(text), _fmt(THEME["fg_text"]))
        for pattern, fmt in self.rules:
            it = pattern.globalMatch(text)
            while it.hasNext():
                m = it.next()
                self.setFormat(m.capturedStart(), m.capturedLength(), fmt)
        self._apply_block_state(text)

    def _apply_block_state(self, text):
        prev  = self.previousBlockState()
        state = prev if prev in (1, 2) else 0
        i = 0
        while i < len(text):
            if state == 0:
                sm  = re.search(r'<style[^>]*>',  text[i:], re.IGNORECASE)
                scm = re.search(r'<script[^>]*>', text[i:], re.IGNORECASE)
                if sm and (not scm or sm.start() < scm.start()):
                    i += sm.end(); state = 1
                elif scm:
                    i += scm.end(); state = 2
                else:
                    break
            elif state == 1:
                em  = re.search(r'</style>', text[i:], re.IGNORECASE)
                end = i + em.start() if em else len(text)
                chunk = text[i:end]
                for pat, fmt in self.css_rules:
                    it = pat.globalMatch(chunk)
                    while it.hasNext():
                        m = it.next()
                        self.setFormat(i + m.capturedStart(), m.capturedLength(), fmt)
                if em: i += em.end(); state = 0
                else:  break
            elif state == 2:
                em  = re.search(r'</script>', text[i:], re.IGNORECASE)
                end = i + em.start() if em else len(text)
                chunk = text[i:end]
                for pat, fmt in self.js_rules:
                    it = pat.globalMatch(chunk)
                    while it.hasNext():
                        m = it.next()
                        self.setFormat(i + m.capturedStart(), m.capturedLength(), fmt)
                if em: i += em.end(); state = 0
                else:  break
        self.setCurrentBlockState(state)


# ============================================================
# CSS シンタックスハイライター
# ============================================================
class CssSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        self.rules = [
            (QRegularExpression(r'/\*[\s\S]*?\*/'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'[.#]?[a-zA-Z][a-zA-Z0-9_\-]*(?=\s*\{)'),
             _fmt(THEME["fg_tag"], bold=True)),
            (QRegularExpression(r'[\w\-]+(?=\s*:)'),
             _fmt(THEME["fg_css_prop"])),
            (QRegularExpression(r':\s*[^;{}]+'),
             _fmt(THEME["fg_css_val"])),
            (QRegularExpression(r'"[^"]*"|\'[^\']*\''),
             _fmt(THEME["fg_string"])),
            (QRegularExpression(r'\b\d+(\.\d+)?(px|em|rem|%|vh|vw|pt|s|ms)?\b'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(r'!important'),
             _fmt(THEME["fg_value"], bold=True)),
            (QRegularExpression(r'@[a-zA-Z\-]+'),
             _fmt(THEME["fg_doctype"], bold=True)),
        ]


# ============================================================
# JavaScript シンタックスハイライター
# ============================================================
class JsSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        kw = (r'\b(var|let|const|function|return|if|else|for|while|do|switch|case|'
              r'break|continue|class|new|this|super|import|export|default|from|of|in|'
              r'async|await|try|catch|finally|throw|typeof|instanceof|void|delete|'
              r'null|undefined|true|false|NaN|Infinity)\b')
        self.rules = [
            (QRegularExpression(r'/\*[\s\S]*?\*/'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'//[^\n]*'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'`[^`]*`'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'"[^"\\]*(?:\\.[^"\\]*)*"|\'[^\'\\]*(?:\\.[^\'\\]*)*\''),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(kw),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(r'\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*\()'),
             _fmt(THEME["fg_js_func"])),
            (QRegularExpression(r'\b0x[0-9a-fA-F]+|\b\d+(\.\d+)?\b'),
             _fmt(THEME["fg_entity"])),
        ]


# ============================================================
# JSON シンタックスハイライター
# ============================================================
class JsonSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        self.rules = [
            (QRegularExpression(r'"[^"]*"(?=\s*:)'),
             _fmt(THEME["fg_attr"], bold=True)),
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'\b-?\d+(\.\d+)?([eE][+-]?\d+)?\b'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(r'\b(true|false|null)\b'),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(r'[{}\[\]]'),
             _fmt(THEME["fg_tag"])),
        ]


# ============================================================
# XML シンタックスハイライター
# ============================================================
class XmlSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        self.rules = [
            (QRegularExpression(r'<!--[\s\S]*?-->'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'<!\[CDATA\[[\s\S]*?\]\]>'),
             _fmt(THEME["fg_string"])),
            (QRegularExpression(r'<\?[^?]*\?>|<!DOCTYPE[^>]*>',
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_doctype"], bold=True)),
            (QRegularExpression(r'</?([a-zA-Z][a-zA-Z0-9:_\-]*)'),
             _fmt(THEME["fg_tag"], bold=True)),
            (QRegularExpression(r'\b([a-zA-Z:_][a-zA-Z0-9:_\-]*)(?=\s*=)'),
             _fmt(THEME["fg_attr"])),
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_value"])),
            (QRegularExpression(r"'[^']*'"),
             _fmt(THEME["fg_string"])),
            (QRegularExpression(r'&[a-zA-Z#][a-zA-Z0-9]*;'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(r'[<>/]'),
             _fmt(THEME["fg_tag"])),
        ]


# ============================================================
# Python シンタックスハイライター
# ============================================================
class PythonSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        kw = (r'\b(False|None|True|and|as|assert|async|await|break|class|continue|'
              r'def|del|elif|else|except|finally|for|from|global|if|import|in|is|'
              r'lambda|nonlocal|not|or|pass|raise|return|try|while|with|yield)\b')
        builtin = (r'\b(print|len|range|type|int|str|float|list|dict|set|tuple|bool|'
                   r'open|input|super|self|cls|enumerate|zip|map|filter|sorted|'
                   r'reversed|isinstance|hasattr|getattr|setattr|staticmethod|'
                   r'classmethod|property)\b')
        self.rules = [
            (QRegularExpression(r'"""[\s\S]*?"""|\'\'\'[\s\S]*?\'\'\''),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'#[^\n]*'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'f"[^"]*"|f\'[^\']*\''),
             _fmt(THEME["fg_string"])),
            (QRegularExpression(r'"[^"\\]*(?:\\.[^"\\]*)*"|\'[^\'\\]*(?:\\.[^\'\\]*)*\''),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(kw),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(builtin),
             _fmt(THEME["fg_css_prop"])),
            (QRegularExpression(r'@[a-zA-Z_][a-zA-Z0-9_.]*'),
             _fmt(THEME["fg_doctype"], bold=True)),
            (QRegularExpression(r'(?<=def\s)[a-zA-Z_][a-zA-Z0-9_]*'),
             _fmt(THEME["fg_js_func"], bold=True)),
            (QRegularExpression(r'(?<=class\s)[a-zA-Z_][a-zA-Z0-9_]*'),
             _fmt(THEME["fg_tag"], bold=True)),
            (QRegularExpression(r'\b\d+(\.\d+)?\b'),
             _fmt(THEME["fg_entity"])),
        ]


# ============================================================
# Batch シンタックスハイライター
# ============================================================
class BatchSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        kw = (r'\b(echo|set|if|else|goto|call|exit|pause|for|do|in|not|'
              r'exist|errorlevel|defined|rem|cls|dir|copy|move|del|mkdir|'
              r'rmdir|cd|pushd|popd|start|taskkill|findstr|find|sort|'
              r'type|more|choice|timeout|ping|ipconfig|netstat)\b')
        self.rules = [
            (QRegularExpression(r'(?i)^(\s*)(rem\b|::)[^\n]*',
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'^\s*:[a-zA-Z_][a-zA-Z0-9_]*'),
             _fmt(THEME["fg_doctype"], bold=True)),
            (QRegularExpression(r'%[^%\n]+%'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(r'![^!\n]+!'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(kw,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'[|<>]+'),
             _fmt(THEME["fg_tag"])),
        ]


# ============================================================
# PowerShell シンタックスハイライター
# ============================================================
class PowerShellSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        kw = (r'\b(if|else|elseif|switch|for|foreach|while|do|until|break|'
              r'continue|return|function|param|begin|process|end|try|catch|'
              r'finally|throw|exit|class|enum|True|False|Null)\b')
        cmdlet = (r'\b(Write-Host|Write-Output|Write-Error|Write-Warning|'
                  r'Get-Item|Set-Item|Remove-Item|New-Item|Copy-Item|Move-Item|'
                  r'Get-Content|Set-Content|Add-Content|Get-ChildItem|'
                  r'Get-Process|Stop-Process|Start-Process|'
                  r'Invoke-Command|Invoke-Expression|Import-Module|'
                  r'ForEach-Object|Where-Object|Select-Object|Sort-Object|'
                  r'Get-Variable|Set-Variable|Test-Path|Split-Path|'
                  r'Join-Path|Resolve-Path)\b')
        self.rules = [
            (QRegularExpression(r'#[^\n]*'),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'@"[\s\S]*?"@|@\'[\s\S]*?\'@'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'"[^"]*"|\'[^\']*\''),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'\$[a-zA-Z_][a-zA-Z0-9_]*'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(cmdlet,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_css_prop"], bold=True)),
            (QRegularExpression(kw,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(r'\b\d+(\.\d+)?\b'),
             _fmt(THEME["fg_entity"])),
            (QRegularExpression(r'-[a-zA-Z]+'),
             _fmt(THEME["fg_attr"])),
            (QRegularExpression(r'\|'),
             _fmt(THEME["fg_tag"], bold=True)),
        ]


# ============================================================
# VBScript シンタックスハイライター
# ============================================================
class VbsSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        kw = (r'\b(Dim|Set|Let|If|Then|Else|ElseIf|End|For|Each|Next|'
              r'While|Wend|Do|Loop|Until|Select|Case|Sub|Function|Class|'
              r'Property|Get|Call|Return|Exit|On|Error|Resume|'
              r'GoTo|And|Or|Not|Is|Mod|Eqv|Imp|Xor|'
              r'True|False|Nothing|Null|Empty)\b')
        builtin = (r'\b(MsgBox|InputBox|WScript|CreateObject|GetObject|'
                   r'CStr|CInt|CLng|CDbl|CBool|CDate|'
                   r'Len|Left|Right|Mid|InStr|Replace|Split|Join|'
                   r'UBound|LBound|Array|IsArray|IsNull|IsEmpty|IsObject|'
                   r'Now|Date|Time|DateAdd|DateDiff|Int|Abs|Sqr|Round)\b')
        self.rules = [
            (QRegularExpression(r"'[^\n]*|(?i)\bRem\b[^\n]*"),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(kw,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            (QRegularExpression(builtin,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_css_prop"])),
            (QRegularExpression(r'\b\d+(\.\d+)?\b'),
             _fmt(THEME["fg_entity"])),
        ]


# ============================================================
# VBA シンタックスハイライター（Excel/Access/Word対応）
# ============================================================
class VbaSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        # VBA 言語キーワード
        kw = (r'\b(Dim|Set|Let|If|Then|Else|ElseIf|End|For|Each|Next|'
              r'While|Wend|Do|Loop|Until|Select|Case|Sub|Function|'
              r'Property|Get|Let|Set|Call|Return|Exit|On|Error|Resume|'
              r'GoTo|GoSub|And|Or|Not|Is|Mod|Eqv|Imp|Xor|Like|'
              r'True|False|Nothing|Null|Empty|'
              r'Public|Private|Protected|Friend|Static|Const|'
              r'ByVal|ByRef|Optional|ParamArray|'
              r'New|With|WithEvents|Implements|'
              r'ReDim|Preserve|Erase|'
              r'Type|Enum|Class|Module|'
              r'Option|Explicit|Base|Compare|'
              r'As|Integer|Long|Single|Double|String|Boolean|'
              r'Byte|Currency|Date|Object|Variant|LongLong|LongPtr)\b')

        # Excel オブジェクト・プロパティ
        excel_obj = (r'\b(Application|Workbook|Workbooks|Worksheet|Worksheets|'
                     r'ActiveWorkbook|ActiveSheet|ActiveCell|ActiveWindow|'
                     r'ThisWorkbook|Selection|Range|Cells|Rows|Columns|'
                     r'Cell|Row|Column|Offset|Resize|EntireRow|EntireColumn|'
                     r'UsedRange|CurrentRegion|End|Interior|Font|Border|Borders|'
                     r'Chart|Charts|ChartObject|ChartObjects|'
                     r'PivotTable|PivotTables|ListObject|ListObjects|'
                     r'Name|Names|Comment|Comments|'
                     r'Shape|Shapes|Picture|Pictures|'
                     r'CommandButton|TextBox|Label|ComboBox|ListBox|'
                     r'UserForm|Controls|'
                     r'xlUp|xlDown|xlLeft|xlRight|xlToLeft|xlToRight|'
                     r'xlEdgeLeft|xlEdgeTop|xlEdgeBottom|xlEdgeRight|'
                     r'xlContinuous|xlThin|xlMedium|xlThick|'
                     r'xlSolid|xlNone|xlAutomatic|'
                     r'vbYes|vbNo|vbOK|vbCancel|vbYesNo|vbOKCancel|'
                     r'vbInformation|vbQuestion|vbExclamation|vbCritical|'
                     r'vbCrLf|vbCr|vbLf|vbTab|vbNullString)\b')

        # VBA 組み込み関数
        builtin = (r'\b(MsgBox|InputBox|'
                   r'CStr|CInt|CLng|CDbl|CBool|CDate|CSng|CCur|CByte|'
                   r'Len|Left|Right|Mid|InStr|InStrRev|Replace|Split|Join|'
                   r'Trim|LTrim|RTrim|UCase|LCase|StrComp|StrConv|'
                   r'UBound|LBound|Array|IsArray|IsNull|IsEmpty|IsObject|'
                   r'IsNumeric|IsDate|IsError|IsMissing|'
                   r'Now|Date|Time|DateAdd|DateDiff|DatePart|DateSerial|'
                   r'TimeSerial|Year|Month|Day|Hour|Minute|Second|Weekday|'
                   r'Int|Fix|Abs|Sqr|Round|Rnd|Randomize|'
                   r'Asc|Chr|Hex|Oct|Val|Str|Format|'
                   r'Dir|FileLen|FileDateTime|GetAttr|SetAttr|'
                   r'Open|Close|Print|Write|Input|Line|Get|Put|'
                   r'Shell|Environ|CreateObject|GetObject|'
                   r'Debug|Print|Assert|'
                   r'Err|Error|CVErr)\b')

        self.rules = [
            # コメント（' または Rem）
            (QRegularExpression(r"'[^\n]*"),
             _fmt(THEME["fg_comment"], italic=True)),
            (QRegularExpression(r'(?i)\bRem\b[^\n]*'),
             _fmt(THEME["fg_comment"], italic=True)),
            # 文字列
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_js_string"])),
            # Excel オブジェクト（タグ色で強調）
            (QRegularExpression(excel_obj,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_tag"], bold=True)),
            # VBA キーワード
            (QRegularExpression(kw,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_js_keyword"], bold=True)),
            # 組み込み関数
            (QRegularExpression(builtin,
                QRegularExpression.PatternOption.CaseInsensitiveOption),
             _fmt(THEME["fg_css_prop"])),
            # 数値
            (QRegularExpression(r'\b\d+(\.\d+)?\b'),
             _fmt(THEME["fg_entity"])),
            # プロパティアクセス（.Value .Name 等）
            (QRegularExpression(r'\.[A-Z][a-zA-Z0-9_]*'),
             _fmt(THEME["fg_attr"])),
        ]


# ============================================================
# CSV シンタックスハイライター
# ============================================================
class CsvSyntaxHighlighter(BaseHighlighter):
    def _build_rules(self):
        self.rules = [
            (QRegularExpression(r'"[^"]*"'),
             _fmt(THEME["fg_js_string"])),
            (QRegularExpression(r'(?<!["\w])-?\d+(\.\d+)?(?!["\w])'),
             _fmt(THEME["fg_entity"])),
        ]

    def highlightBlock(self, text):
        self.setFormat(0, len(text), _fmt(THEME["fg_text"]))
        if self.currentBlock().blockNumber() == 0:
            self.setFormat(0, len(text), _fmt(THEME["fg_attr"], bold=True))
            return
        parts = text.split(',')
        pos = 0
        colors = [THEME["fg_text"], THEME["fg_css_val"]]
        for idx, part in enumerate(parts):
            self.setFormat(pos, len(part), _fmt(colors[idx % 2]))
            pos += len(part) + 1
        for pattern, fmt in self.rules:
            it = pattern.globalMatch(text)
            while it.hasNext():
                m = it.next()
                self.setFormat(m.capturedStart(), m.capturedLength(), fmt)


# ============================================================
# ハイライター選択ファクトリ
# ============================================================
def create_highlighter(ext: str, document):
    ext = ext.lower()
    mapping = {
        ".html": HtmlSyntaxHighlighter,
        ".htm":  HtmlSyntaxHighlighter,
        ".css":  CssSyntaxHighlighter,
        ".js":   JsSyntaxHighlighter,
        ".json": JsonSyntaxHighlighter,
        ".xml":  XmlSyntaxHighlighter,
        ".py":   PythonSyntaxHighlighter,
        ".bat":  BatchSyntaxHighlighter,
        ".ps1":  PowerShellSyntaxHighlighter,
        ".vbs":  VbsSyntaxHighlighter,
        ".vba":  VbaSyntaxHighlighter,
        ".csv":  CsvSyntaxHighlighter,
    }
    return mapping.get(ext, HtmlSyntaxHighlighter)(document)


# ============================================================
# 行番号ウィジェット
# ============================================================
class LineNumberArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor

    def sizeHint(self):
        return QSize(self.editor.line_number_area_width(), 0)

    def paintEvent(self, event):
        self.editor.line_number_area_paint_event(event)


# ============================================================
# コードエディタ本体
# ============================================================
class CodeEditor(QPlainTextEdit):
    cursorLineChanged = pyqtSignal(int)

    HTML_SNIPPETS = {
        "html": '<!DOCTYPE html>\n<html lang="ja">\n<head>\n    <meta charset="UTF-8">\n    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n    <title>タイトル</title>\n</head>\n<body>\n    \n</body>\n</html>',
        "div":     '<div class=""></div>',
        "p":       "<p></p>",
        "a":       '<a href=""></a>',
        "img":     '<img src="" alt="">',
        "ul":      "<ul>\n    <li></li>\n</ul>",
        "ol":      "<ol>\n    <li></li>\n</ol>",
        "table":   "<table>\n    <thead>\n        <tr><th></th></tr>\n    </thead>\n    <tbody>\n        <tr><td></td></tr>\n    </tbody>\n</table>",
        "form":    '<form action="" method="post">\n    \n</form>',
        "input":   '<input type="text" name="" id="">',
        "button":  '<button type="button"></button>',
        "section": "<section>\n    \n</section>",
        "header":  "<header>\n    \n</header>",
        "footer":  "<footer>\n    \n</footer>",
        "nav":     "<nav>\n    \n</nav>",
        "main":    "<main>\n    \n</main>",
        "article": "<article>\n    \n</article>",
        "aside":   "<aside>\n    \n</aside>",
        "span":    '<span class=""></span>',
        "h1": "<h1></h1>", "h2": "<h2></h2>", "h3": "<h3></h3>",
        "style":   "<style>\n    \n</style>",
        "script":  "<script>\n    \n</script>",
        "link":    '<link rel="stylesheet" href="">',
        "meta":    '<meta name="" content="">',
    }

    VOID_TAGS = {"area","base","br","col","embed","hr","img","input",
                 "link","meta","param","source","track","wbr"}

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_ext = ".html"
        self._setup_appearance()
        self.line_number_area = LineNumberArea(self)
        self.highlighter = create_highlighter(".html", self.document())

        self.blockCountChanged.connect(self.update_line_number_area_width)
        self.updateRequest.connect(self.update_line_number_area)
        self.cursorPositionChanged.connect(self._on_cursor_changed)
        self.cursorPositionChanged.connect(self.highlight_current_line)

        self.update_line_number_area_width(0)
        self.highlight_current_line()

        self._completer_list = QListWidget()
        self._completer_list.setWindowFlags(Qt.WindowType.Popup)
        self._completer_list.setStyleSheet(f"""
            QListWidget {{
                background: {THEME['bg_panel']};
                color: {THEME['fg_text']};
                border: 1px solid {THEME['accent']};
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
            }}
            QListWidget::item:selected {{
                background: {THEME['accent']};
                color: {THEME['bg_editor']};
            }}
            QListWidget::item:hover {{ background: {THEME['bg_line']}; }}
        """)
        self._completer_list.itemActivated.connect(self._insert_completion)
        self._completer_list.hide()
        self._current_word = ""

    def set_language(self, ext: str):
        if ext == self._current_ext:
            return
        self._current_ext = ext
        self.highlighter.setDocument(None)
        self.highlighter = create_highlighter(ext, self.document())

    def _setup_appearance(self):
        font = QFont("Consolas", 11)
        font.setStyleHint(QFont.StyleHint.Monospace)
        self.setFont(font)
        self.setStyleSheet(f"""
            QPlainTextEdit {{
                background-color: {THEME['bg_editor']};
                color: {THEME['fg_text']};
                border: none;
                selection-background-color: {THEME['selection']};
            }}
        """)
        self.setTabStopDistance(
            QFontMetrics(self.font()).horizontalAdvance(' ') * 4)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)

    def line_number_area_width(self):
        digits = max(3, len(str(self.blockCount())))
        return 10 + self.fontMetrics().horizontalAdvance('9') * digits

    def update_line_number_area_width(self, _):
        self.setViewportMargins(self.line_number_area_width(), 0, 0, 0)

    def update_line_number_area(self, rect, dy):
        if dy:
            self.line_number_area.scroll(0, dy)
        else:
            self.line_number_area.update(0, rect.y(),
                self.line_number_area.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.update_line_number_area_width(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.line_number_area.setGeometry(
            QRect(cr.left(), cr.top(),
                  self.line_number_area_width(), cr.height()))

    def line_number_area_paint_event(self, event):
        painter = QPainter(self.line_number_area)
        painter.fillRect(event.rect(), QColor(THEME["gutter_bg"]))
        block = self.firstVisibleBlock()
        block_num = block.blockNumber()
        top = round(self.blockBoundingGeometry(block)
                    .translated(self.contentOffset()).top())
        bottom = top + round(self.blockBoundingRect(block).height())
        current_line = self.textCursor().blockNumber()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                num_str = str(block_num + 1)
                if block_num == current_line:
                    painter.setPen(QColor(THEME["accent"]))
                    font = painter.font(); font.setBold(True); painter.setFont(font)
                else:
                    painter.setPen(QColor(THEME["gutter_fg"]))
                    font = painter.font(); font.setBold(False); painter.setFont(font)
                painter.drawText(0, top, self.line_number_area.width() - 4,
                    self.fontMetrics().height(),
                    Qt.AlignmentFlag.AlignRight, num_str)
            block = block.next()
            top = bottom
            bottom = top + round(self.blockBoundingRect(block).height())
            block_num += 1

    def highlight_current_line(self):
        extra = []
        if not self.isReadOnly():
            sel = QTextEdit.ExtraSelection()
            sel.format.setBackground(QColor(THEME["current_line"]))
            sel.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
            sel.cursor = self.textCursor()
            sel.cursor.clearSelection()
            extra.append(sel)
        self.setExtraSelections(extra)

    def _on_cursor_changed(self):
        self.cursorLineChanged.emit(self.textCursor().blockNumber() + 1)

    def keyPressEvent(self, event):
        key = event.key()
        mod = event.modifiers()

        if self._completer_list.isVisible():
            if key == Qt.Key.Key_Down:
                row = self._completer_list.currentRow()
                self._completer_list.setCurrentRow(
                    min(row + 1, self._completer_list.count() - 1))
                return
            elif key == Qt.Key.Key_Up:
                row = self._completer_list.currentRow()
                self._completer_list.setCurrentRow(max(row - 1, 0))
                return
            elif key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                # ★ Enter のみ補完確定。それ以外は補完を閉じて通常処理へ
                item = self._completer_list.currentItem()
                if item:
                    self._insert_completion(item)
                return
            elif key == Qt.Key.Key_Escape:
                self._completer_list.hide()
                return
            else:
                # ★ Tab・文字入力・BackSpace 等は補完を閉じて通常処理に流す
                self._completer_list.hide()
                # fall through ↓

        if key == Qt.Key.Key_Space and mod == Qt.KeyboardModifier.ControlModifier:
            self._show_completion()
            return

        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self._handle_enter()
            return

        if key == Qt.Key.Key_Tab and mod == Qt.KeyboardModifier.NoModifier:
            cursor = self.textCursor()
            if cursor.hasSelection():
                self._indent_selection(cursor)
            else:
                cursor.insertText("    ")
            return

        if key == Qt.Key.Key_Backtab:
            self._unindent_selection(self.textCursor())
            return

        if key == Qt.Key.Key_Greater and self._current_ext in (".html", ".htm", ".xml"):
            super().keyPressEvent(event)
            self._auto_close_tag()
            return

        if key == Qt.Key.Key_QuoteDbl:
            cursor = self.textCursor()
            if not cursor.hasSelection():
                cursor.insertText('""')
                cursor.movePosition(QTextCursor.MoveOperation.Left)
                self.setTextCursor(cursor)
                return

        super().keyPressEvent(event)

        if key not in (Qt.Key.Key_Escape, Qt.Key.Key_Return,
                       Qt.Key.Key_Enter, Qt.Key.Key_Backspace):
            self._update_completion()

    def _handle_enter(self):
        cursor = self.textCursor()
        text   = cursor.block().text()
        indent = re.match(r'^(\s*)', text).group(1)
        pos_in_block = cursor.positionInBlock()
        before = text[:pos_in_block]
        after  = text[pos_in_block:]
        open_tag  = re.search(r'<([a-zA-Z][a-zA-Z0-9\-]*)([^>]*)>$', before)
        close_tag = re.match(r'^</([a-zA-Z][a-zA-Z0-9\-]*)>', after)
        if (open_tag and close_tag and
                open_tag.group(1).lower() == close_tag.group(1).lower()):
            cursor.insertText(f"\n{indent}    \n{indent}")
            cursor.movePosition(QTextCursor.MoveOperation.Up)
            cursor.movePosition(QTextCursor.MoveOperation.EndOfLine)
            self.setTextCursor(cursor)
        else:
            extra = ("    " if open_tag and
                     open_tag.group(1).lower() not in self.VOID_TAGS else "")
            cursor.insertText(f"\n{indent}{extra}")
        self._completer_list.hide()

    def _auto_close_tag(self):
        cursor = self.textCursor()
        before = cursor.block().text()[:cursor.positionInBlock()]
        m = re.search(r'<([a-zA-Z][a-zA-Z0-9\-]*)([^>]*)>$', before)
        if m:
            tag = m.group(1).lower()
            if tag not in self.VOID_TAGS and not m.group(2).endswith('/'):
                cursor.insertText(f"</{tag}>")
                for _ in range(len(tag) + 3):
                    cursor.movePosition(QTextCursor.MoveOperation.Left)
                self.setTextCursor(cursor)

    def _indent_selection(self, cursor):
        start, end = cursor.selectionStart(), cursor.selectionEnd()
        cursor.setPosition(start)
        cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        cursor.movePosition(QTextCursor.MoveOperation.EndOfLine,
                            QTextCursor.MoveMode.KeepAnchor)
        lines = cursor.selectedText().split('\u2029')
        cursor.insertText('\u2029'.join('    ' + l for l in lines))

    def _unindent_selection(self, cursor):
        start, end = cursor.selectionStart(), cursor.selectionEnd()
        cursor.setPosition(start)
        cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        cursor.movePosition(QTextCursor.MoveOperation.EndOfLine,
                            QTextCursor.MoveMode.KeepAnchor)
        lines = cursor.selectedText().split('\u2029')
        new_lines = []
        for l in lines:
            if l.startswith('    '): new_lines.append(l[4:])
            elif l.startswith('\t'): new_lines.append(l[1:])
            else: new_lines.append(l)
        cursor.insertText('\u2029'.join(new_lines))

    def _get_current_word(self):
        cursor = self.textCursor()
        text   = cursor.block().text()
        pos    = cursor.positionInBlock()
        before = text[:pos]
        after  = text[pos:]

        m = re.search(r'<([a-zA-Z][a-zA-Z0-9\-]*)$', before)
        if m:
            # ★ カーソルの直後に英数字・属性・> が続く場合は既存タグ編集中なので除外
            if re.match(r'^[a-zA-Z0-9\-\s>"/=]', after):
                return "", "none"
            return m.group(1), "tag"

        m2 = re.search(r'\s([a-zA-Z\-]*)$', before)
        if m2 and '<' in before:
            return m2.group(1), "attr"

        return "", "none"

    def _update_completion(self):
        if self._current_ext not in (".html", ".htm"):
            self._completer_list.hide()
            return
        word, kind = self._get_current_word()
        # ★ 2文字以上入力された場合のみ補完を表示（1文字は誤爆しやすいため抑制）
        if kind == "tag" and len(word) >= 2:
            candidates = [k for k in self.HTML_SNIPPETS
                          if k.startswith(word) and k != word]
            if candidates:
                self._show_popup(candidates, word, kind)
                return
        self._completer_list.hide()

    def _show_completion(self):
        if self._current_ext not in (".html", ".htm"):
            return
        word, kind = self._get_current_word()
        candidates = ([k for k in self.HTML_SNIPPETS if k.startswith(word)]
                      if kind == "tag" else list(self.HTML_SNIPPETS.keys()))
        if candidates:
            self._show_popup(candidates, word, "tag")

    def _show_popup(self, candidates, word, kind):
        self._completer_list.clear()
        for c in candidates[:15]:
            item = QListWidgetItem(f"<{c}>  ✦ snippet")
            item.setData(Qt.ItemDataRole.UserRole, (c, kind))
            self._completer_list.addItem(item)
        self._completer_list.setCurrentRow(0)
        pos = self.mapToGlobal(self.cursorRect(self.textCursor()).bottomLeft())
        self._completer_list.move(pos)
        self._completer_list.resize(280, min(len(candidates), 10) * 24 + 4)
        self._completer_list.show()
        self._completer_list.raise_()
        self._current_word = word

    def _insert_completion(self, item):
        tag, kind = item.data(Qt.ItemDataRole.UserRole)
        snippet = self.HTML_SNIPPETS.get(tag, f"<{tag}></{tag}>")
        cursor = self.textCursor()
        before = cursor.block().text()[:cursor.positionInBlock()]
        m = re.search(r'<([a-zA-Z][a-zA-Z0-9\-]*)$', before)
        if m:
            for _ in range(len(m.group(0))):
                cursor.deletePreviousChar()
        cursor.insertText(snippet)
        self._completer_list.hide()

    def goto_line(self, line_number):
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        cursor.movePosition(QTextCursor.MoveOperation.Down,
                            QTextCursor.MoveMode.MoveAnchor, line_number - 1)
        self.setTextCursor(cursor)
        self.centerCursor()
        self.setFocus()

    def set_font_size(self, size: int):
        """フォントサイズを変更する"""
        size = max(6, min(32, size))  # 6〜32pt の範囲に制限
        font = self.font()
        font.setPointSize(size)
        self.setFont(font)
        self.setTabStopDistance(
            QFontMetrics(self.font()).horizontalAdvance(' ') * 4
        )
        self.update_line_number_area_width(0)

    def wheelEvent(self, event):
        """Ctrl+ホイールでフォントサイズ変更"""
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                self.set_font_size(self.get_font_size() + 1)
            elif delta < 0:
                self.set_font_size(self.get_font_size() - 1)
            event.accept()
        else:
            super().wheelEvent(event)
            
    def get_font_size(self) -> int:
        return self.font().pointSize()
        

# ============================================================
# 構文チェッカー（拡張版）
# ============================================================
class SyntaxChecker:
    VOID_TAGS = {"area","base","br","col","embed","hr","img","input",
                 "link","meta","param","source","track","wbr"}
    HTML5_TAGS = {
        "header","footer","nav","main","section","article","aside",
        "figure","figcaption","details","summary","mark","time",
        "progress","meter","output","datalist","template","dialog",
        "picture","data"
    }

    def check(self, text: str, ext: str = ".html") -> list[dict]:
        ext = ext.lower()
        if ext in (".html", ".htm"): return self._check_html(text)
        elif ext == ".json":         return self._check_json(text)
        elif ext == ".xml":          return self._check_xml(text)
        elif ext == ".py":           return self._check_python(text)
        elif ext == ".csv":          return self._check_csv(text)
        elif ext == ".css":          return self._check_css(text)
        elif ext == ".js":           return self._check_js(text)
        elif ext == ".bat":          return self._check_batch(text)
        elif ext == ".ps1":          return self._check_powershell(text)
        elif ext == ".vbs":          return self._check_vbs(text)
        elif ext == ".vba":          return self._check_vba(text)
        return []

    # ---- HTML ----
    def _check_html(self, text):
        errors = []
        lines  = text.split('\n')
        try:
            parser = etree.HTMLParser(recover=True)
            etree.fromstring(text.encode('utf-8'), parser)
            for e in parser.error_log:
                if self._is_html5_fp(e.message):
                    continue
                errors.append({"line": e.line, "col": e.column,
                    "type": "error", "message": e.message,
                    "suggestion": self._suggest_html(e.message)})
        except Exception as ex:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": str(ex),
                "suggestion": "HTMLの基本構造を確認してください。"})
        errors += self._custom_html_checks(lines)
        return errors

    def _is_html5_fp(self, msg):
        ml = msg.lower()
        if "tag" in ml and "invalid" in ml:
            for tag in self.HTML5_TAGS:
                if tag in ml:
                    return True
        return False

    def _custom_html_checks(self, lines):
        errors = []
        full = '\n'.join(lines)
        for i, line in enumerate(lines, 1):
            if re.search(r'<img(?![^>]*alt=)[^>]*>', line, re.I):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "<img> タグに alt 属性がありません",
                    "suggestion": 'alt="画像の説明" を追加してください'})
            if re.search(r'<a(?![^>]*href=)[^>]*>', line, re.I):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "<a> タグに href 属性がありません",
                    "suggestion": 'href="#" または適切なURLを追加してください'})
        if not re.search(r'<!DOCTYPE\s+html', full, re.I):
            errors.append({"line": 1, "col": 1, "type": "warning",
                "message": "DOCTYPE 宣言がありません",
                "suggestion": "ファイル先頭に <!DOCTYPE html> を追加してください"})
        if not re.search(r'<meta[^>]+charset', full, re.I):
            errors.append({"line": 1, "col": 1, "type": "info",
                "message": "文字コード宣言がありません",
                "suggestion": '<meta charset="UTF-8"> を <head> 内に追加してください'})
        return errors

    def _suggest_html(self, msg):
        ml = msg.lower()
        if "unexpected end tag" in ml:
            return "対応する開始タグがない閉じタグです。"
        if "attribute" in ml:
            return '属性の記述形式を確認してください。例: attr="value"'
        return "HTMLの構文を確認してください。"

    # ---- JSON ----
    def _check_json(self, text):
        errors = []
        try:
            json.loads(text)
        except json.JSONDecodeError as e:
            errors.append({"line": e.lineno, "col": e.colno,
                "type": "error", "message": f"JSON構文エラー: {e.msg}",
                "suggestion": "カンマの過不足・括弧の対応・クォートを確認してください。"})
        return errors

    # ---- XML ----
    def _check_xml(self, text):
        errors = []
        try:
            etree.fromstring(text.encode('utf-8'))
        except etree.XMLSyntaxError as e:
            errors.append({"line": e.lineno or 1, "col": e.offset or 1,
                "type": "error", "message": f"XML構文エラー: {e.msg}",
                "suggestion": "タグの対応・属性のクォート・エンコーディングを確認してください。"})
        return errors

    # ---- Python ----
    def _check_python(self, text):
        errors = []
        import ast
        try:
            ast.parse(text)
        except SyntaxError as e:
            errors.append({"line": e.lineno or 1, "col": e.offset or 1,
                "type": "error", "message": f"Python構文エラー: {e.msg}",
                "suggestion": "インデント・括弧・コロンの有無を確認してください。"})
        lines = text.split('\n')
        for i, line in enumerate(lines, 1):
            if re.match(r'^\s*print\s+[^(]', line):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "Python2スタイルの print 文です",
                    "suggestion": "print() 関数を使用してください。"})
            if '\t' in line and '    ' in line:
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "タブとスペースが混在しています",
                    "suggestion": "インデントをスペース4つに統一してください。"})
        return errors

    # ---- CSV ----
    def _check_csv(self, text):
        errors = []
        lines = text.split('\n')
        if not lines:
            return errors
        try:
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)
            if not rows:
                return errors
            header_cols = len(rows[0])
            for i, row in enumerate(rows[1:], 2):
                if len(row) != header_cols:
                    errors.append({"line": i, "col": 1, "type": "warning",
                        "message": f"列数が不一致（ヘッダー: {header_cols}列, この行: {len(row)}列）",
                        "suggestion": "カンマの数を確認してください。"})
        except Exception as e:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": f"CSV解析エラー: {e}",
                "suggestion": "CSVの形式を確認してください。"})
        if lines:
            for j, h in enumerate(lines[0].split(','), 1):
                if not h.strip().strip('"'):
                    errors.append({"line": 1, "col": j, "type": "info",
                        "message": f"列{j}のヘッダーが空です",
                        "suggestion": "ヘッダー名を設定してください。"})
        return errors

    # ---- CSS ----
    def _check_css(self, text):
        errors = []
        lines = text.split('\n')
        ob = text.count('{')
        cb = text.count('}')
        if ob != cb:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": f"波括弧の数が不一致（{{: {ob}, }}: {cb}）",
                "suggestion": "{ と } の対応を確認してください。"})
        for i, line in enumerate(lines, 1):
            s = line.strip()
            if ':' in s and not s.endswith((';', '{', '}', ',')):
                if not s.startswith(('/*', '//', '@')):
                    errors.append({"line": i, "col": 1, "type": "warning",
                        "message": "セミコロンが不足している可能性があります",
                        "suggestion": "プロパティ値の末尾に ; を追加してください。"})
        return errors

    # ---- JavaScript ----
    def _check_js(self, text):
        errors = []
        lines = text.split('\n')
        op = text.count('('); cp = text.count(')')
        ob = text.count('{'); cb = text.count('}')
        if op != cp:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": f"括弧 () の数が不一致（(: {op}, ): {cp}）",
                "suggestion": "括弧の対応を確認してください。"})
        if ob != cb:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": f"波括弧 {{}} の数が不一致（{{: {ob}, }}: {cb}）",
                "suggestion": "波括弧の対応を確認してください。"})
        for i, line in enumerate(lines, 1):
            if re.search(r'\bvar\b', line):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "var の使用が検出されました",
                    "suggestion": "const または let の使用を推奨します。"})
            if re.search(r'(?<!=)==(?!=)', line):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "== の使用が検出されました",
                    "suggestion": "=== （厳密等価演算子）の使用を推奨します。"})
        return errors

    # ---- Batch ----
    def _check_batch(self, text):
        errors = []
        lines  = text.split('\n')
        labels = set()
        gotos  = set()
        for i, line in enumerate(lines, 1):
            s = line.strip()
            m = re.match(r'^:([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m:
                labels.add(m.group(1).lower())
            m2 = re.match(r'(?i)goto\s+([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m2:
                gotos.add((m2.group(1).lower(), i))
            if re.match(r'(?i)^exit\b', s):
                prev = [l.strip().lower() for l in lines[max(0,i-3):i-1]]
                if not any('pause' in pl for pl in prev):
                    errors.append({"line": i, "col": 1, "type": "info",
                        "message": "exit の前に pause がありません",
                        "suggestion": "pause を追加することを検討してください。"})
        for label, lineno in gotos:
            if label not in labels and label != 'eof':
                errors.append({"line": lineno, "col": 1, "type": "warning",
                    "message": f"goto 先のラベル :{label} が見つかりません",
                    "suggestion": f":{label} ラベルを定義してください。"})
        return errors

    # ---- PowerShell ----
    def _check_powershell(self, text):
        errors = []
        lines  = text.split('\n')
        ob = text.count('{'); cb = text.count('}')
        if ob != cb:
            errors.append({"line": 1, "col": 1, "type": "error",
                "message": f"波括弧の数が不一致（{{: {ob}, }}: {cb}）",
                "suggestion": "波括弧の対応を確認してください。"})
        for i, line in enumerate(lines, 1):
            if re.search(r'(?i)Invoke-Expression|iex\b', line):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "Invoke-Expression (iex) の使用が検出されました",
                    "suggestion": "セキュリティリスクがあります。使用目的を確認してください。"})
            if re.search(r'(?i)Write-Host\b', line):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "Write-Host の使用が検出されました",
                    "suggestion": "パイプライン処理には Write-Output の使用を検討してください。"})
        return errors

    # ---- VBScript ----
    def _check_vbs(self, text):
        errors = []
        lines  = text.split('\n')
        for i, line in enumerate(lines, 1):
            s = line.strip()
            if re.match(r'(?i)on\s+error\s+resume\s+next', s):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "On Error Resume Next が使用されています",
                    "suggestion": "エラーが隠蔽されます。適切なエラーハンドリングを検討してください。"})
            if re.match(r'(?i)^\s*Sub\s+', s):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "Sub プロシージャが定義されています",
                    "suggestion": "対応する End Sub があることを確認してください。"})
        return errors

    # ---- VBA ----
    def _check_vba(self, text):
        errors = []
        lines  = text.split('\n')
        full   = text

        # Option Explicit チェック
        if not re.search(r'(?i)^\s*Option\s+Explicit', full, re.MULTILINE):
            errors.append({"line": 1, "col": 1, "type": "warning",
                "message": "Option Explicit が宣言されていません",
                "suggestion": "モジュール先頭に 'Option Explicit' を追加すると変数の未宣言を防げます。"})

        # Sub/Function の End 対応チェック
        sub_stack   = []
        func_stack  = []
        for i, line in enumerate(lines, 1):
            s = line.strip()
            # Sub 開始
            m = re.match(r'(?i)^(Public\s+|Private\s+|Friend\s+|Static\s+)?'
                         r'Sub\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\(', s)
            if m and not re.match(r'(?i)^End\s+Sub', s):
                sub_stack.append((m.group(2), i))
            # End Sub
            if re.match(r'(?i)^End\s+Sub\b', s):
                if sub_stack:
                    sub_stack.pop()
                else:
                    errors.append({"line": i, "col": 1, "type": "error",
                        "message": "対応する Sub がない End Sub です",
                        "suggestion": "Sub の定義を確認してください。"})
            # Function 開始
            m2 = re.match(r'(?i)^(Public\s+|Private\s+|Friend\s+|Static\s+)?'
                          r'Function\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\(', s)
            if m2 and not re.match(r'(?i)^End\s+Function', s):
                func_stack.append((m2.group(2), i))
            # End Function
            if re.match(r'(?i)^End\s+Function\b', s):
                if func_stack:
                    func_stack.pop()
                else:
                    errors.append({"line": i, "col": 1, "type": "error",
                        "message": "対応する Function がない End Function です",
                        "suggestion": "Function の定義を確認してください。"})

            # On Error Resume Next 警告
            if re.match(r'(?i)on\s+error\s+resume\s+next', s):
                errors.append({"line": i, "col": 1, "type": "warning",
                    "message": "On Error Resume Next が使用されています",
                    "suggestion": "エラーが隠蔽されます。On Error GoTo エラーハンドラ の使用を検討してください。"})

            # Select Case の End Select チェック（簡易）
            if re.match(r'(?i)^Select\s+Case\b', s):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "Select Case ブロックが開始されています",
                    "suggestion": "対応する End Select があることを確認してください。"})

            # MsgBox の過多使用（デバッグ残り）
            if re.search(r'(?i)\bMsgBox\b', s):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "MsgBox が使用されています",
                    "suggestion": "デバッグ用の MsgBox が残っていないか確認してください。"})

            # ActiveSheet / ActiveWorkbook の直接使用警告
            if re.search(r'(?i)\bActiveSheet\b|\bActiveWorkbook\b', s):
                errors.append({"line": i, "col": 1, "type": "info",
                    "message": "ActiveSheet / ActiveWorkbook の使用が検出されました",
                    "suggestion": "明示的に Worksheets(\"シート名\") や ThisWorkbook を使用することを推奨します。"})

        # 閉じられていない Sub
        for name, lineno in sub_stack:
            errors.append({"line": lineno, "col": 1, "type": "error",
                "message": f"Sub '{name}' に対応する End Sub がありません",
                "suggestion": "End Sub を追加してください。"})
        # 閉じられていない Function
        for name, lineno in func_stack:
            errors.append({"line": lineno, "col": 1, "type": "error",
                "message": f"Function '{name}' に対応する End Function がありません",
                "suggestion": "End Function を追加してください。"})

        return errors


# ============================================================
# セクションナビゲーター（拡張版）
# ============================================================
class SectionNavigator(QTreeWidget):
    sectionSelected = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderLabel("📁 ドキュメント構造")
        self.setStyleSheet(f"""
            QTreeWidget {{
                background: {THEME['bg_panel']};
                color: {THEME['fg_text']};
                border: none;
                font-size: 12px;
            }}
            QTreeWidget::item:selected {{
                background: {THEME['accent']};
                color: {THEME['bg_editor']};
            }}
            QTreeWidget::item:hover {{ background: {THEME['bg_line']}; }}
            QHeaderView::section {{
                background: {THEME['bg_panel']};
                color: {THEME['accent']};
                border: none;
                padding: 4px;
                font-weight: bold;
            }}
        """)
        self.itemClicked.connect(self._on_item_clicked)

    def update_structure(self, text: str, ext: str = ".html"):
        self.clear()
        ext = ext.lower()
        dispatch = {
            ".html": self._build_html, ".htm": self._build_html,
            ".py":   self._build_python,
            ".js":   self._build_js,
            ".json": self._build_json,
            ".xml":  self._build_xml,
            ".css":  self._build_css,
            ".bat":  self._build_batch,
            ".ps1":  self._build_powershell,
            ".vbs":  self._build_vbs,
            ".vba":  self._build_vba,
            ".csv":  self._build_csv,
        }
        fn = dispatch.get(ext)
        if fn:
            fn(text)

    def _add_item(self, parent, label, line, color=None):
        item = QTreeWidgetItem(parent, [label])
        item.setData(0, Qt.ItemDataRole.UserRole, line)
        if color:
            item.setForeground(0, QColor(color))
        return item

    def _build_html(self, text):
        lines = text.split('\n')
        sections = {
            "head":     QTreeWidgetItem(self, ["🔧 <head>"]),
            "style":    QTreeWidgetItem(self, ["🎨 <style>"]),
            "script":   QTreeWidgetItem(self, ["⚡ <script>"]),
            "body":     QTreeWidgetItem(self, ["📄 <body>"]),
            "headings": QTreeWidgetItem(self, ["📝 見出し"]),
            "links":    QTreeWidgetItem(self, ["🔗 リンク"]),
            "images":   QTreeWidgetItem(self, ["🖼 画像"]),
            "forms":    QTreeWidgetItem(self, ["📋 フォーム"]),
        }
        for s in sections.values():
            s.setExpanded(True)
        for i, line in enumerate(lines, 1):
            s = line.strip()
            if re.match(r'<head[\s>]', s, re.I):
                sections["head"].setData(0, Qt.ItemDataRole.UserRole, i)
            if re.match(r'<style[\s>]', s, re.I):
                self._add_item(sections["style"], f"  Line {i}: <style>", i, THEME["fg_css_prop"])
            if re.match(r'<script[\s>]', s, re.I):
                src = re.search(r'src=["\']([^"\']+)["\']', s)
                self._add_item(sections["script"],
                    f"  Line {i}: " + (f"src={src.group(1)}" if src else "<script>"),
                    i, THEME["fg_js_keyword"])
            if re.match(r'<body[\s>]', s, re.I):
                sections["body"].setData(0, Qt.ItemDataRole.UserRole, i)
            m = re.match(r'<(h[1-6])[\s>]', s, re.I)
            if m:
                content = re.sub(r'<[^>]+>', '', s)[:40]
                self._add_item(sections["headings"],
                    f"  {m.group(1).upper()}: {content}", i, THEME["fg_tag"])
            if re.match(r'<a\s', s, re.I):
                href = re.search(r'href=["\']([^"\']+)["\']', s)
                self._add_item(sections["links"],
                    f"  Line {i}: " + (href.group(1)[:40] if href else "(no href)"),
                    i, THEME["accent2"])
            if re.match(r'<img[\s>]', s, re.I):
                src = re.search(r'src=["\']([^"\']+)["\']', s)
                self._add_item(sections["images"],
                    f"  Line {i}: " + (src.group(1)[:40] if src else "(no src)"), i)
            if re.match(r'<form[\s>]', s, re.I):
                self._add_item(sections["forms"], f"  Line {i}: <form>", i)
        for item in sections.values():
            if item.childCount() == 0 and item.data(0, Qt.ItemDataRole.UserRole) is None:
                item.setHidden(True)

    def _build_python(self, text):
        classes   = QTreeWidgetItem(self, ["🏛 クラス"])
        functions = QTreeWidgetItem(self, ["🔧 関数"])
        imports   = QTreeWidgetItem(self, ["📦 インポート"])
        for w in [classes, functions, imports]:
            w.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            s = line.strip()
            m = re.match(r'class\s+([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m:
                self._add_item(classes, f"  Line {i}: {m.group(1)}", i, THEME["fg_tag"])
            m = re.match(r'(?:async\s+)?def\s+([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m:
                self._add_item(functions, f"  Line {i}: {m.group(1)}()", i, THEME["fg_js_func"])
            if re.match(r'(?:import|from)\s+', s):
                self._add_item(imports, f"  Line {i}: {s[:50]}", i, THEME["fg_css_prop"])
        for w in [classes, functions, imports]:
            if w.childCount() == 0:
                w.setHidden(True)

    def _build_js(self, text):
        functions = QTreeWidgetItem(self, ["🔧 関数"])
        classes   = QTreeWidgetItem(self, ["🏛 クラス"])
        imports   = QTreeWidgetItem(self, ["📦 インポート"])
        for w in [functions, classes, imports]:
            w.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            s = line.strip()
            m = re.match(r'(?:async\s+)?function\s+([a-zA-Z_$][a-zA-Z0-9_$]*)', s)
            if m:
                self._add_item(functions, f"  Line {i}: {m.group(1)}()", i, THEME["fg_js_func"])
            m2 = re.match(r'(?:const|let|var)\s+([a-zA-Z_$][a-zA-Z0-9_$]*)\s*=\s*(?:async\s+)?(?:function|\()', s)
            if m2:
                self._add_item(functions, f"  Line {i}: {m2.group(1)}()", i, THEME["fg_js_func"])
            m3 = re.match(r'class\s+([a-zA-Z_$][a-zA-Z0-9_$]*)', s)
            if m3:
                self._add_item(classes, f"  Line {i}: {m3.group(1)}", i, THEME["fg_tag"])
            if re.match(r'(?:import|export)\s+', s):
                self._add_item(imports, f"  Line {i}: {s[:50]}", i, THEME["fg_css_prop"])
        for w in [functions, classes, imports]:
            if w.childCount() == 0:
                w.setHidden(True)

    def _build_json(self, text):
        root = QTreeWidgetItem(self, ["🔑 トップレベルキー"])
        root.setExpanded(True)
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                for key in list(data.keys())[:30]:
                    vtype = type(data[key]).__name__
                    self._add_item(root, f"  {key}  [{vtype}]", 1, THEME["fg_attr"])
            elif isinstance(data, list):
                QTreeWidgetItem(self, [f"📋 配列  ({len(data)} 件)"])
        except Exception:
            err = QTreeWidgetItem(self, ["❌ JSON解析エラー"])
            err.setForeground(0, QColor(THEME["error"]))

    def _build_xml(self, text):
        elements = QTreeWidgetItem(self, ["🏷 要素"])
        elements.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            m = re.match(r'\s*<([a-zA-Z][a-zA-Z0-9:_\-]*)[\s>]', line)
            if m and not line.strip().startswith('</') and not line.strip().startswith('<?'):
                self._add_item(elements, f"  Line {i}: <{m.group(1)}>", i, THEME["fg_tag"])
                if elements.childCount() >= 50:
                    break

    def _build_css(self, text):
        selectors = QTreeWidgetItem(self, ["🎨 セレクタ"])
        media     = QTreeWidgetItem(self, ["📱 メディアクエリ"])
        for w in [selectors, media]:
            w.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            s = line.strip()
            if re.match(r'@media', s, re.I):
                self._add_item(media, f"  Line {i}: {s[:50]}", i, THEME["fg_doctype"])
            elif re.search(r'\{', s) and not s.startswith('@'):
                sel = s.replace('{', '').strip()[:40]
                self._add_item(selectors, f"  Line {i}: {sel}", i, THEME["fg_tag"])
        for w in [selectors, media]:
            if w.childCount() == 0:
                w.setHidden(True)

    def _build_batch(self, text):
        labels   = QTreeWidgetItem(self, ["🏷 ラベル"])
        commands = QTreeWidgetItem(self, ["⚡ 主要コマンド"])
        for w in [labels, commands]:
            w.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            s = line.strip()
            m = re.match(r'^:([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m:
                self._add_item(labels, f"  Line {i}: :{m.group(1)}", i, THEME["fg_doctype"])
            if re.match(r'(?i)^(call|goto|if|for)\b', s):
                self._add_item(commands, f"  Line {i}: {s[:50]}", i, THEME["fg_js_keyword"])
        for w in [labels, commands]:
            if w.childCount() == 0:
                w.setHidden(True)

    def _build_powershell(self, text):
        functions = QTreeWidgetItem(self, ["🔧 関数"])
        functions.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            m = re.match(r'(?i)function\s+([a-zA-Z][a-zA-Z0-9_\-]*)', line.strip())
            if m:
                self._add_item(functions, f"  Line {i}: {m.group(1)}", i, THEME["fg_js_func"])
        if functions.childCount() == 0:
            functions.setHidden(True)

    def _build_vbs(self, text):
        subs = QTreeWidgetItem(self, ["🔧 Sub / Function"])
        subs.setExpanded(True)
        for i, line in enumerate(text.split('\n'), 1):
            m = re.match(r'(?i)(Sub|Function)\s+([a-zA-Z_][a-zA-Z0-9_]*)', line.strip())
            if m:
                self._add_item(subs, f"  Line {i}: {m.group(1)} {m.group(2)}",
                               i, THEME["fg_js_func"])
        if subs.childCount() == 0:
            subs.setHidden(True)

    def _build_vba(self, text):
        """VBA専用ナビゲーター：Sub/Function/Property/定数/モジュール変数"""
        subs      = QTreeWidgetItem(self, ["🔧 Sub プロシージャ"])
        functions = QTreeWidgetItem(self, ["⚙ Function プロシージャ"])
        props     = QTreeWidgetItem(self, ["📌 Property"])
        consts    = QTreeWidgetItem(self, ["🔒 定数 (Const)"])
        mod_vars  = QTreeWidgetItem(self, ["📦 モジュール変数"])
        for w in [subs, functions, props, consts, mod_vars]:
            w.setExpanded(True)

        for i, line in enumerate(text.split('\n'), 1):
            s = line.strip()

            # Sub
            m = re.match(
                r'(?i)^(Public\s+|Private\s+|Friend\s+|Static\s+)?'
                r'Sub\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\(', s)
            if m:
                scope = (m.group(1) or "").strip()
                name  = m.group(2)
                label = f"  Line {i}: {name}()"
                if scope:
                    label += f"  [{scope}]"
                self._add_item(subs, label, i, THEME["fg_js_func"])

            # Function
            m2 = re.match(
                r'(?i)^(Public\s+|Private\s+|Friend\s+|Static\s+)?'
                r'Function\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\(', s)
            if m2:
                scope = (m2.group(1) or "").strip()
                name  = m2.group(2)
                label = f"  Line {i}: {name}()"
                if scope:
                    label += f"  [{scope}]"
                self._add_item(functions, label, i, THEME["fg_css_prop"])

            # Property
            m3 = re.match(
                r'(?i)^(Public\s+|Private\s+)?'
                r'Property\s+(Get|Let|Set)\s+([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m3:
                self._add_item(props,
                    f"  Line {i}: Property {m3.group(2)} {m3.group(3)}",
                    i, THEME["fg_attr"])

            # Const
            m4 = re.match(
                r'(?i)^(Public\s+|Private\s+)?'
                r'Const\s+([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m4:
                self._add_item(consts,
                    f"  Line {i}: {m4.group(2)}", i, THEME["fg_entity"])

            # モジュールレベル変数（Dim/Public/Private がインデントなし）
            m5 = re.match(
                r'(?i)^(Public|Private|Friend)\s+'
                r'(?!Sub|Function|Property|Const)'
                r'([a-zA-Z_][a-zA-Z0-9_]*)', s)
            if m5:
                self._add_item(mod_vars,
                    f"  Line {i}: {m5.group(2)}  [{m5.group(1)}]",
                    i, THEME["fg_string"])

        for w in [subs, functions, props, consts, mod_vars]:
            if w.childCount() == 0:
                w.setHidden(True)

    def _build_csv(self, text):
        lines   = text.split('\n')
        headers = QTreeWidgetItem(self, ["📋 ヘッダー列"])
        stats   = QTreeWidgetItem(self, ["📊 統計"])
        for w in [headers, stats]:
            w.setExpanded(True)
        if lines:
            cols = lines[0].split(',')
            for j, col in enumerate(cols, 1):
                self._add_item(headers,
                    f"  列{j}: {col.strip().strip(chr(34))}", 1, THEME["fg_attr"])
            data_rows = len([l for l in lines[1:] if l.strip()])
            QTreeWidgetItem(stats, [f"  データ行数: {data_rows}"])
            QTreeWidgetItem(stats, [f"  列数: {len(cols)}"])

    def _on_item_clicked(self, item, col):
        line = item.data(0, Qt.ItemDataRole.UserRole)
        if line:
            self.sectionSelected.emit(line)


# ============================================================
# 疎結合化ダイアログ
# ============================================================
class DecoupleDialog(QDialog):
    def __init__(self, parent, has_style, has_script):
        super().__init__(parent)
        self.setWindowTitle("🔗 外部ファイル化（疎結合化）")
        self.setStyleSheet(f"""
            QDialog {{ background: {THEME['bg_panel']}; color: {THEME['fg_text']}; }}
            QLabel {{ color: {THEME['fg_text']}; }}
            QCheckBox {{ color: {THEME['fg_text']}; }}
            QLineEdit {{
                background: {THEME['bg_editor']}; color: {THEME['fg_text']};
                border: 1px solid {THEME['border']}; padding: 4px;
            }}
            QPushButton {{
                background: {THEME['accent']}; color: {THEME['bg_editor']};
                border: none; padding: 6px 16px; font-weight: bold;
            }}
            QPushButton:hover {{ background: {THEME['accent2']}; }}
        """)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("選択した要素を外部ファイルとして切り出します。"))
        self.cb_style = QCheckBox("🎨 <style> → .css ファイルに切り出す")
        self.cb_style.setChecked(has_style)
        self.cb_style.setEnabled(has_style)
        layout.addWidget(self.cb_style)
        self.style_name = QLineEdit("style.css")
        layout.addWidget(self.style_name)
        self.cb_script = QCheckBox("⚡ <script> → .js ファイルに切り出す")
        self.cb_script.setChecked(has_script)
        self.cb_script.setEnabled(has_script)
        layout.addWidget(self.cb_script)
        self.script_name = QLineEdit("script.js")
        layout.addWidget(self.script_name)
        btn = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addWidget(btn)


# ============================================================
# PreviewPage
# ============================================================
class PreviewPage(QWebEnginePage):
    elementClicked      = pyqtSignal(str, int, int)
    elementRightClicked = pyqtSignal(str, str, str, str, str)  # ★ str を1つ追加

    def __init__(self, parent=None):
        super().__init__(parent)

    def javaScriptConsoleMessage(self, level, message, line, source):
        if message.startswith("CLICK:"):
            try:
                data = json.loads(message[6:])
                self.elementClicked.emit(
                    data.get("tag",""), data.get("line",1), data.get("col",1))
            except Exception: pass
        elif message.startswith("RIGHTCLICK:"):
            try:
                data = json.loads(message[11:])
                self.elementRightClicked.emit(
                    data.get("tag",   ""),
                    data.get("id",    ""),
                    data.get("cls",   ""),
                    data.get("text",  ""),
                    data.get("outer", "")   # ★ outer を追加
                )
            except Exception: pass


# ============================================================
# PreviewWindow
# ============================================================
class PreviewWindow(QMainWindow):
    closed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🌐 プレビュー")
        self.resize(900, 700)
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0,0,0,0)
        layout.setSpacing(0)

        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(4,2,4,2)
        toolbar.setSpacing(4)
        self.url_bar = QLabel("プレビュー")
        self.url_bar.setStyleSheet(f"color:{THEME['fg_comment']};font-size:10px;")
        toolbar.addWidget(self.url_bar)
        toolbar.addStretch()

        for label, width, tip in [
            ("🖥",0,"PC表示"),("📱",390,"スマホ表示"),("📟",768,"タブレット表示")]:
            btn = QPushButton(label)
            btn.setFixedSize(24,22)
            btn.setToolTip(tip)
            btn.clicked.connect(lambda _, w=width: self._set_viewport(w))
            btn.setStyleSheet(
                f"QPushButton{{background:{THEME['bg_line']};color:{THEME['fg_text']};"
                f"border:none;font-size:12px;border-radius:3px;}}"
                f"QPushButton:hover{{background:{THEME['accent']};color:{THEME['bg_editor']};}}")
            toolbar.addWidget(btn)

        self.btn_refresh = QPushButton("🔄")
        self.btn_refresh.setFixedSize(24,22)
        self.btn_refresh.setStyleSheet(
            f"QPushButton{{background:{THEME['bg_line']};color:{THEME['fg_text']};"
            f"border:none;font-size:12px;border-radius:3px;}}"
            f"QPushButton:hover{{background:{THEME['accent']};color:{THEME['bg_editor']};}}")
        toolbar.addWidget(self.btn_refresh)

        tw = QWidget()
        tw.setLayout(toolbar)
        tw.setFixedHeight(28)
        tw.setStyleSheet(
            f"background:{THEME['bg_panel']};"
            f"border-bottom:1px solid {THEME['border']};")
        layout.addWidget(tw)

        self.view = QWebEngineView()
        self.page = PreviewPage(self.view)
        self.view.setPage(self.page)
        layout.addWidget(self.view)

        self.view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._show_context_menu)

        self.status = QStatusBar()
        self.status.setStyleSheet(
            f"background:{THEME['bg_panel']};color:{THEME['fg_comment']};font-size:11px;")
        self.setStatusBar(self.status)

    def _set_viewport(self, width):
        if width == 0:
            self.resize(900,700); self.status.showMessage("PC表示",2000)
        else:
            self.resize(width+40, self.height())
            self.status.showMessage(f"幅{width}px",2000)

    def _show_context_menu(self, pos):
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{THEME['bg_panel']};color:{THEME['fg_text']};"
            f"border:1px solid {THEME['border']};padding:4px 0;font-size:13px;}}"
            f"QMenu::item{{padding:6px 24px 6px 12px;}}"
            f"QMenu::item:selected{{background:{THEME['accent']};color:{THEME['bg_editor']};}}"
            f"QMenu::separator{{height:1px;background:{THEME['border']};margin:4px 0;}}")
        ab = QAction("◀  Back",    self); ab.setEnabled(self.view.history().canGoBack());    ab.triggered.connect(self.view.back);    menu.addAction(ab)
        af = QAction("▶  Forward", self); af.setEnabled(self.view.history().canGoForward()); af.triggered.connect(self.view.forward); menu.addAction(af)
        ar = QAction("🔄  Reload", self); ar.triggered.connect(self.view.reload);            menu.addAction(ar)
        menu.addSeparator()
        aso = QAction("🔍  View Page Source", self)
        aso.triggered.connect(self._request_jump)
        menu.addAction(aso)
        menu.exec(self.view.mapToGlobal(pos))

    def _request_jump(self):
        pw = self.parent()
        if pw and hasattr(pw, '_jump_to_source'):
            pw._jump_to_source()

    def closeEvent(self, event):
        self.closed.emit()
        super().closeEvent(event)


# ============================================================
# PreviewWidget
# ============================================================
class PreviewWidget(QWidget):
    elementClicked = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._popup: Optional[PreviewWindow] = None
        self._last_html = ""
        self._last_base = ""
        self._last_right_clicked_tag  = ""
        self._last_right_clicked_id   = ""
        self._last_right_clicked_cls  = ""
        self._last_right_clicked_text = ""
        self._last_right_clicked_outer = ""   # ★ 追加

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
        layout.setSpacing(0)

        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(4,2,4,2)
        toolbar.setSpacing(4)
        self.url_bar = QLabel("プレビュー")
        self.url_bar.setStyleSheet(f"color:{THEME['fg_comment']};font-size:10px;")
        toolbar.addWidget(self.url_bar)
        toolbar.addStretch()

        self.btn_popout = QPushButton("🗗 別ウィンドウ")
        self.btn_popout.setFixedWidth(110)          # 幅を文字数に合わせて固定
        self.btn_popout.setFixedHeight(22)
        self.btn_popout.setToolTip("別ウィンドウで開く / 埋め込みに戻す")
        self.btn_popout.clicked.connect(self._toggle_popup)
        self.btn_popout.setStyleSheet(
            f"QPushButton{{background:{THEME['accent']};color:{THEME['bg_editor']};"
            f"border:none;font-size:12px;border-radius:3px;}}"
            f"QPushButton:hover{{background:{THEME['accent2']};}}")
        toolbar.addWidget(self.btn_popout)

        self.btn_refresh = QPushButton("🔄")
        self.btn_refresh.setFixedSize(24,22)
        self.btn_refresh.setStyleSheet(
            f"QPushButton{{background:{THEME['bg_line']};color:{THEME['fg_text']};"
            f"border:none;font-size:12px;border-radius:3px;}}"
            f"QPushButton:hover{{background:{THEME['accent']};color:{THEME['bg_editor']};}}")
        toolbar.addWidget(self.btn_refresh)

        tw = QWidget()
        tw.setLayout(toolbar)
        tw.setFixedHeight(28)
        tw.setStyleSheet(
            f"background:{THEME['bg_panel']};"
            f"border-bottom:1px solid {THEME['border']};")
        layout.addWidget(tw)

        self.view = QWebEngineView()
        self.page = PreviewPage(self.view)
        self.view.setPage(self.page)
        layout.addWidget(self.view)

        self.view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.view.customContextMenuRequested.connect(self._show_context_menu)
        self.page.elementRightClicked.connect(self._on_right_click_element)

        self.placeholder = QLabel(
            "🗗 プレビューは別ウィンドウで表示中\n\nウィンドウを閉じると\nここに戻ります")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet(
            f"QLabel{{color:{THEME['fg_comment']};font-size:14px;"
            f"background:{THEME['bg_panel']};border:2px dashed {THEME['border']};"
            f"border-radius:8px;padding:40px;}}")
        self.placeholder.hide()
        layout.addWidget(self.placeholder)

        self._inject_js = """
        <script>
        document.addEventListener('DOMContentLoaded', function() {
            document.body.addEventListener('click', function(e) {
                var el = e.target;
                var info = {
                    tag: el.tagName ? el.tagName.toLowerCase() : '',
                    id: el.id || '',
                    cls: (typeof el.className === 'string') ? el.className : '',
                    text: (el.innerText || '').substring(0, 80),
                    line: 1, col: 1
                };
                console.log('CLICK:' + JSON.stringify(info));
                document.querySelectorAll('.__hme_selected').forEach(function(x){
                    x.style.outline = ''; x.classList.remove('__hme_selected');
                });
                el.style.outline = '2px solid #89B4FA';
                el.classList.add('__hme_selected');
                e.stopPropagation();
            }, true);

            // ★ 右クリック：outerHTML も送信
            document.body.addEventListener('contextmenu', function(e) {
                var el = e.target;
                var outer = (el.outerHTML || '').substring(0, 200);
                var info = {
                    tag:   el.tagName ? el.tagName.toLowerCase() : '',
                    id:    el.id || '',
                    cls:   (typeof el.className === 'string') ? el.className : '',
                    text:  (el.innerText || '').substring(0, 80),
                    outer: outer
                };
                console.log('RIGHTCLICK:' + JSON.stringify(info));
            }, true);
        });
        </script>
        """

    def _on_right_click_element(self, tag, eid, cls, text, outer=""):
        self._last_right_clicked_tag   = tag
        self._last_right_clicked_id    = eid
        self._last_right_clicked_cls   = cls
        self._last_right_clicked_text  = text
        self._last_right_clicked_outer = outer

    def _toggle_popup(self):
        if self._popup is None: self._open_popup()
        else: self._close_popup()

    def _open_popup(self):
        self._popup = PreviewWindow(self)
        self._popup.btn_refresh.clicked.connect(
            lambda: self.load_html(self._last_html, self._last_base))
        # ★ 修正後（outer引数が増えたので明示的にラムダで受ける）
        self._popup.page.elementRightClicked.connect(
            self._on_right_click_element
        )
        self._popup.closed.connect(self._on_popup_closed)
        self.view.hide()
        self.placeholder.show()
        self.btn_popout.setText("🗗 埋め込みに戻す")
        self.btn_popout.setFixedWidth(130)   # ★ 追加：長いテキストに合わせて幅を広げる
        self.btn_popout.setStyleSheet(
            f"QPushButton{{background:{THEME['warning']};color:{THEME['bg_editor']};"
            f"border:none;padding:3px 8px;font-size:11px;border-radius:3px;font-weight:bold;}}"
            f"QPushButton:hover{{background:{THEME['error']};}}")
        if self._last_html:
            self._load_to_popup(self._last_html, self._last_base)
        self._popup.show()

    def _close_popup(self):
        if self._popup:
            self._popup.closed.disconnect(self._on_popup_closed)
            self._popup.close()
            self._popup = None
        self._on_popup_closed()

    def _on_popup_closed(self):
        self._popup = None
        self.placeholder.hide()
        self.view.show()
        self.btn_popout.setText("🗗 別ウィンドウ")
        self.btn_popout.setFixedWidth(110)   # ★ 追加：元の幅に戻す
        self.btn_popout.setStyleSheet(
            f"QPushButton{{background:{THEME['accent']};color:{THEME['bg_editor']};"
            f"border:none;padding:3px 8px;font-size:11px;border-radius:3px;font-weight:bold;}}"
            f"QPushButton:hover{{background:{THEME['accent2']};}}")
        if self._last_html:
            self._load_to_view(self.view, self.page, self._last_html, self._last_base)

    def _inject_html(self, html):
        if '</head>' in html:
            return html.replace('</head>', self._inject_js + '</head>', 1)
        elif '<body' in html:
            return re.sub(r'(<body[^>]*>)', r'\1' + self._inject_js, html, 1)
        return self._inject_js + html

    def _load_to_view(self, view, page, html, base_url):
        injected = self._inject_html(html)
        if base_url: view.setHtml(injected, QUrl.fromLocalFile(base_url))
        else:        view.setHtml(injected)

    def _load_to_popup(self, html, base_url):
        if self._popup:
            self._load_to_view(self._popup.view, self._popup.page, html, base_url)

    def load_html(self, html, base_url=""):
        self._last_html = html
        self._last_base = base_url
        if self._popup: self._load_to_popup(html, base_url)
        else:           self._load_to_view(self.view, self.page, html, base_url)

    def load_text(self, text, lang=""):
        """HTML以外のファイルをプレビュー表示"""
        escaped = (text.replace('&','&amp;').replace('<','&lt;')
                       .replace('>','&gt;').replace('"','&quot;'))
        html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
body{{background:#1E1E2E;color:#CDD6F4;
     font-family:'Consolas',monospace;font-size:13px;
     margin:0;padding:16px;white-space:pre-wrap;word-break:break-all;}}
.lang-badge{{position:fixed;top:8px;right:12px;background:#313244;
             color:#89B4FA;padding:2px 10px;border-radius:4px;font-size:11px;}}
</style></head>
<body><div class="lang-badge">{lang}</div>{escaped}</body></html>"""
        self._last_html = html
        self._last_base = ""
        if self._popup: self._load_to_popup(html, "")
        else:           self._load_to_view(self.view, self.page, html, "")

    def _show_context_menu(self, pos):
        menu = QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{THEME['bg_panel']};color:{THEME['fg_text']};"
            f"border:1px solid {THEME['border']};padding:4px 0;font-size:13px;}}"
            f"QMenu::item{{padding:6px 24px 6px 12px;}}"
            f"QMenu::item:selected{{background:{THEME['accent']};color:{THEME['bg_editor']};}}"
            f"QMenu::item:disabled{{color:{THEME['fg_comment']};}}"
            f"QMenu::separator{{height:1px;background:{THEME['border']};margin:4px 0;}}")
        ab  = QAction("◀  Back",    self); ab.setEnabled(self.view.history().canGoBack());    ab.triggered.connect(self.view.back);    menu.addAction(ab)
        af  = QAction("▶  Forward", self); af.setEnabled(self.view.history().canGoForward()); af.triggered.connect(self.view.forward); menu.addAction(af)
        ar  = QAction("🔄  Reload", self); ar.triggered.connect(lambda: self.load_html(self._last_html, self._last_base)); menu.addAction(ar)
        menu.addSeparator()
        asv = QAction("💾  Save Page",       self); asv.triggered.connect(self._save_page);      menu.addAction(asv)
        aso = QAction("🔍  View Page Source", self); aso.triggered.connect(self._jump_to_source); menu.addAction(aso)
        menu.exec(self.view.mapToGlobal(pos))

    def _save_page(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "ページを保存", "page.html",
            "HTML Files (*.html *.htm);;All Files (*)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(self._last_html)
                parent = self.window()
                if hasattr(parent, 'status'):
                    parent.status.showMessage(f"保存しました: {path}", 3000)
            except Exception as e:
                QMessageBox.warning(self, "保存エラー", str(e))

    def _jump_to_source(self):
        """右クリックした要素をスコアリングで精度よくエディタジャンプ"""
        tag   = self._last_right_clicked_tag
        eid   = self._last_right_clicked_id
        cls   = self._last_right_clicked_cls
        text  = self._last_right_clicked_text
        outer = getattr(self, '_last_right_clicked_outer', '')

        html = self._last_html
        if not html or not tag:
            return

        lines = html.split('\n')

        # ---- outerHTML から属性を抽出（最高精度） ----
        outer_attrs = {}
        if outer:
            m_id   = re.search(r'\bid=["\']([^"\']+)["\']',   outer)
            m_cls  = re.search(r'\bclass=["\']([^"\']+)["\']', outer)
            m_src  = re.search(r'\bsrc=["\']([^"\']+)["\']',  outer)
            m_href = re.search(r'\bhref=["\']([^"\']+)["\']', outer)
            m_name = re.search(r'\bname=["\']([^"\']+)["\']', outer)
            if m_id:   outer_attrs['id']   = m_id.group(1)
            if m_cls:  outer_attrs['cls']  = m_cls.group(1).split()[0]
            if m_src:  outer_attrs['src']  = m_src.group(1)
            if m_href: outer_attrs['href'] = m_href.group(1)
            if m_name: outer_attrs['name'] = m_name.group(1)

        # ---- 全行をスコアリング ----
        def score_line(i, line):
            s  = 0
            ll = line.lower()

            # タグ名が含まれていない行は対象外
            if f'<{tag}' not in ll:
                return -1

            # id 完全一致（最高点）
            target_id = outer_attrs.get('id') or eid
            if target_id:
                if (f'id="{target_id}"' in ll or
                        f"id='{target_id}'" in ll):
                    s += 100

            # class 一致
            target_cls = outer_attrs.get('cls') or (
                cls.strip().split()[0] if cls.strip() else "")
            if target_cls and target_cls.lower() in ll:
                s += 50

            # src / href / name 一致
            for attr in ('src', 'href', 'name'):
                val = outer_attrs.get(attr, '')
                if val and val.lower() in ll:
                    s += 40

            # テキスト内容一致（同一行 or 前後2行）
            if text:
                short = text.strip()[:30].lower()
                if short and short in ll:
                    s += 30
                else:
                    context = '\n'.join(
                        lines[max(0, i-2): min(len(lines), i+2)]
                    ).lower()
                    if short and short in context:
                        s += 15

            # タグだけ一致（最低点）
            if s == 0:
                s += 1

            return s

        # 全行スコアリングして最高スコアの行を選択
        best_line  = None
        best_score = 0
        for i, line in enumerate(lines, 1):
            sc = score_line(i, line)
            if sc > best_score:
                best_score = sc
                best_line  = i

        # ---- エディタにジャンプ ----
        main_win = self.window()
        if best_line and hasattr(main_win, 'editor'):
            main_win.editor.goto_line(best_line)
            main_win.editor.setFocus()
            if hasattr(main_win, 'status'):
                main_win.status.showMessage(
                    f"ジャンプ: <{tag}> → Line {best_line}  (score={best_score})",
                    4000)
        elif hasattr(main_win, 'status'):
            main_win.status.showMessage(
                f"<{tag}> が見つかりませんでした", 3000)


# ============================================================
# エラーパネル
# ============================================================
class ErrorPanel(QWidget):
    errorClicked = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
        header = QLabel("🔍 構文チェック")
        header.setStyleSheet(
            f"color:{THEME['accent']};font-weight:bold;font-size:12px;"
            f"padding:4px 8px;background:{THEME['bg_panel']};")
        layout.addWidget(header)
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet(
            f"QListWidget{{background:{THEME['bg_panel']};color:{THEME['fg_text']};"
            f"border:none;font-size:11px;}}"
            f"QListWidget::item{{padding:4px 8px;border-bottom:1px solid {THEME['border']};}}"
            f"QListWidget::item:selected{{background:{THEME['bg_line']};}}")
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        layout.addWidget(self.list_widget)

    def update_errors(self, errors):
        self.list_widget.clear()
        icons  = {"error":"❌","warning":"⚠️","info":"ℹ️"}
        colors = {"error":THEME["error"],"warning":THEME["warning"],"info":THEME["accent2"]}
        for e in errors:
            kind = e.get("type","error")
            item = QListWidgetItem(
                f"{icons.get(kind,'❌')} Line {e['line']}: {e['message']}")
            item.setForeground(QColor(colors.get(kind, THEME["error"])))
            item.setData(Qt.ItemDataRole.UserRole, e)
            item.setToolTip(f"💡 提案: {e.get('suggestion','')}")
            self.list_widget.addItem(item)
        if not errors:
            item = QListWidgetItem("✅ 構文エラーなし")
            item.setForeground(QColor(THEME["accent2"]))
            self.list_widget.addItem(item)

    def _on_item_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data and "line" in data:
            self.errorClicked.emit(data["line"])


# ============================================================
# メインウィンドウ
# ============================================================
class HtmlMasterEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_file: Optional[str] = None
        self._current_ext = ".html"
        self.checker = SyntaxChecker()
        self._check_timer   = QTimer(); self._check_timer.setSingleShot(True)
        self._check_timer.timeout.connect(self._run_checks)
        self._preview_timer = QTimer(); self._preview_timer.setSingleShot(True)
        self._preview_timer.timeout.connect(self._update_preview)

        self.setWindowTitle("⚡ HTML Master Editor")
        self.resize(1600, 900)
        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._apply_theme()
        self._load_default()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左パネル
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0,0,0,0)
        left_layout.setSpacing(0)
        self.navigator = SectionNavigator()
        self.navigator.sectionSelected.connect(self._goto_line)
        left_layout.addWidget(self.navigator)
        self.btn_decouple = QPushButton("🔗 外部ファイル化")
        self.btn_decouple.setStyleSheet(
            f"QPushButton{{background:{THEME['bg_line']};color:{THEME['accent']};"
            f"border:none;padding:8px;font-weight:bold;font-size:12px;}}"
            f"QPushButton:hover{{background:{THEME['accent']};color:{THEME['bg_editor']};}}")
        self.btn_decouple.clicked.connect(self._decouple_files)
        left_layout.addWidget(self.btn_decouple)
        left_panel.setMinimumWidth(200)
        left_panel.setMaximumWidth(320)
        self.main_splitter.addWidget(left_panel)

        # 中央パネル
        center_splitter = QSplitter(Qt.Orientation.Vertical)
        self.editor = CodeEditor()
        self.editor.document().contentsChanged.connect(self._on_content_changed)
        center_splitter.addWidget(self.editor)
        self.error_panel = ErrorPanel()
        self.error_panel.errorClicked.connect(self._goto_line)
        self.error_panel.setMaximumHeight(200)
        center_splitter.addWidget(self.error_panel)
        center_splitter.setSizes([700, 150])
        self.main_splitter.addWidget(center_splitter)

        # 右パネル
        self.preview = PreviewWidget()
        self.preview.btn_refresh.clicked.connect(self._update_preview)
        self.preview.setMinimumWidth(300)
        self.main_splitter.addWidget(self.preview)

        self.main_splitter.setSizes([250, 750, 600])
        main_layout.addWidget(self.main_splitter)
        self.main_splitter.setStretchFactor(0, 0)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setStretchFactor(2, 1)

        # ステータスバー
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.setStyleSheet(
            f"QStatusBar{{background:{THEME['bg_panel']};"
            f"color:{THEME['fg_comment']};font-size:11px;}}")
        self.lbl_cursor = QLabel("行: 1  列: 1")
        self.lbl_cursor.setStyleSheet(f"color:{THEME['fg_text']};padding:0 8px;")
        self.lbl_file = QLabel("新規ファイル")
        self.lbl_file.setStyleSheet(f"color:{THEME['accent']};padding:0 8px;")
        self.lbl_errors = QLabel("✅ OK")
        self.lbl_errors.setStyleSheet(f"color:{THEME['accent2']};padding:0 8px;")
        self.lbl_lang = QLabel("HTML")
        self.lbl_lang.setStyleSheet(
            f"color:{THEME['fg_doctype']};padding:0 8px;font-weight:bold;")
        self.status.addWidget(self.lbl_file)
        self.status.addPermanentWidget(self.lbl_lang)
        self.status.addPermanentWidget(self.lbl_errors)
        self.status.addPermanentWidget(self.lbl_cursor)

        self.editor.cursorLineChanged.connect(self._update_cursor_status)
        self.editor.cursorPositionChanged.connect(self._update_cursor_col)

    def _setup_menu(self):
        menubar = self.menuBar()
        menubar.setStyleSheet(
            f"QMenuBar{{background:{THEME['bg_panel']};color:{THEME['fg_text']};}}"
            f"QMenuBar::item:selected{{background:{THEME['accent']};color:{THEME['bg_editor']};}}"
            f"QMenu{{background:{THEME['bg_panel']};color:{THEME['fg_text']};"
            f"border:1px solid {THEME['border']};}}"
            f"QMenu::item:selected{{background:{THEME['accent']};color:{THEME['bg_editor']};}}")

        file_menu = menubar.addMenu("ファイル(&F)")
        self._add_action(file_menu, "新規(&N)",            self._new_file,  "Ctrl+N")
        self._add_action(file_menu, "開く(&O)",            self._open_file, "Ctrl+O")
        self._add_action(file_menu, "保存(&S)",            self._save_file, "Ctrl+S")
        self._add_action(file_menu, "名前を付けて保存(&A)", self._save_as,   "Ctrl+Shift+S")
        file_menu.addSeparator()
        self._add_action(file_menu, "終了(&Q)", self.close, "Ctrl+Q")

        edit_menu = menubar.addMenu("編集(&E)")
        self._add_action(edit_menu, "元に戻す", self.editor.undo, "Ctrl+Z")
        self._add_action(edit_menu, "やり直し", self.editor.redo, "Ctrl+Y")
        edit_menu.addSeparator()
        self._add_action(edit_menu, "検索",       self._find_text,    "Ctrl+F")
        self._add_action(edit_menu, "行へジャンプ", self._jump_to_line, "Ctrl+G")

        view_menu = menubar.addMenu("表示(&V)")
        self._add_action(view_menu, "プレビュー更新",              self._update_preview,       "F5")
        self._add_action(view_menu, "構文チェック実行",            self._run_checks,           "F6")
        self._add_action(view_menu, "プレビューを別ウィンドウで開く", self.preview._toggle_popup, "F7")
        self._add_action(view_menu, "レイアウトをリセット",         self._reset_layout,         "F8")

        tool_menu = menubar.addMenu("ツール(&T)")
        self._add_action(tool_menu, "外部ファイル化（疎結合化）", self._decouple_files)
        self._add_action(tool_menu, "HTMLフォーマット整形",       self._format_html)

    def _add_action(self, menu, name, slot, shortcut=None):
        action = QAction(name, self)
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        action.triggered.connect(slot)
        menu.addAction(action)
        return action

    def _setup_toolbar(self):
        tb = QToolBar("メインツールバー")
        tb.setMovable(False)
        tb.setStyleSheet(
            f"QToolBar{{background:{THEME['bg_panel']};"
            f"border-bottom:1px solid {THEME['border']};spacing:4px;padding:2px 6px;}}"
            f"QToolButton{{background:transparent;color:{THEME['fg_text']};"
            f"border:none;padding:4px 10px;font-size:13px;}}"
            f"QToolButton:hover{{background:{THEME['bg_line']};border-radius:4px;}}")
        self.addToolBar(tb)
        buttons = [
            ("📄 新規",      self._new_file),
            ("📂 開く",      self._open_file),
            ("💾 保存",      self._save_file),
            ("|", None),
            ("🔄 プレビュー",  self._update_preview),
            ("🔍 構文チェック", self._run_checks),
            ("🔗 外部ファイル化", self._decouple_files),
            ("|", None),
            ("✨ フォーマット", self._format_html),
            ("|", None),          # ★ 追加
            ("🔡 文字サイズ", self._change_font_size),  # ★ 追加
        ]
        for label, slot in buttons:
            if label == "|":
                tb.addSeparator()
            else:
                btn = QPushButton(label)
                btn.clicked.connect(slot)
                btn.setStyleSheet(
                    f"QPushButton{{background:transparent;color:{THEME['fg_text']};"
                    f"border:none;padding:4px 10px;font-size:12px;}}"
                    f"QPushButton:hover{{background:{THEME['bg_line']};"
                    f"border-radius:4px;color:{THEME['accent']};}}")
                tb.addWidget(btn)

    def _apply_theme(self):
        self.setStyleSheet(
            f"QMainWindow{{background:{THEME['bg_editor']};}}"
            f"QSplitter::handle{{background:{THEME['border']};width:1px;height:1px;}}")

    def _reset_layout(self):
        total = self.main_splitter.width()
        self.main_splitter.setSizes([250, total - 850, 600])
        self.status.showMessage("レイアウトをリセットしました", 2000)

    def _set_language(self, ext: str):
        self._current_ext = ext
        self.editor.set_language(ext)
        lang = LANGUAGE_MAP.get(ext, ext.upper().lstrip('.'))
        self.lbl_lang.setText(lang)
        is_html = ext in (".html", ".htm")
        self.preview.url_bar.setText("プレビュー" if is_html else f"{lang} ビュー")
        self.btn_decouple.setEnabled(is_html)

    def _load_default(self):
        default = """<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HTML Master Editor へようこそ</title>
    <style>
        body {
            font-family: 'Segoe UI', sans-serif;
            max-width: 800px;
            margin: 40px auto;
            padding: 20px;
            background: #f5f5f5;
            color: #333;
        }
        h1 { color: #89B4FA; }
        .card {
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
    </style>
</head>
<body>
    <div class="card">
        <h1>⚡ HTML Master Editor</h1>
        <p>VSCodeを超えるHTML特化エディタへようこそ！</p>
        <ul>
            <li>✅ HTML / CSS / JS / Python / VBA / Batch 他 対応</li>
            <li>✅ リアルタイムプレビュー</li>
            <li>✅ 自動構文チェック</li>
            <li>✅ コード補完（Ctrl+Space）</li>
            <li>✅ セクションナビゲーター</li>
            <li>✅ 外部ファイル化（疎結合化）</li>
        </ul>
    </div>
    <script>
        console.log('HTML Master Editor 起動完了');
    </script>
</body>
</html>"""
        self.editor.setPlainText(default)
        self._set_language(".html")

    def _on_content_changed(self):
        self._check_timer.start(800)
        self._preview_timer.start(1200)
        self.navigator.update_structure(
            self.editor.toPlainText(), self._current_ext)

    def _run_checks(self):
        text   = self.editor.toPlainText()
        errors = self.checker.check(text, self._current_ext)
        self.error_panel.update_errors(errors)
        err_count  = sum(1 for e in errors if e.get("type") == "error")
        warn_count = sum(1 for e in errors if e.get("type") == "warning")
        if err_count:
            self.lbl_errors.setText(f"❌ {err_count}エラー {warn_count}警告")
            self.lbl_errors.setStyleSheet(f"color:{THEME['error']};padding:0 8px;")
        elif warn_count:
            self.lbl_errors.setText(f"⚠️ {warn_count}警告")
            self.lbl_errors.setStyleSheet(f"color:{THEME['warning']};padding:0 8px;")
        else:
            self.lbl_errors.setText("✅ OK")
            self.lbl_errors.setStyleSheet(f"color:{THEME['accent2']};padding:0 8px;")

    def _update_preview(self):
        text = self.editor.toPlainText()
        if self._current_ext in (".html", ".htm"):
            base = (str(Path(self.current_file).parent) + os.sep
                    if self.current_file else "")
            self.preview.load_html(text, base)
        else:
            lang = LANGUAGE_MAP.get(self._current_ext,
                                    self._current_ext.upper().lstrip('.'))
            self.preview.load_text(text, lang)

    def _goto_line(self, line: int):
        self.editor.goto_line(line)

    def _update_cursor_status(self, line: int):
        col = self.editor.textCursor().positionInBlock() + 1
        self.lbl_cursor.setText(f"行: {line}  列: {col}")

    def _update_cursor_col(self):
        cursor = self.editor.textCursor()
        self.lbl_cursor.setText(
            f"行: {cursor.blockNumber()+1}  列: {cursor.positionInBlock()+1}")

    def _new_file(self):
        self.editor.clear()
        self.current_file = None
        self.lbl_file.setText("新規ファイル")
        self.setWindowTitle("⚡ HTML Master Editor - 新規ファイル")
        self._set_language(".html")

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "ファイルを開く", "", FILE_FILTER)
        if path:
            try:
                with open(path, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()
            except Exception as e:
                QMessageBox.warning(self, "読み込みエラー", str(e))
                return
            self.editor.setPlainText(content)
            self.current_file = path
            ext = Path(path).suffix.lower()
            self.lbl_file.setText(Path(path).name)
            self.setWindowTitle(f"⚡ HTML Master Editor - {Path(path).name}")
            self._set_language(ext)

    def _save_file(self):
        if self.current_file: self._write_file(self.current_file)
        else: self._save_as()

    def _save_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "名前を付けて保存", "", FILE_FILTER)
        if path:
            self._write_file(path)
            self.current_file = path
            ext = Path(path).suffix.lower()
            self.lbl_file.setText(Path(path).name)
            self.setWindowTitle(f"⚡ HTML Master Editor - {Path(path).name}")
            self._set_language(ext)

    def _write_file(self, path: str):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.editor.toPlainText())
            self.status.showMessage(f"保存しました: {path}", 3000)
        except Exception as e:
            QMessageBox.warning(self, "保存エラー", str(e))

    def _decouple_files(self):
        if self._current_ext not in (".html", ".htm"):
            QMessageBox.information(self, "情報",
                "外部ファイル化はHTMLファイルのみ対応しています。")
            return
        if not self.current_file:
            QMessageBox.warning(self, "警告",
                "先にファイルを保存してください。\n"
                "保存先ディレクトリに .css/.js ファイルを生成します。")
            return
        html = self.editor.toPlainText()
        has_style  = bool(re.search(r'<style[^>]*>[\s\S]*?</style>',          html, re.I))
        has_script = bool(re.search(r'<script(?![^>]*src=)[^>]*>[\s\S]*?</script>', html, re.I))
        if not has_style and not has_script:
            QMessageBox.information(self, "情報",
                "切り出し対象の <style> / <script> ブロックが見つかりません。")
            return
        dlg = DecoupleDialog(self, has_style, has_script)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        base_dir = Path(self.current_file).parent
        new_html  = html
        if dlg.cb_style.isChecked():
            css_name = dlg.style_name.text().strip() or "style.css"
            css_contents = []
            def extract_css(m):
                css_contents.append(m.group(1).strip())
                return f'<link rel="stylesheet" href="{css_name}">'
            new_html = re.sub(r'<style[^>]*>([\s\S]*?)</style>',
                              extract_css, new_html, flags=re.I)
            with open(base_dir / css_name, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(css_contents))
        if dlg.cb_script.isChecked():
            js_name = dlg.script_name.text().strip() or "script.js"
            js_contents = []
            def extract_js(m):
                js_contents.append(m.group(1).strip())
                return f'<script src="{js_name}"></script>'
            new_html = re.sub(r'<script(?![^>]*src=)[^>]*>([\s\S]*?)</script>',
                              extract_js, new_html, flags=re.I)
            with open(base_dir / js_name, 'w', encoding='utf-8') as f:
                f.write('\n\n'.join(js_contents))
        self.editor.setPlainText(new_html)
        self._save_file()
        QMessageBox.information(self, "完了",
            f"外部ファイル化が完了しました！\n保存先: {base_dir}")

    def _format_html(self):
        if self._current_ext not in (".html", ".htm"):
            QMessageBox.information(self, "情報",
                "HTMLフォーマットはHTMLファイルのみ対応しています。")
            return
        try:
            html  = self.editor.toPlainText()
            lines = html.split('\n')
            formatted = []
            indent = 0
            INDENT = "    "
            BLOCK_TAGS = {"html","head","body","div","section","article","aside",
                          "header","footer","nav","main","ul","ol","table","thead",
                          "tbody","tr","form","select","style","script"}
            VOID = CodeEditor.VOID_TAGS
            for line in lines:
                stripped = line.strip()
                if not stripped:
                    formatted.append(""); continue
                close = re.match(r'</([a-zA-Z][a-zA-Z0-9\-]*)', stripped)
                if close and close.group(1).lower() in BLOCK_TAGS:
                    indent = max(0, indent - 1)
                formatted.append(INDENT * indent + stripped)
                open_tag = re.match(r'<([a-zA-Z][a-zA-Z0-9\-]*)([^>]*)>', stripped)
                if open_tag:
                    tag   = open_tag.group(1).lower()
                    attrs = open_tag.group(2)
                    if (tag in BLOCK_TAGS and tag not in VOID
                            and not stripped.startswith('</') and not attrs.endswith('/')):
                        if f"</{tag}>" not in stripped:
                            indent += 1
            self.editor.setPlainText('\n'.join(formatted))
            self.status.showMessage("フォーマット完了！", 2000)
        except Exception as e:
            self.status.showMessage(f"フォーマットエラー: {e}", 3000)

    def _find_text(self):
        text, ok = QInputDialog.getText(self, "検索", "検索文字列:")
        if ok and text:
            cursor = self.editor.document().find(text, self.editor.textCursor())
            if not cursor.isNull():
                self.editor.setTextCursor(cursor)
            else:
                self.status.showMessage(f"'{text}' が見つかりません", 2000)

    def _jump_to_line(self):
        line, ok = QInputDialog.getInt(
            self, "行へジャンプ", "行番号:",
            self.editor.textCursor().blockNumber() + 1,
            1, self.editor.document().blockCount())
        if ok:
            self._goto_line(line)
            
    def _change_font_size(self):
        """フォントサイズ変更ダイアログ"""
        current = self.editor.get_font_size()
        size, ok = QInputDialog.getInt(
            self,
            "文字サイズ変更",
            "フォントサイズ (6〜32pt):",
            current, 6, 32, 1
        )
        if ok:
            self.editor.set_font_size(size)
            self.status.showMessage(f"フォントサイズを {size}pt に変更しました", 2000)
    

# ============================================================
# エントリーポイント
# ============================================================
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("HTML Master Editor")
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window,          QColor(THEME["bg_editor"]))
    palette.setColor(QPalette.ColorRole.WindowText,      QColor(THEME["fg_text"]))
    palette.setColor(QPalette.ColorRole.Base,            QColor(THEME["bg_editor"]))
    palette.setColor(QPalette.ColorRole.AlternateBase,   QColor(THEME["bg_panel"]))
    palette.setColor(QPalette.ColorRole.Text,            QColor(THEME["fg_text"]))
    palette.setColor(QPalette.ColorRole.Button,          QColor(THEME["bg_panel"]))
    palette.setColor(QPalette.ColorRole.ButtonText,      QColor(THEME["fg_text"]))
    palette.setColor(QPalette.ColorRole.Highlight,       QColor(THEME["accent"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(THEME["bg_editor"]))
    app.setPalette(palette)
    
    app.setStyleSheet(f"QToolTip {{ background-color: {THEME['bg_panel']}; color: {THEME['fg_text']}; border: 1px solid {THEME['border']}; padding: 4px; }}")

    window = HtmlMasterEditor()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()