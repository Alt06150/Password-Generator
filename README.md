# 【For overseas】
# Password-Generator（VBA）

## Overview
This is an Excel tool that generates 16-character passwords.
It includes symbols, uppercase and lowercase letters, and numbers,
and the generated passwords are saved in the history.

## Key Features
- Generates passwords with a fixed length of 16 characters.
- Secure composition that includes symbols, uppercase letters, lowercase letters, and numbers.
- Saves generated passwords in a history.
- Runs as an Excel macro (no additional software required).

## File Structure
```
Password_Generator/
├── Password_Generator.xlsm     # Excel file for execution
└── modules/
     └── PasswordGenerator.bas  # VBA code (for viewing)
```

## How to Use
1. Open “Password_Generator.xlsm”.
2. If an Excel warning appears, click “Enable Content”.
3. Click the “パスワード生成” button to generate the password(“パスワード生成” means “password generation” in English).
4. Click the “コピー” button to copy the password to the clipboard(“コピー” means “copy” in English).

## Bonus
・　When you generate a password, it will be recorded in the “履歴” sheet(“履歴” means “history” in English).

## System Requirements
- Windows 10 / 11
- Microsoft Excel（with macros enabled）

## Important Notes
- This tool is intended for personal use only.
- You are solely responsible for managing your passwords.
- This tool will not work unless macros are enabled.
 
---
 
# 【日本向け説明】
パスワード生成ツール（VBA）

## 概要
16文字のパスワードを生成する Excel ツールです。
記号・大小英数字を含み、生成したパスワードは履歴として保存されます。

## 主な機能
- 16 文字固定のパスワード生成
- 記号・大文字・小文字・数字を含む安全な構成
- 生成したパスワードを履歴として保存
- Excel マクロで動作（追加ソフト不要）

## ファイル構成
```
Password_Generator/
├── Password_Generator.xlsm     # 実行用Excelファイル
└── modules/
     └── PasswordGenerator.bas  # VBAコード（閲覧用）
```

## 使い方
1. 「Password_Generator.xlsm」を開く。
2. Excel の警告が出た場合は 「コンテンツの有効化」 をクリックする。
3. 「パスワード生成」ボタンを押して生成。
4. 「コピー」ボタンで、クリップボードにパスワードをコピーする。

## おまけ
・ パスワードを生成すると、「履歴」シートに生成したパスワードが記録されます。

## 動作環境
- Windows 10 / 11
- Microsoft Excel（マクロ有効版）

## 注意事項
- 本ツールは個人利用を想定しています。
- パスワードの管理は自己責任で行ってください。
- マクロを有効化しないと動作しません。

