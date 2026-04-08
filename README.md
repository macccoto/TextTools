# TextTools.dll

Excel VBA から呼び出せる、文字列正規化ネイティブ DLL（Windows / C言語）。

Excel の `Replace` / `Trim` では対応しづらい、改行・タブ・制御文字の混在した文字列を安全にクレンジングします。

---

## 収録ファイル

| ファイル | 説明 |
|---|---|
| `TextTools.c` | DLL ソースコード（C言語） |
| `TextTools.h` | ヘッダファイル |
| `TextTools.def` | エクスポート定義（32bit 名前装飾回避用） |
| `build.bat` | MSVC ビルドスクリプト |
| `TextTools64.dll` | ビルド済み DLL（64bit） |
| `TextTools32.dll` | ビルド済み DLL（32bit） |
| `TestTextTools.bas` | VBA テストモジュール |

---

## 実装済み関数

### `NormalizeTextW`

```c
int __stdcall NormalizeTextW(
    const wchar_t* src,
    wchar_t*       dest,
    int            destChars
);
```

UTF-16 文字列を正規化して `dest` バッファに書き込みます。

#### 正規化ルール

1. 先頭・末尾の半角スペース / タブ / CR / LF を除去
2. CRLF / CR / LF をすべて半角スペース 1 個に統一
3. タブを半角スペース 1 個に変換
4. 連続する半角スペースを 1 個に圧縮
5. 制御文字（0x01〜0x1F、CR / LF / TAB を除く）を削除

#### 変換例

| 入力 | 出力 |
|---|---|
| `"  ABC   DEF  "` | `"ABC DEF"` |
| `"A\r\nB\rC\nD"` | `"A B C D"` |
| `"A\t\tB"` | `"A B"` |
| `"A" & Chr(1) & "B"` | `"AB"` |

#### 戻り値

| 値 | 意味 |
|---|---|
| `> 0` | 書き込んだ文字数（NUL を除く）。`Left$(dest, ret)` で結果を取り出す |
| `= 0` | 結果が空文字列、または引数エラー |
| `>= destChars` | バッファ不足。戻り値 + 1 のサイズで再試行 |

---

## 使い方

### 1. DLL を配置する

Office の bitness に合わせた DLL を用意してください。

```
Office 32bit → TextTools32.dll を使用
Office 64bit → TextTools64.dll を使用
```

> Office の bitness は `ファイル → アカウント → Excel のバージョン情報` で確認できます。

フルパスで指定するのが最も確実です。

```
C:\Users\yourname\lib\TextTools64.dll
```

### 2. VBA に Declare 文を追加する

標準モジュールの先頭に貼り付けます。

```vb
#If VBA7 Then
    ' Office 2010 以降（32bit / 64bit 両対応）
    Private Declare PtrSafe Function NormalizeTextW Lib _
        "C:\path\to\TextTools64.dll" ( _
        ByVal src       As LongPtr, _
        ByVal dest      As LongPtr, _
        ByVal destChars As Long _
    ) As Long
#Else
    ' Office 2007 以前（32bit のみ）
    Private Declare Function NormalizeTextW Lib _
        "C:\path\to\TextTools32.dll" ( _
        ByVal src       As Long, _
        ByVal dest      As Long, _
        ByVal destChars As Long _
    ) As Long
#End If
```

> **重要:** `ByVal` を必ず付けてください。`ByRef` にするとクラッシュします。

### 3. 呼び出す

```vb
Function Normalize(ByVal src As String) As String
    Dim dest    As String
    Dim bufSize As Long
    Dim ret     As Long

    bufSize = Len(src) + 2
    If bufSize < 16 Then bufSize = 16

    dest = String$(bufSize, Chr$(0))
    ret  = NormalizeTextW(StrPtr(src), StrPtr(dest), bufSize)

    If ret >= bufSize Then
        ' バッファ不足 → 必要サイズで再試行
        bufSize = ret + 1
        dest    = String$(bufSize, Chr$(0))
        ret     = NormalizeTextW(StrPtr(src), StrPtr(dest), bufSize)
    End If

    Normalize = Left$(dest, ret)
End Function
```

---

## ビルド方法

**Visual Studio Build Tools 2022**（または VS 2022）が必要です。

### 64bit

`x64 Native Tools Command Prompt` を開いて実行：

```cmd
cl /LD /W4 /nologo /utf-8 TextTools.c /link /DEF:TextTools.def /OUT:TextTools64.dll /MACHINE:X64
```

### 32bit

`x86 Native Tools Command Prompt` を開いて実行：

```cmd
cl /LD /W4 /nologo /utf-8 TextTools.c /link /DEF:TextTools.def /OUT:TextTools32.dll /MACHINE:X86
```

### build.bat を使う場合

`build.bat` をダブルクリックするか、コマンドプロンプトで実行してください。  
64bit / 32bit を自動で続けてビルドします。

---

## よくあるエラー

| エラー | 原因 | 対処 |
|---|---|---|
| `ファイルが見つかりません` | `Lib` パスが間違っている | フルパスで指定する |
| `Entry point not found` | DLL のエクスポート名が違う | `dumpbin /EXPORTS TextTools.dll` で確認。`.def` ファイルを使ってビルドし直す |
| `Bad DLL calling convention` | 呼び出し規約の不一致 | ソースに `__stdcall` が付いているか確認 |
| `型が一致しません` | `ByRef` になっている | `ByVal` に直す |
| Excel がクラッシュ | DLL の bitness が Office と合っていない | 32bit Office には 32bit DLL を使う |

---

## 設計方針

- **C 言語**で実装（C++ の名前マングリングを避ける）
- **呼び出し元バッファ渡し**（DLL 内で `malloc` しない → メモリリーク・クロスランタイム問題を排除）
- **`__stdcall`**（Windows API 標準 / VBA `Declare` のデフォルト）
- **`.def` ファイル**で 32bit の名前装飾（`_NormalizeTextW@12`）を回避

---

## 今後の追加候補

| 関数名（案） | 内容 |
|---|---|
| `NormalizeFullHalfW` | 全角 → 半角 / 半角 → 全角変換（`LCMapStringW` 使用） |
| `RemoveCharsW` | 指定文字セットの一括削除 |
| `CountTokensW` | 区切り文字でのトークンカウント |
| `ReplaceTextW` | 高速文字列置換 |

---

## ライセンス

MIT
