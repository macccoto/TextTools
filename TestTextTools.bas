Attribute VB_Name = "TextToolsTest"
'===========================================================
' TextToolsTest.bas
' -----------------
' TextTools.dll の NormalizeTextW 関数テストコード
' 標準モジュールとして貼り付けて使用する
'
' 使い方:
'   1. このモジュールを VBE で標準モジュールに貼り付ける
'   2. TestNormalizeTextW を実行（F5 or マクロ一覧から）
'   3. イミディエイトウィンドウ（Ctrl+G）で結果を確認
'===========================================================

Option Explicit

' ----------------------------------------------------------
' DLL 宣言
' ----------------------------------------------------------

#If VBA7 Then
    ' VBA7 = Office 2010 以降（32bit / 64bit 両対応）
    ' PtrSafe: 64bit VBA で必須のキーワード
    ' LongPtr: 32bit では Long(4byte)、64bit では LongLong(8byte) に自動切り替わる

    Private Declare PtrSafe Function NormalizeTextW Lib _
        "C:\Users\konan\OneDrive\ClaudeCode用\TextTools\TextTools64.dll" ( _
        ByVal src       As LongPtr, _
        ByVal dest      As LongPtr, _
        ByVal destChars As Long _
    ) As Long

#Else
    ' VBA6 = Office 2007 以前（32bit のみ）

    Private Declare Function NormalizeTextW Lib _
        "C:\Users\konan\OneDrive\ClaudeCode用\TextTools\TextTools32.dll" ( _
        ByVal src       As Long, _
        ByVal dest      As Long, _
        ByVal destChars As Long _
    ) As Long

#End If

' ----------------------------------------------------------
' ヘルパー: NormalizeTextW の安全なラッパー
' ----------------------------------------------------------

'
' SafeNormalize
' -------------
' NormalizeTextW を呼び出すラッパー関数。
' バッファ不足のとき自動的に再試行する。
'
' 引数:
'   src  正規化したい文字列
' 戻り値:
'   正規化後の文字列（失敗時は空文字列）
'
Private Function SafeNormalize(ByVal src As String) As String
    Dim bufSize  As Long    ' バッファサイズ（文字数、NUL 含む）
    Dim dest     As String  ' 出力バッファ（事前に確保する）
    Dim ret      As Long    ' DLL の戻り値

    ' 初期バッファサイズ: 入力長 + 1（NUL 分）で十分なはず
    ' 正規化で文字数が増えることはないが、念のため余裕を持たせる
    bufSize = Len(src) + 2
    If bufSize < 16 Then bufSize = 16  ' 最低 16 文字

    ' --- 1 回目の呼び出し ---
    dest = String$(bufSize, Chr$(0))  ' NUL 文字で埋めた文字列をバッファとして使用
    ret  = NormalizeTextW(StrPtr(src), StrPtr(dest), bufSize)

    ' バッファ不足チェック: ret >= bufSize ならバッファ不足
    If ret >= bufSize Then
        ' ret = 必要文字数（NUL 除く）なので +1 して再試行
        bufSize = ret + 1
        dest    = String$(bufSize, Chr$(0))
        ret     = NormalizeTextW(StrPtr(src), StrPtr(dest), bufSize)
    End If

    ' 正常終了（ret > 0）: ret 文字分を切り出す
    ' ret = 0 の場合は空文字列
    If ret > 0 Then
        SafeNormalize = Left$(dest, ret)
    Else
        SafeNormalize = ""
    End If
End Function

' ----------------------------------------------------------
' テストメイン
' ----------------------------------------------------------

Public Sub TestNormalizeTextW()
    Dim passCount As Long
    Dim failCount As Long

    Debug.Print "================================================"
    Debug.Print "  NormalizeTextW テスト開始"
    Debug.Print "================================================"

    passCount = 0
    failCount = 0

    ' テスト実行
    RunTest "TC01", "  ABC   DEF  ",        "ABC DEF",    passCount, failCount
    RunTest "TC02", "A" & vbCrLf & "B",    "A B",        passCount, failCount
    RunTest "TC03", "A" & vbCr & "B",       "A B",        passCount, failCount
    RunTest "TC04", "A" & vbLf & "B",       "A B",        passCount, failCount
    RunTest "TC05", "A" & vbTab & vbTab & "B", "A B",     passCount, failCount
    RunTest "TC06", "A" & Chr$(1) & "B",    "AB",         passCount, failCount
    RunTest "TC07", "A" & Chr$(31) & "B",   "AB",         passCount, failCount
    RunTest "TC08", "",                      "",           passCount, failCount
    RunTest "TC09", "   ",                   "",           passCount, failCount
    RunTest "TC10", vbTab & vbTab,           "",           passCount, failCount
    RunTest "TC11", "ABC",                   "ABC",        passCount, failCount
    RunTest "TC12", " A ",                   "A",          passCount, failCount

    ' 複合ケース: 改行 + 制御文字 + タブ + 連続スペース
    RunTest "TC13", vbTab & " A" & Chr$(2) & vbCrLf & "B  C " & vbCr, _
                    "A B C",                              passCount, failCount

    ' 全角文字（変換せず通過するはず）
    RunTest "TC14", "  ＡＢＣ  ",            "ＡＢＣ",    passCount, failCount

    ' 改行が複数連続
    RunTest "TC15", "A" & vbCrLf & vbCrLf & "B",  "A B", passCount, failCount

    Debug.Print "================================================"
    Debug.Print "  結果: " & passCount & " PASS / " & failCount & " FAIL"
    Debug.Print "================================================"

    If failCount = 0 Then
        MsgBox "全テスト PASS (" & passCount & " 件)", vbInformation, "TextTools テスト"
    Else
        MsgBox failCount & " 件 FAIL あり。イミディエイトウィンドウを確認してください。", _
               vbExclamation, "TextTools テスト"
    End If
End Sub

' ----------------------------------------------------------
' テストヘルパー: 1 ケース実行して結果表示
' ----------------------------------------------------------

Private Sub RunTest( _
    ByVal id       As String, _
    ByVal input    As String, _
    ByVal expected As String, _
    ByRef passCount As Long, _
    ByRef failCount As Long _
)
    Dim actual   As String
    Dim result   As String
    Dim dispIn   As String

    ' 入力の制御文字を可視化（表示用）
    dispIn = MakeVisible(input)

    actual = SafeNormalize(input)

    If actual = expected Then
        result = "PASS"
        passCount = passCount + 1
    Else
        result = "FAIL"
        failCount = failCount + 1
    End If

    Debug.Print "[" & id & "] " & result & _
                "  入力=[" & dispIn & "]" & _
                "  期待=[" & expected & "]" & _
                "  実際=[" & actual & "]"
End Sub

' ----------------------------------------------------------
' デバッグ用: 制御文字を可視化する
' ----------------------------------------------------------

Private Function MakeVisible(ByVal s As String) As String
    Dim i   As Long
    Dim c   As Integer
    Dim buf As String

    buf = ""
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        Select Case c
            Case 9:  buf = buf & "<TAB>"
            Case 10: buf = buf & "<LF>"
            Case 13: buf = buf & "<CR>"
            Case Is < 32: buf = buf & "<" & c & ">"
            Case Else: buf = buf & Chr$(c)
        End Select
    Next i
    MakeVisible = buf
End Function

' ----------------------------------------------------------
' 低レベルテスト: DLL を直接呼び出す例（学習用）
' ----------------------------------------------------------

Public Sub LowLevelTest()
    Dim src     As String
    Dim dest    As String
    Dim bufSize As Long
    Dim ret     As Long

    src     = "  Hello" & vbTab & "World  " & vbCrLf
    bufSize = 256
    dest    = String$(bufSize, Chr$(0))

    '-------------------------
    ' StrPtr(s) : VBA 文字列の内部バッファ（UTF-16）へのポインタを返す
    ' String$(n, c): 文字 c を n 個並べた文字列を生成（バッファ確保に使う）
    ' Left$(s, n)  : 文字列 s の左 n 文字を取得（戻り値で切り詰めに使う）
    '-------------------------
    ret = NormalizeTextW(StrPtr(src), StrPtr(dest), bufSize)

    Debug.Print "LowLevel ret=" & ret
    Debug.Print "LowLevel result=[" & Left$(dest, ret) & "]"
End Sub
