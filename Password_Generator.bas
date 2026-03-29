Attribute VB_Name = "Module1"
Option Explicit

' パスワード生成メインマクロ
Public Sub Password_Generator()
    Const PASSWORD_LENGTH As Long = 16
    Const TARGET_SHEET_NAME As String = "【パスワード生成】"
    Const TARGET_CELL_ADDRESS As String = "B4"
    Const HISTORY_SHEET_NAME As String = "【履歴】"
    Const HISTORY_HEADER_TIME As String = "生成時間"
    Const HISTORY_HEADER_PASSWORD As String = "生成パスワード"
    Const MAX_RETRY As Long = 1000
    
    Dim chars As String
    Dim pw As String
    Dim wsTarget As Worksheet
    Dim wsHistory As Worksheet
    Dim retryCount As Long
    
    ' 利用文字セット（固定）
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
            "abcdefghijklmnopqrstuvwxyz" & _
            "0123456789" & _
            "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
    
    ' 防御的チェック：文字セットが空でないか
    If Len(chars) = 0 Then
        MsgBox "利用文字セットが空です。コードを確認してください。", vbCritical
        Exit Sub
    End If
    
    ' 対象シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    On Error GoTo 0
    If wsTarget Is Nothing Then
        MsgBox "シート「" & TARGET_SHEET_NAME & "」が見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' 履歴シート取得（なければ作成）
    Set wsHistory = GetOrCreateHistorySheet(HISTORY_SHEET_NAME, _
                                            HISTORY_HEADER_TIME, _
                                            HISTORY_HEADER_PASSWORD)
    If wsHistory Is Nothing Then
        MsgBox "履歴シートの準備に失敗しました。", vbCritical
        Exit Sub
    End If
    
    ' 乱数初期化
    Randomize
    
    ' 重複しないパスワードができるまで生成＆チェック
    retryCount = 0
    Do
        pw = CreateRandomPassword(PASSWORD_LENGTH, chars)
        retryCount = retryCount + 1
        
        If retryCount > MAX_RETRY Then
            MsgBox "重複しないパスワードの生成に失敗しました。試行回数上限に達しました。", vbCritical
            Exit Sub
        End If
    Loop While IsPasswordExists(wsHistory, pw, HISTORY_HEADER_PASSWORD)
    
    ' B4セルに上書き（式誤認回避のため先頭に ' を付ける）
    wsTarget.Range(TARGET_CELL_ADDRESS).Value = "'" & pw
    
    ' 履歴に追記
    AppendHistory wsHistory, pw, HISTORY_HEADER_TIME, HISTORY_HEADER_PASSWORD

    ' 完了ダイアログ
    MsgBox "パスワードを生成しました", vbInformation, "完了"
End Sub

' ランダムパスワード生成
Private Function CreateRandomPassword(ByVal length As Long, ByVal chars As String) As String
    Dim i As Long
    Dim result As String
    Dim index As Long
    Dim maxIndex As Long
    
    maxIndex = Len(chars)
    
    For i = 1 To length
        index = Int(maxIndex * Rnd) + 1
        result = result & Mid$(chars, index, 1)
    Next i
    
    CreateRandomPassword = result
End Function

' 履歴シート取得 or 作成＋ヘッダ設定
Private Function GetOrCreateHistorySheet( _
    ByVal sheetName As String, _
    ByVal headerTime As String, _
    ByVal headerPassword As String) As Worksheet
    
    Dim ws As Worksheet
    Dim isNew As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' シートがない場合は作成
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        isNew = True
    End If
    
    ' ヘッダ行の設定（A1=生成時間、B1=生成パスワード）
    If isNew Or (ws.Range("A1").Value = "" And ws.Range("B1").Value = "") Then
        ws.Range("A1").Value = headerTime
        ws.Range("B1").Value = headerPassword
    End If
    
    Set GetOrCreateHistorySheet = ws
End Function

' パスワードが履歴に存在するかチェック
Private Function IsPasswordExists( _
    ByVal wsHistory As Worksheet, _
    ByVal password As String, _
    ByVal headerPassword As String) As Boolean
    
    Dim lastRow As Long
    Dim passwordCol As Long
    Dim found As Range
    
    ' パスワード列はB列固定
    passwordCol = 2
    
    lastRow = wsHistory.Cells(wsHistory.Rows.Count, passwordCol).End(xlUp).Row
    If lastRow < 2 Then
        IsPasswordExists = False
        Exit Function
    End If
    
    Set found = wsHistory.Range(wsHistory.Cells(2, passwordCol), _
                                wsHistory.Cells(lastRow, passwordCol)) _
                         .Find(What:=password, LookIn:=xlValues, LookAt:=xlWhole)
    
    IsPasswordExists = Not (found Is Nothing)
End Function

' 履歴に1行追記（A列=時間、B列=パスワード）
Private Sub AppendHistory( _
    ByVal wsHistory As Worksheet, _
    ByVal password As String, _
    ByVal headerTime As String, _
    ByVal headerPassword As String)
    
    Dim lastRow As Long
    
    lastRow = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    lastRow = lastRow + 1
    
    ' 記録（式誤認回避のため先頭に ' を付ける）
    wsHistory.Cells(lastRow, 1).Value = Now
    wsHistory.Cells(lastRow, 1).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    wsHistory.Cells(lastRow, 2).Value = "'" & password
    
    ' 格子線（縁取り）
    With wsHistory.Range(wsHistory.Cells(lastRow, 1), wsHistory.Cells(lastRow, 2)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

' コピー機能
Public Sub CopyPassword()
    Const TARGET_SHEET_NAME As String = "【パスワード生成】"
    Const TARGET_CELL_ADDRESS As String = "B4"
    
    Dim ws As Worksheet
    Dim pw As String
    
    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    pw = ws.Range(TARGET_CELL_ADDRESS).Value
    
    If pw = "" Then
        MsgBox "コピーするパスワードがありません。", vbExclamation, "注意"
        Exit Sub
    End If
    
    ' クリップボードへコピー
    With CreateObject("htmlfile")
        .parentWindow.clipboardData.setData "text", pw
    End With
    
    MsgBox "コピーしました。", vbInformation, "完了"
End Sub

