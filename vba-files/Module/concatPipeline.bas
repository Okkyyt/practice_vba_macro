Option Explicit

Public Function Concat_Sheet( _
    ByVal destSheetName As String, _
    ByVal srcSheetName As String, _
    ByVal destStartRow As Long, _
    ByVal srcStartRow As Long, _
    ByVal copyColCount As Long, _
    ByVal folderPath As String, _
    ByVal fileList As Collection _
) As Long


    ' 貼り付け先ワークブック・シート
    Dim wbDest As Workbook
    Dim wsDest As Worksheet

    ' 貼り付け元ワークブック・シート
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet

    Dim filePath As Variant
    Dim appended As Long

    ' マージ先シート
    Set wbDest = ThisWorkbook
    Set wsDest = wbDest.Worksheets(destSheetName)

    ' マージ先シートの既存データをクリア：8行目から最終行（通常 1048576）まで
    wsDest.Rows(destStartRow & ":" & wsDest.Rows.Count).Clear

    ' 進捗バーの準備
    Dim total As Long
    Dim i As Long
    total = 0
    On Error Resume Next
    total = fileList.Count
    On Error GoTo 0
    If total > 0 Then
        UserForm4.StartProgress total, "データ結合中…"
    End If

    ' 各ファイルごとにデータを抽出し、マージ先ファイルの様式１－３シートに追記
    ' 8行目以降のデータをすべて取得し、マージ先ファイルにconcatで追記していく
    For Each filePath In fileList

        ' マージ元ファイルを開く（読み取り専用）
        Set wbSrc = Workbooks.Open(CStr(filePath), ReadOnly:=True)

        ' マージ元シートの特定
        Set wsSrc = wbSrc.Worksheets(srcSheetName)

        ' マージ
        appended = concat_sheet_data(wsDest, wsSrc, destStartRow, srcStartRow, copyColCount)

        ' 次の貼り付け開始行を更新
        destStartRow = destStartRow + appended

        wbSrc.Close SaveChanges:=False

        ' 進捗更新
        i = i + 1
        Dim msg As String
        msg = "処理中: " & VBA.Mid$(CStr(filePath), VBA.InStrRev(CStr(filePath), "\") + 1)
        On Error Resume Next
        UserForm4.UpdateProgress i, msg
        On Error GoTo 0
    Next filePath

    ' 進捗終了
    If total > 0 Then
        On Error Resume Next
        UserForm4.FinishProgress
        On Error GoTo 0
    End If
    Concat_Sheet = destStartRow - 1 ' 最終行を返す
End Function
