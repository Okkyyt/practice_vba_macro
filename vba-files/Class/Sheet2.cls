VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' 構成１－２

Option Explicit

Public Sub Merge_Youshiki1_2()
    Dim folderPath As String
    Dim sourceSheetName As String
    Dim destSheetName As String
    Dim sourceStartRow As Long
    Dim outputStartRow As Long

    ' 貼り付け元ワークブック・シート
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    ' 貼り付け先ワークブック・シート
    Dim wbDest As Workbook
    Dim wsDest As Worksheet

    Dim fileList As Collection
    Dim filePath As Variant
    Dim appended As Long

    ' folderPath = wbDest.Path & "\xls-files"
    sourceSheetName = "様式１－２"
    destSheetName = "様式１-２"
    sourceStartRow = 8
    outputStartRow = 7

    ' マージ先シート
    Set wbDest = ThisWorkbook
    Set wsDest = wbDest.Worksheets(destSheetName)

    ' フォルダ選択ボタン作成
    ' 実行ボタン作成

    ' "folderPath" をフォルダーを選択で取得
    folderPath = select_folder()
    If folderPath = "" Then
        MsgBox "フォルダーが選択されませんでした", vbExclamation
        Exit Sub
    End If

    ' マージのためのExcelファイルが格納されたフォルダからファイルを取得（パス：./xls-files）
    Set fileList = get_file_list(folderPath)
    MsgBox "Excelファイル数：" & CStr(fileList.Count), vbInformation
    If fileList Is Nothing Or fileList.Count = 0 Then
        MsgBox "Excelファイルがありません", vbExclamation
        Exit Sub
    End If

    ' マージ先シートの既存データをクリア：8行目から最終行（通常 1048576）まで
    wsDest.Rows(sourceStartRow & ":" & wsDest.Rows.Count).Clear

    ' 各ファイルごとにデータを抽出し、マージ先ファイルの様式１－３シートに追記
    ' 8行目以降のデータをすべて取得し、マージ先ファイルにconcatで追記していく
    For Each filePath In fileList

        Set wbSrc = Workbooks.Open(CStr(filePath), ReadOnly:=True)

        ' マージ元シート
        Set wsSrc = wbSrc.Worksheets(sourceSheetName)

        ' マージ
        appended = concat_sheet_data(wsDest, wsSrc, outputStartRow, sourceStartRow)

        ' 次の貼り付け開始行を更新
        outputStartRow = outputStartRow + appended

        wbSrc.Close SaveChanges:=False
    Next filePath

    MsgBox "マージ完了", vbInformation
End Sub
