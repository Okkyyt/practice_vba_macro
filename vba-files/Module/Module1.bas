Attribute VB_Name = "Module11"
Option Explicit

' フォルダ選択ダイアログを表示し、選択されたフォルダパスを取得
' 引数：
'   prompt … ダイアログの説明文
' 戻り値：
'   String … 選択されたフォルダパス（キャンセル時は空文字列）
Public Function select_folder( _
    Optional ByVal prompt As String = "フォルダを選択してください" _
) As String
    Dim fldr As FileDialog
    Dim folderPath As String

    ' フォルダ選択ダイアログの表示
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = prompt
        If .Show <> -1 Then
            select_folder = ""   ' キャンセル時は空文字列を返す
            Exit Function
        End If
        folderPath = .SelectedItems(1)
    End With

    ' 返り値は関数名にセット
    select_folder = folderPath
End Function


' マージのためのExcelファイルが格納されたフォルダからExcelファイルを取得
' 引数：
'   folderPath … フォルダパス（末尾 \ はあってもなくてもOK）
'   filePattern … 検索パターン（省略時 "*.xls*"）
' 戻り値：
'   Collection … 各要素が「フルパス文字列」のコレクション
Public Function get_file_list( _
    ByVal folderPath As String, _
    Optional ByVal filePattern As String = "*.xls*" _
) As Collection

    Dim file_list As New Collection
    Dim file_name As String

    ' フォルダ末尾に "\" がなければ追加
    If Len(folderPath) > 0 Then
        If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then
            folderPath = folderPath & "\"
        End If
    End If

    ' 指定フォルダ内のExcelファイルをすべて取得
    file_name = Dir$(folderPath & filePattern)
    Do While file_name <> ""
        file_list.Add folderPath & file_name
        file_name = Dir$()
    Loop

    Set get_file_list = file_list
End Function


' ファイルのデータを抽出し、マージ先ファイルのシートの最終行に追記（縦にconcat）
' 引数：
'   wsDest        … マージ先シート
'   wsSrc         … マージ元シート
'   outputStartRow … マージ先シートの追記開始行
'   sourceStartRow … マージ元シートのデータ開始行
' 戻り値：
'   Long … 追記した行数（データ無しなら 0）
Public Function concat_sheet_data( _
    ByVal wsDest As Worksheet, _
    ByVal wsSrc As Worksheet, _
    ByVal outputStartRow As Long, _
    ByVal sourceStartRow As Long _
) As Long

    ' コピー元の最大行を取得
    Dim srcMaxRow As Long
    ' コピー元の最大列を取得
    Dim srcMaxCol As Long
    srcMaxCol = 21 ' U列までコピー
    ' コピー元のコピー範囲
    Dim rngSrc As Range
    ' コピー先の貼り付け範囲
    Dim rngDest As Range
    ' 最大行を検索するための基準列
    Dim baseCol As Long
    baseCol = 2 ' B列を基準に最大行を検索


    ' 下方向で最大行を検索
    srcMaxRow = wsSrc.Cells(sourceStartRow, baseCol).End(xlDown).Row
    ' 最大行がデータ開始行未満, もしくは最終行の場合、データ無しとみなして終了
    If srcMaxRow < sourceStartRow Or srcMaxRow = wsSrc.Rows.Count Then
        concat_sheet_data = 0
        Exit Function
    End If

    ' コピー範囲
    Set rngSrc = wsSrc.Range(wsSrc.Cells(sourceStartRow, 1), wsSrc.Cells(srcMaxRow, srcMaxCol))

    ' デバッグ
    MsgBox "コピー元  shape" & rngSrc.Rows.Count & " x " & rngSrc.Columns.Count

    ' コピー元へ貼り付け
    Set rngDest = wsDest.Cells(outputStartRow, 1).Resize(rngSrc.Rows.Count, rngSrc.Columns.Count)
    rngDest.Value = rngSrc.Value

    ' 追加した行数を返す
    concat_sheet_data = rngSrc.Rows.Count
End Function
