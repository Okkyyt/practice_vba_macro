VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' 実行画面
Option Explicit

' フォルダ選択ボタンが押された際の処理（フォルダ選択ボタン）
Public Sub onClick_SelectFolder()
    Dim folderPath As String
    folderPath = select_folder()
    
    If folderPath <> "" Then
        Range("D8").Value = folderPath
    End If
End Sub

' 実行ボタンが押された際の処理（実行ボタン）
Public Sub onClick_Exec_Button()
    
    Dim folderPath As String
    Dim fileList As Collection
    folderPath = Range("D8").Value
    
    ' 入力チェック
    If folderPath = "" Then
        MsgBox "フォルダーが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "指定されたフォルダーが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' マージのためのExcelファイルが格納されたフォルダからファイルを取得
    Set fileList = get_file_list(folderPath)
    If fileList Is Nothing Or fileList.Count = 0 Then
        MsgBox "Excelファイルがありません", vbExclamation
        Exit Sub
    End If

    ' 画面更新を無効化して処理を実行
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' 構成１－２
    Sheet2.sheet_1_2 folderPath, fileList

    ' 構成１－３
    Sheet3.sheet_1_3 folderPath, fileList

    ' 構成１－４
    Sheet4.sheet_1_4 folderPath, fileList

    ' 構成１－５
    Sheet5.sheet_1_5 folderPath, fileList
    
    ' 正常終了
    Application.ScreenUpdating = True
    MsgBox "完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
End Sub

' フォルダ選択ダイアログを表示するヘルパー関数
Private Function select_folder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダーを選択してください"
        If .Show = -1 Then
            select_folder = .SelectedItems(1)
        Else
            select_folder = ""
        End If
    End With
End Function
