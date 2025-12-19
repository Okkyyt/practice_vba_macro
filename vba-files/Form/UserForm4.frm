VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   1930
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 進捗バー管理用の簡易 API（モジュールレベル変数）
Private m_Total As Long
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()

   ' 既定の初期化（実際の値は StartProgress で設定）
   ProgressBar1.Min = 0
   ProgressBar1.Max = 100
   ProgressBar1.Value = 0
   On Error Resume Next
   Label1.Caption = ""
   ' 少しフォントサイズを小さく
   Me.Font.Size = 9
   Label1.Font.Size = 9
   On Error GoTo 0
   
End Sub

 

' 進捗の開始：最大値やタイトルを設定してフォームをモデルレス表示
Public Sub StartProgress(ByVal total As Long, Optional ByVal title As String = "")
   m_Total = IIf(total > 0, total, 100)
   With Me.ProgressBar1
      .Min = 0
      .Max = m_Total
      .Value = 0
   End With
   If Len(title) > 0 Then Me.Caption = title
   On Error Resume Next
   Me.Label1.Caption = "準備中…"
   On Error GoTo 0
   Me.Show vbModeless
   DoEvents
End Sub

' 現在値に応じて進捗を更新（メッセージも任意で更新）
Public Sub UpdateProgress(ByVal current As Long, Optional ByVal message As String = "")
   With Me.ProgressBar1
      Dim newVal As Long
      newVal = current
      If newVal < .Min Then newVal = .Min
      If newVal > .Max Then newVal = .Max
      .Value = newVal
   End With
   If Len(message) > 0 Then
      On Error Resume Next
      Me.Label1.Caption = message
      On Error GoTo 0
   End If
   DoEvents
End Sub

' 進捗の終了：フォームを閉じる
Public Sub FinishProgress()
   Unload Me
End Sub


