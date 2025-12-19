VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ç\ê¨ÇPÅ|ÇR

Option Explicit

Public Sub sheet_1_3( _
  ByVal folderPath As String _
)
    
    Dim destSheetName As String
    Dim srcSheetName As String
    Dim destStartRow As Long
    Dim srcStartRow As Long
    Dim copyColCount As Long
    Dim appended As Long

    ' folderPath = wbDest.Path & "\xls-files"
    destSheetName = "óléÆÇP-ÇR"
    srcSheetName = "óléÆÇPÅ|ÇR"
    destStartRow = 7
    srcStartRow = 8
    copyColCount = 25


    appended = Concat_Sheet( _
      destSheetName, _
      srcSheetName, _
      destStartRow, _
      srcStartRow, _
      copyColCount, _
      folderPath _
    )
End Sub


