Attribute VB_Name = "CreateMenu"
'@Folder("Utility")
Option Explicit

Sub CreateRightClickMenu()
    Select Case ActiveSheet.Name
        Case "個別取引"
            CreateDetailPdfMenu
    End Select
End Sub

Private Sub CreateDetailPdfMenu()
    CommandBars("Cell").Reset
    Dim rightClickMenu As CommandBarButton
    Set rightClickMenu = CommandBars("Cell").Controls.Add(Before:=1, Temporary:=True)
    With rightClickMenu
        .Caption = "個別取引PDF作成"
        .OnAction = "CreateTradingDetailPDF"
        .Tag = "Documentフォルダ内に個別取引PDFを作成します"
        .State = msoButtonDown
    End With
End Sub
