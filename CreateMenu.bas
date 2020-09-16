Attribute VB_Name = "CreateMenu"
'@Folder("Utility")
Option Explicit

Sub CreateRightClickMenu()
    Select Case ActiveSheet.Name
        Case "�ʎ��"
            CreateDetailPdfMenu
    End Select
End Sub

Private Sub CreateDetailPdfMenu()
    CommandBars("Cell").Reset
    Dim rightClickMenu As CommandBarButton
    Set rightClickMenu = CommandBars("Cell").Controls.Add(Before:=1, Temporary:=True)
    With rightClickMenu
        .Caption = "�ʎ��PDF�쐬"
        .OnAction = "CreateTradingDetailPDF"
        .Tag = "Document�t�H���_���Ɍʎ��PDF���쐬���܂�"
        .State = msoButtonDown
    End With
End Sub
