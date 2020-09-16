Attribute VB_Name = "CreateMenu"
'@Folder("Utility")
Option Explicit

Sub CreateRightClickMenu()
    Select Case ActiveSheet.Name
        Case "ŒÂ•Êæˆø"
            CreateDetailPdfMenu
    End Select
End Sub

Private Sub CreateDetailPdfMenu()
    CommandBars("Cell").Reset
    Dim rightClickMenu As CommandBarButton
    Set rightClickMenu = CommandBars("Cell").Controls.Add(Before:=1, Temporary:=True)
    With rightClickMenu
        .Caption = "ŒÂ•ÊæˆøPDFì¬"
        .OnAction = "CreateTradingDetailPDF"
        .Tag = "DocumentƒtƒHƒ‹ƒ_“à‚ÉŒÂ•ÊæˆøPDF‚ğì¬‚µ‚Ü‚·"
        .State = msoButtonDown
    End With
End Sub
