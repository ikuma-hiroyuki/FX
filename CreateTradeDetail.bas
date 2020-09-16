Attribute VB_Name = "CreateTradeDetail"
'@Folder("CreateTradeDetail")
Option Explicit

Public Sub 個別取引作成()
    Dim order As Orders: Set order = New Orders
    order.GetOrder
    
    If order.OrderNum <> vbNullString Then
        CurrentImageDelete
        
        Dim entryFX  As CreateFxNote: Set entryFX = New CreateFxNote
        entryFX.GetImage order.OrderNum, order.EntryURL, "entry", implantAddress:="A7"
        entryFX.ImplantationImage
        
        Dim exitFX As CreateFxNote: Set exitFX = New CreateFxNote
        exitFX.GetImage order.OrderNum, order.ExitURL, "exit", implantAddress:="A31"
        exitFX.ImplantationImage
    
        Dim twoLHFX As CreateFxNote: Set twoLHFX = New CreateFxNote
        twoLHFX.GetImage order.OrderNum, order.TwoHourLaterURL, "2hoursLater", implantAddress:="F31"
        twoLHFX.ImplantationImage
        
        Worksheets("個別取引").Select
    Else
        MsgBox "選択中の行には注文番号が含まれていません。再度実行して下さい。", vbExclamation, "エラー"
    End If
End Sub

Private Sub CurrentImageDelete()
    Dim sh As Shape
    For Each sh In Worksheets("個別取引").Shapes
        If sh.Name <> "GetImage" Then sh.Delete
    Next
End Sub
