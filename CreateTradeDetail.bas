Attribute VB_Name = "CreateTradeDetail"
'@Folder("CreateTradeDetail")
Option Explicit

Public Sub �ʎ���쐬()
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
        
        Worksheets("�ʎ��").Select
    Else
        MsgBox "�I�𒆂̍s�ɂ͒����ԍ����܂܂�Ă��܂���B�ēx���s���ĉ������B", vbExclamation, "�G���["
    End If
End Sub

Private Sub CurrentImageDelete()
    Dim sh As Shape
    For Each sh In Worksheets("�ʎ��").Shapes
        If sh.Name <> "GetImage" Then sh.Delete
    Next
End Sub
