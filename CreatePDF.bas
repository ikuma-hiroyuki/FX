Attribute VB_Name = "CreatePDF"
'@folder("CreatePDF")
Option Explicit

Sub CreateTradingDetailPDF()
    If MsgBox("�ʎ����PDF���쐬���܂����H", vbInformation + vbYesNo) = vbYes Then
        Dim dealDetail As TradeDetail
        Set dealDetail = New TradeDetail
        
        If dealDetail.IsTradeExist Then
            dealDetail.CreatePdf
        Else
            MsgBox "����ԍ�: " & dealDetail.OrderNum & " ���}�X�^�[�f�[�^�ɑ��݂��Ă��܂���B", vbExclamation
        End If
    End If
End Sub
