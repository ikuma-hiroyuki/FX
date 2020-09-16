Attribute VB_Name = "CreatePDF"
'@folder("CreatePDF")
Option Explicit

Sub CreateTradingDetailPDF()
    If MsgBox("個別取引のPDFを作成しますか？", vbInformation + vbYesNo) = vbYes Then
        Dim dealDetail As TradeDetail
        Set dealDetail = New TradeDetail
        
        If dealDetail.IsTradeExist Then
            dealDetail.CreatePdf
        Else
            MsgBox "取引番号: " & dealDetail.OrderNum & " がマスターデータに存在していません。", vbExclamation
        End If
    End If
End Sub
