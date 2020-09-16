VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TradeDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CreatePDF")
Option Explicit

Private m_OrderNum As String
Private m_Currency As String
Private m_EntryDateTime As Date
Private m_WsDetail As Worksheet
Private m_WsMasta As Worksheet
Private m_IsTradeExist As Boolean

Public Sub CreatePdf()
    Dim wsh As Object:    Set wsh = CreateObject("WScript.Shell")
    Dim saveDir As String: saveDir = wsh.SpecialFolders("MyDocuments")
    Dim fileName As String: fileName = Format(m_EntryDateTime, "yyyy-mm-dd-hhmm") & "-" & m_Currency
    Dim savePath As String: savePath = saveDir & "\" & Replace(fileName, "/", "_") & ".pdf"
    
    m_WsDetail.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=savePath, _
        OpenAfterPublish:=True
End Sub

Private Sub Class_Initialize()
    Set m_WsDetail = Worksheets("個別取引")
    Set m_WsMasta = Worksheets("マスターデータ")
    m_OrderNum = m_WsDetail.Cells(1, "J")
    Dim foundOrder As Range:    Set foundOrder = m_WsMasta.Columns(1).Find(what:=m_OrderNum)
    If Not foundOrder Is Nothing Then
        m_Currency = m_WsMasta.Cells(foundOrder.Row, Col.通貨ペア)
        m_EntryDateTime = m_WsMasta.Cells(foundOrder.Row, Col.エントリー日時)
        m_IsTradeExist = True
    End If
End Sub

Public Property Get IsTradeExist() As Boolean
    IsTradeExist = m_IsTradeExist
End Property

Public Property Get OrderNum() As String
    OrderNum = m_OrderNum
End Property
