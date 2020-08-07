VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
e '@Folder("Orders")
Option Explicit

Private m_orderNum As String
Private m_wsMasta As Worksheet
Private m_entryURL As String
Private m_exitURL As String
Private m_2HourLaterURL As String

Private m_currentRow As Long

Public Sub GetOrder()
    If ActiveSheet.Name = "マスターデータ" Then
        Set m_wsMasta = Worksheets("マスターデータ")
        With m_wsMasta
            If IsNumeric(.Cells(m_currentRow, "A")) Then
                m_orderNum = .Cells(m_currentRow, "A")
                m_entryURL = .Cells(m_currentRow, "N")
                m_exitURL = .Cells(m_currentRow, "V")
                m_2HourLaterURL = .Cells(m_currentRow, "W")
            End If
        End With
    Else
        MsgBox "マスターデータシートで実行して下さい。", vbExclamation
        End
    End If
End Sub

Private Sub Class_Initialize()
    m_currentRow = Selection.Row
End Sub

Public Property Get EntryURL() As String
    EntryURL = m_entryURL
End Property

Public Property Get ExitURL() As String
    ExitURL = m_exitURL
End Property

Public Property Get TwoHourLaterURL() As String
    TwoHourLaterURL = m_2HourLaterURL
End Property

Public Property Get OrderNum() As String
    OrderNum = m_orderNum
End Property
