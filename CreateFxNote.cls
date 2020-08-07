VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateFxNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '@Folder("GetWebImage")
Option Explicit

#If Win64 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Private m_wsIndividual As Worksheet

'�摜��URL
Private m_imageUrl  As String
'�摜�擾���ʂ�ۑ�
Private m_DLResult As Long
'�摜�ۑ���
Private m_imgSavePath As String
'�摜�����A�h���X
Private m_implantAddress As String

Private m_filename As String
Private m_orderNum As String

Private Sub Class_Initialize()
    Set m_wsIndividual = Worksheets("�ʎ��")
End Sub

Private Sub Class_Terminate()
    Application.StatusBar = False
    On Error Resume Next
    Kill m_imgSavePath
End Sub

Private Sub ExeImplantation(ByVal result As Long, ByVal imagePath As String, implantRange As Range)
    If result = 0 Then
        With m_wsIndividual.Pictures.Insert(imagePath)
            m_wsIndividual.Range("J1") = m_orderNum
            .Top = implantRange.Top
            .Left = implantRange.Left
            .Width = implantRange.MergeArea.Width - 5
        End With
        implantRange.Value = vbNullString
    Else
        implantRange.Value = "�摜�_�E�����[�h�Ɏ��s���܂����B"
    End If
End Sub

Public Sub GetImage(ByVal OrderNum As String, _
                    ByVal imageUrl As String, _
                    ByVal fileName As String, _
                    ByVal implantAddress As String)
                    
    m_orderNum = OrderNum
    m_filename = fileName
    Application.StatusBar = m_filename & "�摜���擾��"
    m_implantAddress = implantAddress
    m_imageUrl = imageUrl
    ImageDl fileName, m_DLResult, m_imageUrl, m_imgSavePath
End Sub

Private Sub ImageDl(ByVal fileType As String, ByRef result As Long, ByVal imageUrl As String, ByRef saveAs As String)
    saveAs = ThisWorkbook.Path & "\" & m_orderNum & "_" & fileType & ".png"
    result = URLDownloadToFile(0, imageUrl, saveAs, 0, 0) '�_�E�����[�h�ɐ��������0���Ԃ�B
End Sub

Public Sub ImplantationImage()
    Dim implantRange As Range: Set implantRange = m_wsIndividual.Range(m_implantAddress)
    ExeImplantation m_DLResult, m_imageUrl, implantRange
End Sub
