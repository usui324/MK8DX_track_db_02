VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �ϐ�
    ' �o�^�L�[
    Private pRegistKey As String
    ' ���t
    Private pPlayDate As Date
    ' tier
    Private pTier As String
    ' �`��
    Private pFormat As String
    ' �R�[�X�f�[�^
    Private pTrackData() As New TrackData

' �v���p�e�B
    Public Property Get registKey() As String
        registKey = pRegistKey
    End Property
    
    Public Property Get playDate() As Date
        playDate = pPlayDate
    End Property
    
    Public Property Get tier() As String
        tier = pTier
    End Property
    
    Public Property Get format() As String
        format = pFormat
    End Property
    
    Public Property Get TrackData() As TrackData()
        TrackData = pTrackData
    End Property

' ���\�b�h
    Public Sub init(registKey As String, tier As String, format As String, playDate As Date, trackDataList() As TrackData)
    ' ������
    '
        pRegistKey = registKey
        pTier = tier
        pFormat = format
        pPlayDate = playDate
        pTrackData = trackDataList
    End Sub