VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �ϐ�
    ' �R�[�X�L�[
    Private pTrackKey As String
    ' �R�[�X���{��
    Private pTrackNameJp As String
    ' �R�[�X�p��
    Private pTrackNameEn As String
    ' �X�^�[�g����
    Private pStartingRank As Integer
    ' ���ʏ���
    Private pResultRank As Integer
    ' ���_
    Private pResultPoint As Integer
    ' ���l
    Private pRemark As String

' �v���p�e�B
    Public Property Get trackKey() As String
        trackKey = pTrackKey
    End Property
    
    Public Property Get trackNameJp() As String
        trackNameJp = pTrackNameJp
    End Property
    
    Public Property Get trackNameEn() As String
        trackNameEn = pTrackNameEn
    End Property
    
    Public Property Get startingRank() As Integer
        startingRank = pStartingRank
    End Property
    
    Public Property Get resultRank() As Integer
        resultRank = pResultRank
    End Property
    
    Public Property Get resultPoint() As Integer
        resultPoint = pResultPoint
    End Property
    
    Public Property Get remark() As String
        remark = pRemark
    End Property

' ���\�b�h
    Public Function init(trackKey As String, startingRank As Integer, resultRank As Integer, remark As String)
    ' ������
    '
        pTrackKey = trackKey
        pTrackNameJp = getTrackNameJp(trackKey)
        pTrackNameEn = getTrackNameEn(trackKey)
        pStartingRank = startingRank
        pResultRank = resultRank
        pResultPoint = getPoint(resultRank)
        pRemark = remark
    End Function
