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

' 変数
    ' コースキー
    Private pTrackKey As String
    ' コース日本名
    Private pTrackNameJp As String
    ' コース英名
    Private pTrackNameEn As String
    ' スタート順位
    Private pStartingRank As Integer
    ' 結果順位
    Private pResultRank As Integer
    ' 得点
    Private pResultPoint As Integer
    ' 備考
    Private pRemark As String

' プロパティ
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

' メソッド
    Public Function init(trackKey As String, startingRank As Integer, resultRank As Integer, remark As String)
    ' 初期化
    '
        pTrackKey = trackKey
        pTrackNameJp = getTrackNameJp(trackKey)
        pTrackNameEn = getTrackNameEn(trackKey)
        pStartingRank = startingRank
        pResultRank = resultRank
        pResultPoint = getPoint(resultRank)
        pRemark = remark
    End Function

