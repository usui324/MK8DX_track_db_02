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

' 変数
    ' 登録キー
    Private pRegistKey As String
    ' 日付
    Private pPlayDate As Date
    ' tier
    Private pTier As String
    ' 形式
    Private pFormat As String
    ' コースデータ
    Private pTrackData() As New TrackData

' プロパティ
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

' メソッド
    Public Sub init(registKey As String, tier As String, format As String, playDate As Date, trackDataList() As TrackData)
    ' 初期化
    '
        pRegistKey = registKey
        pTier = tier
        pFormat = format
        pPlayDate = playDate
        pTrackData = trackDataList
    End Sub
