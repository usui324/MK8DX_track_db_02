Attribute VB_Name = "MasterService"
Option Explicit

Public Function getTrackNameJpList() As Range
' コースの日本語名のリストを取得する
    
    Set getTrackNameJpList = getMasterColumn(TRACK_MASTER, "trackNameJp")
    
End Function

Public Function getTrackNameEnList() As Range
' コースの英語名のリストを取得する
    
    Set getTrackNameEnList = getMasterColumn(TRACK_MASTER, "trackNameEn")
    
End Function

Public Function getTrackKeyList() As Range
' コースキーリストを取得する
    
    Set getTrackKeyList = getMasterColumn(TRACK_MASTER, "trackKey")

End Function

Public Function getLanguageNameList() As Range
' 言語名リストを取得する
    
    Set getLanguageNameList = getMasterColumn(LANGUAGE_MASTER, "languageName")

End Function

Public Function getLanguageKeyList() As Range
' 言語キーリストを取得する
    
    Set getLanguageKeyList = getMasterColumn(LANGUAGE_MASTER, "languageKey")

End Function

Public Function getTrackNameJp(trackKey As String) As String
' コースの日本名を取得する
'
    Dim recordRow As Long: recordRow = getMasterRecord(TRACK_MASTER, trackKey).Row
    Dim col As Long: col = getMasterColumn(TRACK_MASTER, "trackNameJp").column
    
    getTrackNameJp = Sheets(TRACK_MASTER).Cells(recordRow, col).Value
    
End Function

Public Function getTrackNameEn(trackKey As String) As String
' コースの英名を取得する
'
    Dim recordRow As Long: recordRow = getMasterRecord(TRACK_MASTER, trackKey).Row
    Dim col As Long: col = getMasterColumn(TRACK_MASTER, "trackNameEn").column
    getTrackNameEn = Sheets(TRACK_MASTER).Cells(recordRow, col).Value
    
End Function

Public Function getTierNameList() As Range
' tier名リストを取得する

    Set getTierNameList = getMasterColumn(LOUNGE_TIER_MASTER, "loungeTierName")

End Function

Public Function getFormatNameList() As Range
' フォーマット名リストを取得する

    Set getFormatNameList = getMasterColumn(FORMAT_MASTER, "formatName")

End Function

Public Function getPoint(rankKey As Integer) As Integer
' 得点を取得する
'
    Dim recordRow As Long: recordRow = getMasterRecord(POINT_MASTER, rankKey).Row
    Dim col As Long: col = getMasterColumn(POINT_MASTER, "Point").column
    getPoint = Sheets(POINT_MASTER).Cells(recordRow, col).Value
    
End Function

