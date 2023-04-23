Attribute VB_Name = "MasterService"
Option Explicit

Public Function getTrackNameJpList() As Range
' �R�[�X�̓��{�ꖼ�̃��X�g���擾����
    
    Set getTrackNameJpList = getMasterColumn(TRACK_MASTER, "trackNameJp")
    
End Function

Public Function getTrackNameEnList() As Range
' �R�[�X�̉p�ꖼ�̃��X�g���擾����
    
    Set getTrackNameEnList = getMasterColumn(TRACK_MASTER, "trackNameEn")
    
End Function

Public Function getTrackKeyList() As Range
' �R�[�X�L�[���X�g���擾����
    
    Set getTrackKeyList = getMasterColumn(TRACK_MASTER, "trackKey")

End Function

Public Function getLanguageNameList() As Range
' ���ꖼ���X�g���擾����
    
    Set getLanguageNameList = getMasterColumn(LANGUAGE_MASTER, "languageName")

End Function

Public Function getLanguageKeyList() As Range
' ����L�[���X�g���擾����
    
    Set getLanguageKeyList = getMasterColumn(LANGUAGE_MASTER, "languageKey")

End Function

Public Function getTrackNameJp(trackKey As String) As String
' �R�[�X�̓��{�����擾����
'
    Dim recordRow As Long: recordRow = getMasterRecord(TRACK_MASTER, trackKey).Row
    Dim col As Long: col = getMasterColumn(TRACK_MASTER, "trackNameJp").column
    
    getTrackNameJp = Sheets(TRACK_MASTER).Cells(recordRow, col).Value
    
End Function

Public Function getTrackNameEn(trackKey As String) As String
' �R�[�X�̉p�����擾����
'
    Dim recordRow As Long: recordRow = getMasterRecord(TRACK_MASTER, trackKey).Row
    Dim col As Long: col = getMasterColumn(TRACK_MASTER, "trackNameEn").column
    getTrackNameEn = Sheets(TRACK_MASTER).Cells(recordRow, col).Value
    
End Function

Public Function getTierNameList() As Range
' tier�����X�g���擾����

    Set getTierNameList = getMasterColumn(LOUNGE_TIER_MASTER, "loungeTierName")

End Function

Public Function getFormatNameList() As Range
' �t�H�[�}�b�g�����X�g���擾����

    Set getFormatNameList = getMasterColumn(FORMAT_MASTER, "formatName")

End Function

Public Function getPoint(rankKey As Integer) As Integer
' ���_���擾����
'
    Dim recordRow As Long: recordRow = getMasterRecord(POINT_MASTER, rankKey).Row
    Dim col As Long: col = getMasterColumn(POINT_MASTER, "Point").column
    getPoint = Sheets(POINT_MASTER).Cells(recordRow, col).Value
    
End Function

