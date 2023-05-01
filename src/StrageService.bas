Attribute VB_Name = "StrageService"
' Storageシートに関するサービスクラス
'
Public Sub initStorage(languageKey As String)
' ストレージをセット
'
    Sheets(STORAGE).Unprotect PASSWORD:=PROTECT_PASSWORD

    If languageKey = "jp" Then
        setTrackListJp
        setTrackKeyList
        setLanguages
        setTierList (languageKey)
        setFormatList (languageKey)
    ElseIf languageKey = "en" Then
        setTrackListEn
        setTrackKeyList
        setLanguages
        setTierList (languageKey)
        setFormatList (languageKey)
    Else
        
    End If
    
    Sheets(STORAGE).Protect PASSWORD:=PROTECT_PASSWORD
End Sub

Public Sub setTrackListJp()
' 日本語のコースリストをセット
    
    ' コースリストをコピー
    Dim trackNames As Range: Set trackNames = getTrackNameJpList()
    trackNames.Copy
    
    ' ペースト
    Sheets(STORAGE).Select
    Cells(1, STORAGE_COL_TRACK_NAME).Value = SELECT_TRACK_JP
    Cells(2, STORAGE_COL_TRACK_NAME).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTrackListEn()
' 英語のコースリストをセット

    ' コースリストをコピー
    Dim trackNames As Range: Set trackNames = getTrackNameEnList()
    trackNames.Copy
    
    ' ペースト
    Sheets(STORAGE).Select
    Cells(1, STORAGE_COL_TRACK_NAME).Value = SELECT_TRACK_EN
    Cells(2, STORAGE_COL_TRACK_NAME).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTrackKeyList()
' コースキーリストをセット

    ' コースキーリストをコピー
    Dim trackKeys As Range: Set trackKeys = getTrackKeyList()
    trackKeys.Copy
    
    ' ペースト
    Sheets(STORAGE).Select
    Cells(2, STORAGE_COL_TRACK_KEY).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setLanguages()
' 言語リストをセット
    
    ' 言語名リストをコピー
    Dim languageNames As Range: Set languageNames = getLanguageNameList()
    languageNames.Copy Sheets(STORAGE).Cells(1, STORAGE_COL_LANGUAGE_NAME)
    Application.CutCopyMode = False
    
    ' 言語キーリストをコピー
    Dim languageKeys As Range: Set languageKeys = getLanguageKeyList()
    languageKeys.Copy Sheets(STORAGE).Cells(1, STORAGE_COL_LANGUAGE_KEY)
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTierList(languageKey As String)
' tierリストをセット
    
    
    Sheets(STORAGE).Cells(1, STORAGE_COL_TIER_NAME).Value = getUnselectValue(languageKey)
    
    Dim tierNames As Range: Set tierNames = getTierNameList()
    Dim pasteCell As Range: Set pasteCell = Sheets(STORAGE).Cells(2, STORAGE_COL_TIER_NAME)
    tierNames.Copy pasteCell
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setFormatList(languageKey As String)
' formatリストをセット
    
    Sheets(STORAGE).Cells(1, STORAGE_COL_FORMAT_NAME).Value = getUnselectValue(languageKey)
    
    Dim formatNames As Range: Set formatNames = getFormatNameList()
    Dim pasteCell As Range: Set pasteCell = Sheets(STORAGE).Cells(2, STORAGE_COL_FORMAT_NAME)
    formatNames.Copy pasteCell
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Function getUnselectValue(languageKey As String) As String
' 未選択の文言を取得
'
    If languageKey = "jp" Then
        getUnselectValue = UNSELECT_JP
    ElseIf languageKey = "en" Then
        getUnselectValue = UNSELECT_EN
    Else
        openErrorMsgBox ("invalid languageKey: " + language)
    End If
End Function

Public Function getTrackKey(trackName As String) As String
' コースキーを取得する
' TODO: 設計の根本的な見直し, プルダウンの表示名と内部的な値を分けたい
'
    Dim startCell As Range: Set startCell = Sheets(STORAGE).Cells(STORAGE_ROW_TRACK_NAME, STORAGE_COL_TRACK_NAME)
    Dim endCell As Range: Set endCell = Sheets(STORAGE).Cells(STORAGE_ROW_TRACK_NAME, STORAGE_COL_TRACK_NAME).End(xlDown)
    Dim keyCell As Range: Set keyCell = Sheets(STORAGE).Range(startCell, endCell).Find(trackName)
    
    If keyCell Is Nothing Then
        getTrackKey = ""
    Else
        getTrackKey = Sheets(STORAGE).Cells(keyCell.Row, STORAGE_COL_TRACK_KEY).Value
    End If
        
End Function
