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

Public Function getSelectTrackValue(languageKey As String) As String
' コースを選択の文言を取得
'
    If languageKey = "jp" Then
        getSelectTrackValue = SELECT_TRACK_JP
    ElseIf languageKey = "en" Then
        getSelectTrackValue = SELECT_TRACK_EN
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

Public Function getPointFlg() As Integer
' 得点ソート値を取得
'
    getPointFlg = Sheets(STORAGE).Cells(STORAGE_ROW_POINT_FLG, STORAGE_COL_POINT_FLG).Value
End Function

Public Function getRankFlg() As Integer
' 順位ソート値を取得
'
    getRankFlg = Sheets(STORAGE).Cells(STORAGE_ROW_RANK_FLG, STORAGE_COL_RANK_FLG).Value
End Function

Public Function getTimesFlg() As Integer
' 回数ソート値を取得
'
    getTimesFlg = Sheets(STORAGE).Cells(STORAGE_ROW_TIMES_FLG, STORAGE_COL_TIMES_FLG).Value
End Function

Public Sub incrementPointFlg()
' 得点ソート値を加算
'
    Sheets(STORAGE).Unprotect PASSWORD:=PROTECT_PASSWORD
    
    Sheets(STORAGE).Cells(STORAGE_ROW_POINT_FLG, STORAGE_COL_POINT_FLG).Value = (Sheets(STORAGE).Cells(STORAGE_ROW_POINT_FLG, STORAGE_COL_POINT_FLG).Value + 1) Mod 2
    Sheets(STORAGE).Cells(STORAGE_ROW_RANK_FLG, STORAGE_COL_RANK_FLG).Value = 0
    Sheets(STORAGE).Cells(STORAGE_ROW_TIMES_FLG, STORAGE_COL_TIMES_FLG).Value = 0
    
    Sheets(STORAGE).Protect PASSWORD:=PROTECT_PASSWORD
End Sub

Public Sub incrementRankFlg()
' 順位ソート値を加算
'
    Sheets(STORAGE).Unprotect PASSWORD:=PROTECT_PASSWORD
    
    Sheets(STORAGE).Cells(STORAGE_ROW_POINT_FLG, STORAGE_COL_POINT_FLG).Value = 0
    Sheets(STORAGE).Cells(STORAGE_ROW_RANK_FLG, STORAGE_COL_RANK_FLG).Value = (Sheets(STORAGE).Cells(STORAGE_ROW_RANK_FLG, STORAGE_COL_RANK_FLG).Value + 1) Mod 2
    Sheets(STORAGE).Cells(STORAGE_ROW_TIMES_FLG, STORAGE_COL_TIMES_FLG).Value = 0
    
    Sheets(STORAGE).Protect PASSWORD:=PROTECT_PASSWORD
End Sub

Public Sub incrementTimesFlg()
' 回数ソート値を加算
'
    Sheets(STORAGE).Unprotect PASSWORD:=PROTECT_PASSWORD
    
    Sheets(STORAGE).Cells(STORAGE_ROW_POINT_FLG, STORAGE_COL_POINT_FLG).Value = 0
    Sheets(STORAGE).Cells(STORAGE_ROW_RANK_FLG, STORAGE_COL_RANK_FLG).Value = 0
    Sheets(STORAGE).Cells(STORAGE_ROW_TIMES_FLG, STORAGE_COL_TIMES_FLG).Value = (Sheets(STORAGE).Cells(STORAGE_ROW_TIMES_FLG, STORAGE_COL_TIMES_FLG).Value + 1) Mod 2
    
    Sheets(STORAGE).Protect PASSWORD:=PROTECT_PASSWORD
End Sub
