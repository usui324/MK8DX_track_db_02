Attribute VB_Name = "SettingsService"
' Settingsシートに関するサービスクラス
'
 
Function getLanguage() As String
' 言語設定の取得
'
   getLanguage = Sheets(SETTINGS).Cells(SETTINGS_ROW_LANGUAGE, SETTINGS_COL_VALUE).Value
   
End Function

Function getShowingMap() As String
' マップ表示設定の取得
'
    getShowingMap = Sheets(SETTINGS).Cells(SETTINGS_ROW_SHOWING_MAP, SETTINGS_COL_VALUE).Value
    
End Function
