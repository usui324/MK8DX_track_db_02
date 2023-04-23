Attribute VB_Name = "SettingsService"
' Settingsシートに関するサービスクラス
'
 
Function getLanguage() As String
' 言語設定の取得
'
   getLanguage = Sheets(SETTINGS).Cells(SETTINGS_ROW_LANGUAGE, SETTINGS_COL_VALUE).Value
   

End Function
