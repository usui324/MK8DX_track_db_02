Attribute VB_Name = "SettingsService"
' Settings�V�[�g�Ɋւ���T�[�r�X�N���X
'
 
Function getLanguage() As String
' ����ݒ�̎擾
'
   getLanguage = Sheets(SETTINGS).Cells(SETTINGS_ROW_LANGUAGE, SETTINGS_COL_VALUE).Value
   
End Function

Function getShowingMap() As String
' �}�b�v�\���ݒ�̎擾
'
    getShowingMap = Sheets(SETTINGS).Cells(SETTINGS_ROW_SHOWING_MAP, SETTINGS_COL_VALUE).Value
    
End Function
