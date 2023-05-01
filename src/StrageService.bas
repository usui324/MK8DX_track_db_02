Attribute VB_Name = "StrageService"
' Storage�V�[�g�Ɋւ���T�[�r�X�N���X
'
Public Sub initStorage(languageKey As String)
' �X�g���[�W���Z�b�g
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
' ���{��̃R�[�X���X�g���Z�b�g
    
    ' �R�[�X���X�g���R�s�[
    Dim trackNames As Range: Set trackNames = getTrackNameJpList()
    trackNames.Copy
    
    ' �y�[�X�g
    Sheets(STORAGE).Select
    Cells(1, STORAGE_COL_TRACK_NAME).Value = SELECT_TRACK_JP
    Cells(2, STORAGE_COL_TRACK_NAME).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTrackListEn()
' �p��̃R�[�X���X�g���Z�b�g

    ' �R�[�X���X�g���R�s�[
    Dim trackNames As Range: Set trackNames = getTrackNameEnList()
    trackNames.Copy
    
    ' �y�[�X�g
    Sheets(STORAGE).Select
    Cells(1, STORAGE_COL_TRACK_NAME).Value = SELECT_TRACK_EN
    Cells(2, STORAGE_COL_TRACK_NAME).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTrackKeyList()
' �R�[�X�L�[���X�g���Z�b�g

    ' �R�[�X�L�[���X�g���R�s�[
    Dim trackKeys As Range: Set trackKeys = getTrackKeyList()
    trackKeys.Copy
    
    ' �y�[�X�g
    Sheets(STORAGE).Select
    Cells(2, STORAGE_COL_TRACK_KEY).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setLanguages()
' ���ꃊ�X�g���Z�b�g
    
    ' ���ꖼ���X�g���R�s�[
    Dim languageNames As Range: Set languageNames = getLanguageNameList()
    languageNames.Copy Sheets(STORAGE).Cells(1, STORAGE_COL_LANGUAGE_NAME)
    Application.CutCopyMode = False
    
    ' ����L�[���X�g���R�s�[
    Dim languageKeys As Range: Set languageKeys = getLanguageKeyList()
    languageKeys.Copy Sheets(STORAGE).Cells(1, STORAGE_COL_LANGUAGE_KEY)
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setTierList(languageKey As String)
' tier���X�g���Z�b�g
    
    
    Sheets(STORAGE).Cells(1, STORAGE_COL_TIER_NAME).Value = getUnselectValue(languageKey)
    
    Dim tierNames As Range: Set tierNames = getTierNameList()
    Dim pasteCell As Range: Set pasteCell = Sheets(STORAGE).Cells(2, STORAGE_COL_TIER_NAME)
    tierNames.Copy pasteCell
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Sub setFormatList(languageKey As String)
' format���X�g���Z�b�g
    
    Sheets(STORAGE).Cells(1, STORAGE_COL_FORMAT_NAME).Value = getUnselectValue(languageKey)
    
    Dim formatNames As Range: Set formatNames = getFormatNameList()
    Dim pasteCell As Range: Set pasteCell = Sheets(STORAGE).Cells(2, STORAGE_COL_FORMAT_NAME)
    formatNames.Copy pasteCell
    Application.CutCopyMode = False
    
    Range("A1").Select
    
End Sub

Public Function getUnselectValue(languageKey As String) As String
' ���I���̕������擾
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
' �R�[�X�L�[���擾����
' TODO: �݌v�̍��{�I�Ȍ�����, �v���_�E���̕\�����Ɠ����I�Ȓl�𕪂�����
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
