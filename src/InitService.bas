Attribute VB_Name = "InitService"
' �������Ɋւ���T�[�r�X
'
Public Sub init()
' ����������
'
    ' ���b�Z�[�W�̕\��
    Dim res As VbMsgBoxResult: res = openMsgBox("���͂�����������܂��B��낵���ł��傤���H", , vbOKCancel)
    If res = VbMsgBoxResult.vbCancel Then
        Exit Sub
    End If
    
    ' ����ݒ�
    Dim language As String: language = getLanguage

    ' Storage�̏�����
    initStorage (language)
    
    ' RegistData�̏�����
    initInputData
    
End Sub

