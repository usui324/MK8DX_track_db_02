Attribute VB_Name = "Liblary"
Sub ExportAll()
' ���W���[����S�ăG�N�X�|�[�g����

    ' ���W���[��
    Dim module As VBComponent
    Dim moduleList As VBComponents
    
    ' �g���q
    Dim extension
    ' �Ώۃu�b�N�̃p�X
    Dim targetPath
    ' �G�N�X�|�[�g�t�@�C���p�X
    Dim exportPath
    ' �Ώۃu�b�N�I�u�W�F�N�g
    Dim targetBook


    ' ���̃u�b�N��ΏۂƂ���
    Set targetBook = ThisWorkbook
    targetPath = targetBook.Path

    ' ���W���[���ꗗ���擾
    Set moduleList = targetBook.VBProject.VBComponents
    
    ' �e���W���[���ɑ΂��鏈��
    For Each module In moduleList
        ' �N���X
        If module.Type = vbext_ct_ClassModule Then
            extension = "cls"
        ' �t�H�[��
        ElseIf module.Type = vbext_ct_MSForm Then
            extension = "frm"
        ' �W�����W���[��
        ElseIf module.Type = vbext_ct_StdModule Then
            extension = "bas"
        ' ���̑�
        Else
            GoTo CONTINUE
        End If
        
        ' �G�N�X�|�[�g����
        exportPath = targetPath & "\src\" & module.Name & "." & extension
        Call module.Export(exportPath)
        
        ' �o�͐�m�F�p���O
        Debug.Print exportPath
        
CONTINUE:
    Next

End Sub

