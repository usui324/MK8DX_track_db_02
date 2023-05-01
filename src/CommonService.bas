Attribute VB_Name = "CommonService"
' �ėp�T�[�r�X�N���X
'

Public Sub selectSheet(sheetName As String)
' �V�[�g��I������
' @param sheetName: �V�[�g��
    
On Error GoTo Exception
    
    Range("A1").Select
    
    Sheets(sheetName).Select
    
    Range("A1").Select
    
    Exit Sub
Exception:

    Call openErrorMsgBox("invalid sheetName: " & sheetName)

End Sub

Public Sub openErrorMsgBox(message As String)
' �G���[���b�Z�[�W��\������
' @param message: ���b�Z�[�W
    
   Call openMsgBox(message, "Error")
   End

End Sub

Public Function openMsgBox(message As String, Optional title As String = "MK8DX Track DB", Optional style As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly) As VbMsgBoxResult

' ���b�Z�[�W�{�b�N�X��\������
' @param message: ���b�Z�[�W���e
' @param title: �^�C�g��
' @param style: ���b�Z�[�W�{�b�N�X�̃X�^�C��

    openMsgBox = MsgBox(message, style, title)

End Function

Public Function findWholeMatch(r As Range, target As Variant) As Range
' Range�I�u�W�F�N�g���犮�S��v����I�u�W�F�N�g��T������
' @param r �T����Range�I�u�W�F�N�g
' @param target �T���Ώە�����

    Set findWholeMatch = r.Find(target, LookAt:=xlWhole, MatchCase:=True)
    
End Function

Public Sub goToRegistDataSheet()
' �f�[�^�o�^�V�[�g�ֈړ�
    Application.ScreenUpdating = False
    
    ' �V�[�g�I��
    selectSheet (REGIST_DATA)
    ' �E�B���h�E�T�C�Y�̒���
    Application.WindowState = xlNormal
    ActiveWindow.Width = 430
    ActiveWindow.Height = 720

    Application.ScreenUpdating = True
End Sub

Public Sub goToDataSheet()
' �f�[�^�V�[�g�ֈړ�
'
    Application.ScreenUpdating = False
    
    ' �V�[�g�I��
    selectSheet (DATA)
    ' �E�B���h�E�T�C�Y�̒���
    Application.WindowState = xlMaximized

    Application.ScreenUpdating = True
End Sub

Public Sub goToGraphSheet()
' �O���t�V�[�g�ֈړ�
'
    Application.ScreenUpdating = False
    
    ' �V�[�g�I��
    selectSheet (GRAPH)
    ' �E�B���h�E�T�C�Y�̒���
    Application.WindowState = xlMaximized

    Application.ScreenUpdating = True
End Sub