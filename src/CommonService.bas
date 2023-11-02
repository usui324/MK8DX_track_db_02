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

Public Function findWholeMatch(r As Range, Target As Variant) As Range
' Range�I�u�W�F�N�g���犮�S��v����I�u�W�F�N�g��T������
' @param r �T����Range�I�u�W�F�N�g
' @param target �T���Ώە�����

    Set findWholeMatch = r.Find(Target, LookAt:=xlWhole, MatchCase:=True)
    
End Function

Public Sub goToRegistDataSheet()
' �f�[�^�o�^�V�[�g�ֈړ�
    Application.ScreenUpdating = False
    
    ' �V�[�g�I��
    selectSheet (REGIST_DATA)
    ' �E�B���h�E�T�C�Y�̒���
    Application.WindowState = xlNormal
    ActiveWindow.Width = 470
    ActiveWindow.Height = 720
    ' �Z���̑I��
    Range(INIT_SELECT_REGIST_DATA).Select
    

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
    ' �Z���̑I��
    Range(INIT_SELECT_DATA).Select
    
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
    ' �Z���̑I��
    Range(INIT_SELECT_GRAPH).Select

    Application.ScreenUpdating = True
End Sub

Public Sub goToSettingsSheet()
' �ݒ�V�[�g�ֈړ�
'
    Application.ScreenUpdating = False
    
    ' �V�[�g�I��
    selectSheet (SETTINGS)
    ' �E�B���h�E�T�C�Y�̒���
    Application.WindowState = xlMaximized
    ' �Z���̑I��
    Range(INIT_SELECT_SETTINGS).Select

    Application.ScreenUpdating = True
End Sub

Public Function isCorrectArray(ByVal arrs As Variant) As Boolean
' �z�񂪐��킩���肷��
'
    isCorrectArray = True
    
    ' �ő�C���f�b�N�X���擾
    Dim a As Long
    On Error GoTo err
    a = UBound(arrs)
    
    ' �C���f�b�N�X�������Ȃ�False
    If a < 0 Then
        isCorrectArray = False
    End If
    
err:
    '�G���[���������Ƃ��G���[�ԍ���9��13�̏ꍇ��False
    If err.Number = 9 Or err.Number = 13 Then
        isCorrectArray = False
    End If
    
End Function

Public Function convertLongToStr(longNum As Long, strSize As Integer) As String
' ���l�𕶎���ɕϊ�����
'
    Dim l As Integer: l = Len(CStr(longNum))
    
    If l >= strSize Then
        convertLongToStr = CStr(longNum)
        Exit Function
    End If
    
    Dim i As Integer
    For i = l + 1 To strSize
        convertLongToStr = convertLongToStr + "0"
    Next i
    
    convertLongToStr = convertLongToStr + CStr(longNum)
End Function

Public Sub saveBook()
' �ۑ�����
'
    ThisWorkbook.Save
End Sub
