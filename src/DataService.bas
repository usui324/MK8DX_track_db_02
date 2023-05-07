Attribute VB_Name = "DataService"
Public Sub addNewData(newData As inputData)
' �V�K�f�[�^�̒ǉ�
'
    ' �s�ԍ�
    Dim rowNo As Long: rowNo = getNewRowNo
    
    If Not isCorrectArray(newData.TrackData) Then
        Call openMsgBox("�L���ȓo�^�f�[�^������܂���B", , vbOKOnly)
        End
    End If
    
    For Each track In newData.TrackData
        Sheets(DATA).Cells(rowNo, DATA_COL_REGIST_KEY).Value = newData.registKey
        Sheets(DATA).Cells(rowNo, DATA_COL_DATE).Value = newData.playDate
        Sheets(DATA).Cells(rowNo, DATA_COL_TIER).Value = newData.tier
        Sheets(DATA).Cells(rowNo, DATA_COL_FORMAT).Value = newData.format
        Sheets(DATA).Cells(rowNo, DATA_COL_TRACK_KEY).Value = track.trackKey
        Sheets(DATA).Cells(rowNo, DATA_COL_TRACK_NAME_JP).Value = track.trackNameJp
        Sheets(DATA).Cells(rowNo, DATA_COL_TRACK_NAME_EN).Value = track.trackNameEn
        Sheets(DATA).Cells(rowNo, DATA_COL_STARTING_RANK).Value = track.startingRank
        Sheets(DATA).Cells(rowNo, DATA_COL_RANK).Value = track.resultRank
        Sheets(DATA).Cells(rowNo, DATA_COL_POINT).Value = track.resultPoint
        Sheets(DATA).Cells(rowNo, DATA_COL_REMARK).Value = track.remark
        
        rowNo = rowNo + 1
    Next
    
End Sub

Public Function getNewRowNo() As Long
' �f�[�^�ǉ��p�̍s�ԍ����擾
'
    getNewRowNo = getLastRowNo + 1
End Function

Public Function getLastRowNo() As Long
' ���̓f�[�^�̍ŏI�s�̍s�ԍ����擾
'
    Dim lastRowCell As Range: Set lastRowCell = Sheets(DATA).Cells(Rows.Count, 1).End(xlUp)
    
    ' ��O����
    If lastRowCell.Row = DATA_ROW_HEADER + 1 And lastRowCell.Value = "" Then
        getLastRowNo = DATA_ROW_HEADER
    Else
        getLastRowNo = lastRowCell.Row
    End If

End Function

Public Function getNewRegistKey() As Long
' �V�K�o�^�L�[���擾
'
    getNewRegistKey = getLastRegistKey + 1
    
    If getNewRegistKey > REGIST_KEY_MAX Then
        Call openMsgBox("����ȏ�̃f�[�^�o�^���󂯕t���܂���B")
        End
    End If
   
End Function

Public Function getLastRegistKey() As Long
' ���̓f�[�^�̍ŐV�̓o�^�L�[���擾
'
    Dim rowNo As Long: rowNo = getLastRowNo
    If rowNo = 1 Then
        getLastRegistKey = 0
    Else
        getLastRegistKey = Sheets(DATA).Cells(rowNo, DATA_COL_REGIST_KEY).Value
    End If
End Function

Public Sub exportData()
' �f�[�^���G�N�X�|�[�g
'
     ' �G�N�X�|�[�g�t�@�C�����w��
    ChDir ThisWorkbook.Path
    Dim saveFileName As String
    saveFileName = Application.GetSaveAsFilename(InitialFileName:="mogiData.txt", filefilter:="�͋[�f�[�^,*.txt")

    ' �L�����Z������
    If saveFileName = "False" Then
        Exit Sub
    End If
    
    ' �o�͂���ΏۃV�[�g
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DATA)

    ' �t�@�C���ɏ�������
    Open saveFileName For Output As #1
    
    Dim i As Long
    For i = 2 To getLastRowNo
        Print #1, printLine(ws, i)
    Next i
    
    Close #1
    
    Call openMsgBox(saveFileName & "�Ƀf�[�^���o�͂��܂���", , vbInformation)

End Sub

Private Function printLine(ws As Worksheet, rowNo As Long) As String
' �t�@�C���o�͂����s�̕������Ԃ�
'
    Dim i As Integer
    printLine = ws.Cells(rowNo, 1).Value
    For i = 2 To DATA_COLS
        printLine = printLine & "," & ws.Cells(rowNo, i).Value
    Next i
    
End Function

Public Sub importData()
' �f�[�^���C���|�[�g
'
    Dim openFileName As String
    Dim ws As Worksheet
    Dim line As String
    Dim arr As Variant
    Dim rowNo As Integer: rowNo = 2
    Dim columnNo As Integer

    ' �C���|�[�g�t�@�C�����w��
    ChDir ThisWorkbook.Path
    openFileName = Application.GetOpenFilename("�͋[�f�[�^,*.txt", , "�C���|�[�g����f�[�^�t�@�C�����w��")
    
    ' �L�����Z������
    If openFileName = "False" Then
        Exit Sub
    End If
    
    ' ���͑ΏۃV�[�g
    Set ws = ThisWorkbook.Worksheets(DATA)
    
    Open openFileName For Input As #1
    
    While Not EOF(1)
        Line Input #1, line
        arr = Split(line, ",")
        
        For columnNo = LBound(arr) To UBound(arr)
            ws.Cells(rowNo, columnNo + 1).Value = arr(columnNo)
        Next columnNo
        rowNo = rowNo + 1
    Wend
    
    Close #1
    
    Call openMsgBox("�f�[�^���C���|�[�g���܂���", , vbInformation)
    
End Sub

Public Sub deleteData()
' �f�[�^���폜����
'
    ' ���b�Z�[�W�\��
    Dim response As Integer
    response = openMsgBox("�f�[�^�����S�폜���܂����H", , vbYesNo + vbDefaultButton2)
    
    If response <> 6 Then
        End
    End If
    
    response = openMsgBox("�{���ɂ�낵���ł����H", , vbYesNo + vbDefaultButton2)
    
    If response <> 6 Then
        End
    End If
    
    Sheets(DATA).Range(Cells(2, 1), Cells(getLastRowNo, DATA_COLS)).ClearContents
    
    ' �e�[�u���̃T�C�Y�ύX
    Sheets(DATA).ListObjects(DATA_TABLE_NAME).Resize Range(Cells(DATA_ROW_HEADER, DATA_COL_REGIST_KEY), Cells(DATA_ROW_HEADER + 1, DATA_COLS))
End Sub
