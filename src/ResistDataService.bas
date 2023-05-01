Attribute VB_Name = "ResistDataService"
' RegistData�V�[�g�Ɋւ���T�[�r�X�N���X
'

Public Sub initInputData()
' �f�[�^���͗��̏�����
'
    ' tier���̏�����
    initInputTier
    ' �`�����̏�����
    initInputFormat
    ' �J�n���ʗ��̏�����
    initInputStartingRank
    ' �R�[�X�����̏�����
    initInputTrackName
    ' ���ʗ��̏�����
    initInputRank
    ' ���l���̏�����
    initInputRemark
    
End Sub

Private Sub initInputTier()
' tier���̏�����

    ' �����l�̎擾
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_TIER_NAME).Value
    
    Sheets(REGIST_DATA).Cells(REGIST_ROW_TIER, REGIST_COL_TIER).Value = initValue
    
    Range("A1").Select
End Sub

Private Sub initInputFormat()
' �`�����̏�����

    ' �����l�̎擾
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_FORMAT_NAME).Value
    
    Sheets(REGIST_DATA).Cells(REGIST_ROW_FORMAT, REGIST_COL_FORMAT).Value = initValue
    
    Range("A1").Select
End Sub

Private Sub initInputStartingRank()
' �J�n���ʗ��̏�����
    
    Sheets(REGIST_DATA).Cells(REGIST_ROW_START_RANK, REGIST_COL_START_RANK).Value = initValue
    
    Range("A1").Select
End Sub

Private Sub initInputTrackName()
' �R�[�X�����̏�����
'
    ' �����l�̎擾
    Dim initValue As String: initValue = Sheets(STORAGE).Cells(1, STORAGE_COL_TRACK_NAME).Value
    
    Sheets(REGIST_DATA).Select
    
    ' �������̎��s
    Dim i As Integer
    For i = 1 To RACE_NUM:
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME).Value = initValue
    Next i
    
    Range("A1").Select
    
End Sub

Private Sub initInputRank()
' ���ʗ��̏�����
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK).Value = ""
    Next i
    
    Range("A1").Select
    
End Sub
    
Private Sub initInputRemark()
' ���l���̏�����
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Cells(REGIST_ROW_HEADER + i, REGIST_COL_REMARK).Value = ""
    Next i
    
    Range("A1").Select
    
End Sub

Public Sub registData()
' �f�[�^��o�^����
'
    ' ���̓f�[�^�̌`��
    Dim iData As inputData: Set iData = createInputData
    
    ' �f�[�^�̒ǉ�
    Call addNewData(iData)
    
    ' �o�^�L�[�̉��Z
    addRegistKey
    
    ' Data�V�[�g�̕\��
    Sheets(DATA).Select

End Sub


Private Function createInputData() As inputData
' ���̓f�[�^���`������
'
    ' �o�^�L�[
    Dim registKey As String: registKey = getRegistKey
    ' ���t
    Dim playDate As Date: playDate = Date
    ' tier
    Dim tier As String: tier = Sheets(REGIST_DATA).Cells(REGIST_ROW_TIER, REGIST_COL_TIER).Value
    ' �`��
    Dim format As String: format = Sheets(REGIST_DATA).Cells(REGIST_ROW_FORMAT, REGIST_COL_FORMAT).Value
    ' �R�[�X�f�[�^
    Dim track() As TrackData
    Dim arrSize As Integer: arrSize = 0
    
    Dim i As Integer
    For i = 1 To RACE_NUM
        Dim tmpTrack As TrackData: Set tmpTrack = createTrackData(i)
        If Not tmpTrack Is Nothing Then
            ReDim Preserve track(arrSize)
            Set track(arrSize) = tmpTrack
            
            arrSize = arrSize + 1
        End If
    Next i
    
    ' ���̓f�[�^
    Dim iData As New inputData
    Call iData.init(registKey, tier, format, Date, track)
    Set createInputData = iData
    
End Function

Private Function createTrackData(i As Integer) As TrackData
' �R�[�X�f�[�^���`������
'
    Dim rowNo As Long: rowNo = REGIST_ROW_HEADER + i
    
    ' ���̓`�F�b�N
    Dim isCompleted As Boolean
    isCompleted = isInputTrackName(i) And isInputRank(i)
    
    If Not isCompleted Then
        Set createTrackData = Nothing
    Else
        Dim track As New TrackData
        Call track.init(getTrackKey(getTrackName(i)), getStartingRank(i), getResultRank(i), getRemark(i))
        Set createTrackData = track
    End If
    
End Function

Private Function isInputTrackName(i As Integer)
' �R�[�X�������͂���Ă��邩
'
    ' ���I���̕���
    Dim unselectValue As String: unselectValue = getUnselectValue(getLanguage)
    ' ����ΏۃZ��
    Dim c As Range: Set c = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME)
    
    isInputTrackName = c.Value <> "" And c.Value <> unselectValue
End Function

Private Function isInputRank(i As Integer)
' ���ʂ����͂���Ă��邩

    ' ����ΏۃZ��
    Dim c As Range: Set c = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK)
    
    isInputRank = c.Value <> ""
End Function

Private Function getTrackName(i As Integer) As String
' �R�[�X�����擾
    getTrackName = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_TRACK_NAME)
End Function

Private Function getResultRank(i As Integer) As Integer
' ���ʏ��ʂ��擾
    getResultRank = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK)
End Function

Private Function getRemark(i As Integer) As String
'���l���擾
    getRemark = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_REMARK)
End Function

Private Function getStartingRank(i As Integer) As Integer
' �X�^�[�g���ʂ��擾
    Dim offsetRow As Integer
    If i <> 1 Then
        offsetRow = -1
    Else
        offsetRow = -2
    End If
    
    getStartingRank = Sheets(REGIST_DATA).Cells(REGIST_ROW_HEADER + i, REGIST_COL_RANK).Offset(offsetRow, 0)
End Function