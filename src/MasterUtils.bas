Attribute VB_Name = "MasterUtils"
Option Explicit

' �}�X�^�֘A�T�[�r�X�N���X
' �}�X�^����
' - 1��ڂ�PK�ł��邱��
' - PK�͈�ӂł��邱��
' - PK���󕶎��łȂ�����
' - ���R�[�h�͋�s�����܂���`����Ă��邱��
'

Public Function getMasterRecord(masterName As String, key As Variant) As Range
' �}�X�^���烌�R�[�h���擾
' @param masterName: �}�X�^�e�[�u����
' @param key: �擾���郌�R�[�h�L�[

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �L�[��T��
    Dim keysColumn As Range: Set keysColumn = ActiveSheet.Range("A2", Range("A2").End(xlDown))
    Dim findKey As Range: Set findKey = findWholeMatch(keysColumn, key)
    
    ' �L�[��������Ȃ��ꍇ
    If findKey Is Nothing Then
        Set getMasterRecord = Nothing
        Exit Function
    End If
    
    ' �擾���郌�R�[�h�̍s�ԍ�
    Dim recordRowNo As Long: recordRowNo = findKey.Row
    ' �J������
    Dim columnNum As Long: columnNum = getMasterColumnNum(masterName)
    
    ' ���R�[�h��Ԃ�
    Set getMasterRecord = ActiveSheet.Range(Cells(recordRowNo, 1), Cells(recordRowNo, columnNum))

End Function

Public Function getMasterColumn(masterName As String, column As String) As Range
' �}�X�^����J�������擾
' @param masterName: �}�X�^�e�[�u����
' @param key: �擾����J������

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �J������T��
    Dim columnList As Range: Set columnList = getMasterColumnList(masterName)
    Dim findColumn As Range: Set findColumn = findWholeMatch(columnList, column)
    
    ' �J������������Ȃ��ꍇ
    If findColumn Is Nothing Then
        Debug.Print column; masterName
        Set getMasterColumn = Nothing
    End If
        
    ' �擾����J�����̗�ԍ�
    Dim columnNo As Long: columnNo = findColumn.column
    ' ���R�[�h��
    Dim recordNum As Long: recordNum = getMasterRecordRowNo(masterName)
    
    ' �e���R�[�h�̎擾�J������Ԃ�
    Set getMasterColumn = ActiveSheet.Range(Cells(2, columnNo), Cells(recordNum, columnNo))

End Function

Function getMasterData(masterName As String, key As String, column As String) As Range
' �}�X�^���烌�R�[�h�̓���̃f�[�^���擾
' @param masterName: �}�X�^�e�[�u����
' @param key: �擾���郌�R�[�h�L�[
' @param column: �擾����J������

    ' �擾����J�����̗�ԍ�
    Dim columnNo As Long: columnNo = findWholeMatch(getMasterColumnList(masterName), column).column
    ' �擾���郌�R�[�h�̍s�ԍ�
    Dim rowNo As Long: rowNo = getMasterRecord(masterName, key).Row

    Set getMasterData = Cells(rowNo, columnNo)
    
End Function

Public Function getMasterColumnList(masterName As String) As Range
' �}�X�^�̃J���������X�g���擾
' @param masterName: �}�X�^�e�[�u����

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �J�������X�g��Ԃ�
    Set getMasterColumnList = ActiveSheet.Range("A1", Range("A1").End(xlToRight))

End Function

Public Function getMasterKeyList(masterName As String) As Range
' �}�X�^�̃L�[���X�g���擾
' @param masterName: �}�X�^�e�[�u����

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �L�[���X�g��Ԃ�
    Set getMasterKeyList = ActiveSheet.Range("A1", Range("A1").End(xlDown))

End Function

Public Function getMasterColumnNum(masterName As String) As Long
' �}�X�^�̃J���������擾
' @param masterName: �}�X�^�e�[�u����

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �J��������Ԃ�
     getMasterColumnNum = ActiveSheet.Range("A1").End(xlToRight).column

End Function

Public Function getMasterRecordRowNo(masterName As String) As Long
' �}�X�^�̍ŏI���R�[�h�̍s�ԍ���Ԃ�
' @param masterName: �}�X�^�e�[�u����

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �ŏI�L�[�̍s�ԍ���Ԃ�
    getMasterRecordRowNo = ActiveSheet.Range("A1").End(xlDown).Row

End Function

Public Function getMasterRecords(masterName As String, key As Variant, keyColumnName As String) As Range
' ��PK���畡�����R�[�h���擾����
' @param masterName: �}�X�^�e�[�u����
' @param key: �擾���郌�R�[�h�L�[
' @param keyColumnName: �L�[�̗�

    ' �V�[�g��I��
    Call selectSheet(masterName)
    
    ' �L�[�̗�ԍ�
    Dim keyColumnNo As Long: keyColumnNo = findWholeMatch(getMasterColumnList(masterName), keyColumnName).column
    
    ' �L�[��T��
    Dim keysColumn As Range: Set keysColumn = _
        ActiveSheet.Range(Cells(2, keyColumnNo), Cells(2, keyColumnNo).End(xlDown))
    Dim findKeys As Range: Set findKeys = findAllWholeMatch(keysColumn, key)
    
    ' �L�[��������Ȃ��ꍇ
    If findKeys Is Nothing Then
        Set getMasterRecords = Nothing
        Exit Function
    End If
    
    ' �擾���郌�R�[�h�̍s�ԍ�
    Dim recordRowNoList() As Long: ReDim recordRowNoList(findKeys.Count)
    Dim i As Long, c As Long: c = 0
    For i = 1 To findKeys.Count
        recordRowNoList(c) = findKeys(i).Row
        c = c + 1
    Next i
    
    ' �J������
    Dim columnNum As Long: columnNum = getMasterColumnNum(masterName)
    
    ' ���R�[�h��Ԃ�
    Set getMasterRecords = ActiveSheet.Range(Cells(recordRowNoList(0), 1), Cells(recordRowNoList(0), columnNum))
    For i = 1 To findKeys.Count - 1
        Set getMasterRecords = Union(getMasterRecords, _
            ActiveSheet.Range(Cells(recordRowNoList(i), 1), Cells(recordRowNoList(i), columnNum)))
    Next i

End Function

Public Function getMasterDatas(masterName As String, key As Variant, keyColumnName As String, column As String) As Range
' ��PK���畡���̃��R�[�h�̓���̃f�[�^���擾
' @param masterName: �}�X�^�e�[�u����
' @param key: �擾���郌�R�[�h�L�[
' @param keyColumnName: �L�[�̗�
' @param column: �擾����J������

    ' �擾����J�����̗�ԍ�
    Dim columnNo As Long: columnNo = findWholeMatch(getMasterColumnList(masterName), column).column
    ' �擾���郌�R�[�h���X�g
    Dim records As Range: Set records = getMasterRecords(masterName, key, keyColumnName)
    
    Dim i As Long
    For i = 1 To records.Count
        If records(i).column = columnNo Then
            If getMasterDatas Is Nothing Then
                Set getMasterDatas = records(i)
            Else
                Set getMasterDatas = Union(getMasterDatas, records(i))
            End If
        End If
    Next i

End Function

Sub test()
    
    Dim hoge As Range: Set hoge = getMasterDatas("KnowledgeMaster", "SSC", "trackKey", "value")
    Dim i As Long
    For i = 1 To hoge.Count
        Debug.Print hoge(i)
    Next
    
    
End Sub
