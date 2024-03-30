Attribute VB_Name = "MasterUtils"
Option Explicit
' �}�X�^�̐���
' - 1��ڂ�Primary Key�ł��邱��
' - PK����󕶎���ł��邱��
' - ���R�[�h�͋�s�����܂Ȃ�����
'

Public Function getRecord(tableName As String, pk As String) As Variant
' PK����擾�������R�[�h����Ԃ�
' @param tableName: �e�[�u����
' @param pk: pk
' @return ���R�[�h����1�~n�z��
'
    ' �L�[��T��
    Dim keyList As Variant: keyList = getRecordKeyList(tableName)
    Dim rowNo As Long: rowNo = 0
    Dim i As Long
    For i = 1 To getRecordNum(tableName)
        If keyList(i, 1) = pk Then
            rowNo = i + 1
            Exit For
        End If
    Next i
    
    ' �L�[��������Ȃ��ꍇ
    If rowNo = 0 Then
        getRecord = Empty
        Exit Function
    End If
    
    ' �J������
    Dim columnNum As Long: columnNum = getColumnNum(tableName)
    
    ' ���R�[�h��Ԃ�
    getRecord = Sheets(tableName).Range(Sheets(tableName).Cells(rowNo, 1), Sheets(tableName).Cells(rowNo, columnNum)).Value

End Function

Public Function getRecords(tableName As String, key As String, keyColumnName As String) As Variant
' �L�[����擾����1�ȏ�̃��R�[�h��Ԃ�
' @param tableName: �e�[�u����
' @param key: �L�[
' @param keyColumnName: �L�[�̃J������
' @return ���R�[�h����m�~n�z��
'
    Dim i As Long
    Dim j As Long
    
    ' �L�[��T������J�������擾
    Dim keyList As Variant: keyList = getColumn(tableName, keyColumnName)
    
    ' �J������������Ȃ��ꍇ
    If IsEmpty(keyList) Then
        getRecords = Empty
        Exit Function
    End If
    
    ' ���R�[�h�����擾
    Dim recordNum As Long: recordNum = getRecordNum(tableName)
    
    ' �Y�����郌�R�[�h��
    Dim targetRecordNum As Long: targetRecordNum = 0
    ' �Y�����郌�R�[�h�̃C���f�b�N�X���X�g
    Dim targetIndexList As Variant
    ReDim targetIndexList(recordNum, 1)
    For i = 1 To recordNum
        If keyList(i, 1) = key Then
            targetRecordNum = targetRecordNum + 1
            targetIndexList(targetRecordNum, 1) = i
        End If
    Next i
    
    ' ���R�[�h��������Ȃ��ꍇ
    If targetRecordNum = 0 Then
        getRecords = Empty
        Exit Function
    End If
    
    ' �i�[�pVariant�z����쐬
    Dim targetRecords As Variant
    ReDim targetRecords(1 To targetRecordNum, 1 To getColumnNum(tableName))
    
    ' �e�[�u�����擾
    Dim table As Variant: table = getTable(tableName)
    
    For i = 1 To targetRecordNum
        For j = 1 To getColumnNum(tableName)
            targetRecords(i, j) = table(targetIndexList(i, 1), j)
        Next j
    Next i
    
    getRecords = targetRecords

End Function

Public Function getColumn(tableName As String, columnName As String) As Variant
' �J����������擾�����J���������Ԃ�
' @param tableName: �e�[�u����
' @param columnName: �J������
' @return �J��������n�~1�z��
'
    ' �J�������̃��X�g���擾
    Dim columnNameLIst As Variant: columnNameLIst = getColumnList(tableName)
    ' �J��������T��
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnNameLIst(1, i) = columnName Then
            columnNo = i
            Exit For
        End If
    Next i
    
    ' ������Ȃ��ꍇ
    If columnNo = 0 Then
        getColumn = Empty
        Exit Function
    End If
    
    ' ���R�[�h�����擾
    Dim recordNum As Long: recordNum = getRecordNum(tableName)
    
    ' �J������Ԃ�
    getColumn = Sheets(tableName).Range(Sheets(tableName).Cells(2, columnNo), Sheets(tableName).Cells(recordNum + 1, columnNo)).Value

End Function

Public Function getData(tableName As String, pk As String, columnName As String) As Variant
' ���背�R�[�h�̓���̃J�������擾
' @param tableName: �e�[�u����
' @param pk: pk
' @param columnName: �J������
' @return �擾�f�[�^��1�~1�z��
'
    ' ���R�[�h���擾
    Dim record As Variant: record = getRecord(tableName, pk)
    ' ���R�[�h��������Ȃ��ꍇ
    If IsEmpty(record) Then
        getData = Empty
        Exit Function
    End If
    
    ' �J�����ꗗ���擾
    Dim columnList As Variant: columnList = getColumnList(tableName)
    ' �J��������T��
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnList(1, i) = columnName Then
            columnNo = i
            Exit For
        End If
    Next i
    ' �J������������Ȃ��ꍇ
    If columnNo = 0 Then
        getData = Empty
        Exit Function
    End If
    
    ' ����̃f�[�^��Ԃ�
    getData = record(1, i)

End Function

Public Function getDatas(tableName As String, key As String, keyColumnName As String, targetColumnName) As Variant
' �L�[����擾����1�ȏ�̃��R�[�h�̓���̃J������Ԃ�
' @param tableName: �e�[�u����
' @param key: �L�[
' @param keyColumnName: �L�[�̃J������
' @param targetColumnName: �擾����J������
' @return �擾�f�[�^��n�~1�z��
'
    ' ���R�[�h���擾
    Dim records As Variant: records = getRecords(tableName, key, keyColumnName)
    ' ���R�[�h��������Ȃ��ꍇ
    If IsEmpty(records) Then
        getDatas = Empty
        Exit Function
    End If
    
    ' �J�����ꗗ���擾
    Dim columnList As Variant: columnList = getColumnList(tableName)
    ' �J��������T��
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnList(1, i) = targetColumnName Then
            columnNo = i
            Exit For
        End If
    Next i
    ' �J������������Ȃ��ꍇ
    If columnNo = 0 Then
        getDatas = Empty
        Exit Function
    End If
    
    ' ����̃f�[�^��Ԃ�
    Dim targetDatas As Variant
    ReDim targetDatas(1 To UBound(records, 1), 1 To 1)
    For i = 1 To UBound(targetDatas, 1)
        targetDatas(i, 1) = records(i, columnNo)
    Next i
    
    getDatas = targetDatas

End Function
Public Function getColumnNum(tableName As String) As Long
' �w�肵���e�[�u���̃J���������擾
' @param tableName: �e�[�u����
'
    getColumnNum = Sheets(tableName).Range("A1").End(xlToRight).column

End Function

Public Function getRecordNum(tableName As String) As Long
' �w�肵���e�[�u���̃��R�[�h�����擾
' @param tableName: �e�[�u����
'
    getRecordNum = Sheets(tableName).Range("A1").End(xlDown).Row - 1

End Function

Public Function getColumnList(tableName As String) As Variant
' �w�肵���e�[�u���̃J�������̃��X�g���擾
' @param tableName: �e�[�u����
' @return �J���������i�[����1�~n�̓񎟌��z��
'
    getColumnList = Sheets(tableName).Range(Sheets(tableName).Range("A1"), Sheets(tableName).Range("A1").End(xlToRight)).Value

End Function

Public Function getRecordKeyList(tableName As String) As Variant
' �w�肵���e�[�u����pk�̃��X�g���擾
' @param tableName: �e�[�u����
' @return pk���i�[����n�~1�̓񎟌��z��
'
    getRecordKeyList = Sheets(tableName).Range(Sheets(tableName).Range("A2"), Sheets(tableName).Range("A1").End(xlDown)).Value

End Function

Public Function getTable(tableName As String) As Variant
' �w�肵���e�[�u�����擾
' @paran tableName: �e�[�u����
' @return �e�[�u���̑S���R�[�h
'
    getTable = Sheets(tableName).Range(Sheets(tableName).Range("A2"), Sheets(tableName).Range("A2").End(xlToRight).End(xlDown)).Value

End Function


