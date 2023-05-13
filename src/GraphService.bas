Attribute VB_Name = "GraphService"
Option Explicit

Public Sub updateGraphs()
' �O���t���X�V����
'
    ActiveWorkbook.RefreshAll
End Sub

Public Sub resetGraphFilter()
' �O���t�̃t�B���^�[�����Z�b�g����
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' �t�B���^�[�̃��Z�b�g
    pTable.PivotFields(PIVOT_FILTER_NAME_1).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_2).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_3).CurrentPage = "(ALL)"
    pTable.PivotFields(PIVOT_FILTER_NAME_4).CurrentPage = "(ALL)"
    
End Sub

Public Sub setGraphMinNumOfRace()
' �K�背�[�X���̐ݒ��������
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' �K�背�[�X���̎擾
    Dim reguRaceNum As Long: reguRaceNum = Sheets(SETTINGS).Cells(SETTINGS_ROW_RACE_NUM, SETTINGS_COL_VALUE).Value
    
    ' �ݒ�������� ' TODO: �s�\�[�X���ς�����Ƃ��̑Ή�
    pTable.PivotFields(PIVOT_ROW_NAME).ClearAllFilters
    pTable.PivotFields(PIVOT_ROW_NAME).PivotFilters. _
        Add2 Type:=xlValueIsGreaterThanOrEqualTo, _
        DataField:=pTable.PivotFields(PIVOT_COL_NAME_3), Value1:=reguRaceNum
End Sub

Public Sub sortGraphByPoint()
' ���ϓ��_�Ń\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    ' �t���O�̒l
    Dim flg As Integer: flg = getPointFlg
    
    ' �K�背�[�X�t�B���^�[
    setGraphMinNumOfRace
    
    ' �t���O�̒l��0�Ȃ珸���\�[�g / 1�Ȃ�~���\�[�g
    If flg = 0 Then
        sortGraphByPointAscending
    Else
        sortGraphByPointDescending
    End If
    
    ' �t���O�����Z
    incrementPointFlg
    
End Sub

Public Sub sortGraphByRank()
' ���Ϗ��ʂŃ\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    ' �K�背�[�X�t�B���^�[
    setGraphMinNumOfRace
    
    ' �t���O�̒l
    Dim flg As Integer: flg = getRankFlg
    
    ' �t���O�̒l��0�Ȃ珸���\�[�g / 1�Ȃ�~���\�[�g
    If flg = 0 Then
        sortGraphByRankAscending
    Else
        sortGraphByRankDescending
    End If
    
    ' �t���O�����Z
    incrementRankFlg
    
End Sub

Public Sub sortGraphByTimes()
' �񐔂Ń\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    ' �t�B���^�[�����Z�b�g
    pTable.PivotFields(PIVOT_ROW_NAME).ClearAllFilters
    
    ' �t���O�̒l
    Dim flg As Integer: flg = getTimesFlg
    
    ' �t���O�̒l��0�Ȃ珸���\�[�g / 1�Ȃ�~���\�[�g
    If flg = 0 Then
        sortGraphByTimesAscending
    Else
        sortGraphByTimesDescending
    End If
    
    ' �t���O�����Z
    incrementTimesFlg
    
End Sub

Private Sub sortGraphByPointAscending()
' ���ϓ��_�̏����\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_1
    
End Sub

Private Sub sortGraphByPointDescending()
' ���ϓ��_�̍~���\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_1
    
End Sub

Private Sub sortGraphByRankAscending()
' ���Ϗ��ʂ̏����\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_2
    
End Sub

Private Sub sortGraphByRankDescending()
' ���Ϗ��ʂ̍~���\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_2
    
End Sub

Private Sub sortGraphByTimesAscending()
' �񐔂̏����\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlAscending, PIVOT_COL_NAME_3
    
End Sub

Private Sub sortGraphByTimesDescending()
' �񐔂̍~���\�[�g
'
    ' �s�{�b�g�e�[�u��
    Dim pTable As PivotTable: Set pTable = Sheets(GRAPH).PivotTables(GRAPH_PIVOT_TABLE_NAME)
    
    pTable.PivotFields(PIVOT_ROW_NAME).AutoSort xlDescending, PIVOT_COL_NAME_3
    
End Sub
