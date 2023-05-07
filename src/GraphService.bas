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
