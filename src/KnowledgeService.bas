Attribute VB_Name = "KnowledgeService"
Option Explicit

Public Sub initKnowledge()
' �m�����̏�����
'
    Sheets(REGIST_DATA).Select
    
    Dim i As Integer
    For i = 0 To MAX_KNOWLEDGE - 1
        Cells(REGIST_ROW_KNOWLEDGE + i, REGIST_COL_KNOWLEDGE).Value = ""
    Next i
    
    Cells(REGIST_ROW_KNOWLEDGE, REGIST_COL_TRACK_NAME).Value = ""
    
    Range(INIT_SELECT_REGIST_DATA).Select
    
End Sub

Public Sub setKnowledge(trackKey As String)
' �m�����̃Z�b�g
'
    ' �m���̃Z�b�g
    Dim knowledges As Variant: knowledges = getKnowledgeList(trackKey)
    If Not IsEmpty(knowledges) Then
        
        Sheets(REGIST_DATA).Cells(REGIST_ROW_KNOWLEDGE, REGIST_COL_TRACK_NAME) = getTrackNameJp(trackKey)
        
        Dim i As Long
        For i = 1 To UBound(knowledges)
            Sheets(REGIST_DATA).Cells(REGIST_ROW_KNOWLEDGE + i - 1, REGIST_COL_KNOWLEDGE) = knowledges(i, 1)
        Next i
        
    End If
        
    ' �V�[�g�I��
    Sheets(REGIST_DATA).Select
    ' �Z���I��
    Range(INIT_SELECT_REGIST_DATA).Select
    
End Sub
