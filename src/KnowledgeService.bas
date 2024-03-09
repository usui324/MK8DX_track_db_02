Attribute VB_Name = "KnowledgeService"
Option Explicit

Public Sub initKnowledge()
' 知識欄の初期化
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
' 知識欄のセット
'
    ' 知識のセット
    Dim knowledges As Range: Set knowledges = getKnowledgeList(trackKey)
    If Not knowledges Is Nothing Then
        
        Sheets(REGIST_DATA).Cells(REGIST_ROW_KNOWLEDGE, REGIST_COL_TRACK_NAME) = getTrackNameJp(trackKey)
      
        knowledges.Copy
        Sheets(REGIST_DATA).Cells(REGIST_ROW_KNOWLEDGE, REGIST_COL_KNOWLEDGE).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
        
    ' シート選択
    Sheets(REGIST_DATA).Select
    ' セル選択
    Range(INIT_SELECT_REGIST_DATA).Select
    
End Sub
