Attribute VB_Name = "DataService"
Public Sub addNewData(newData As inputData)
' 新規データの追加
'
    ' 行番号
    Dim rowNo As Long: rowNo = getNewRowNo
    
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
' データ追加用の行番号を取得
'
    getNewRowNo = getLastRowNo + 1
End Function

Public Function getLastRowNo() As Long
' 入力データの最終行の行番号を取得
'
    Dim lastRowCell As Range: Set lastRowCell = Sheets(DATA).Cells(Rows.Count, 1).End(xlUp)
    getLastRowNo = lastRowCell.Row
End Function

Public Sub exportData()
' データをエクスポートする
'
     ' エクスポートファイルを指定
    ChDir ThisWorkbook.Path
    Dim saveFileName As String
    saveFileName = Application.GetSaveAsFilename(InitialFileName:="mogiData.txt", filefilter:="模擬データ,*.txt")

    ' キャンセル処理
    If saveFileName = "False" Then
        Exit Sub
    End If
    
    ' 出力する対象シート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DATA)

    ' ファイルに書き込み
    Open saveFileName For Output As #1
    
    Dim i As Long
    For i = 2 To getLastRowNo
        Print #1, printLine(ws, i)
    Next i
    
    Close #1
    
    Call openMsgBox(saveFileName & "にデータを出力しました", , vbInformation)

End Sub

Function printLine(ws As Worksheet, rowNo As Long) As String
' ファイル出力する一行の文字列を返す
'
    Dim i As Integer
    printLine = ws.Cells(rowNo, 1).Value
    For i = 2 To DATA_COLS
        printLine = printLine & "," & ws.Cells(rowNo, i).Value
    Next i
    
End Function



















