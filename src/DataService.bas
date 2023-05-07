Attribute VB_Name = "DataService"
Public Sub addNewData(newData As inputData)
' 新規データの追加
'
    ' 行番号
    Dim rowNo As Long: rowNo = getNewRowNo
    
    If Not isCorrectArray(newData.TrackData) Then
        Call openMsgBox("有効な登録データがありません。", , vbOKOnly)
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
' データ追加用の行番号を取得
'
    getNewRowNo = getLastRowNo + 1
End Function

Public Function getLastRowNo() As Long
' 入力データの最終行の行番号を取得
'
    Dim lastRowCell As Range: Set lastRowCell = Sheets(DATA).Cells(Rows.Count, 1).End(xlUp)
    
    ' 例外処理
    If lastRowCell.Row = DATA_ROW_HEADER + 1 And lastRowCell.Value = "" Then
        getLastRowNo = DATA_ROW_HEADER
    Else
        getLastRowNo = lastRowCell.Row
    End If

End Function

Public Function getNewRegistKey() As Long
' 新規登録キーを取得
'
    getNewRegistKey = getLastRegistKey + 1
    
    If getNewRegistKey > REGIST_KEY_MAX Then
        Call openMsgBox("これ以上のデータ登録を受け付けません。")
        End
    End If
   
End Function

Public Function getLastRegistKey() As Long
' 入力データの最新の登録キーを取得
'
    Dim rowNo As Long: rowNo = getLastRowNo
    If rowNo = 1 Then
        getLastRegistKey = 0
    Else
        getLastRegistKey = Sheets(DATA).Cells(rowNo, DATA_COL_REGIST_KEY).Value
    End If
End Function

Public Sub exportData()
' データをエクスポート
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

Private Function printLine(ws As Worksheet, rowNo As Long) As String
' ファイル出力する一行の文字列を返す
'
    Dim i As Integer
    printLine = ws.Cells(rowNo, 1).Value
    For i = 2 To DATA_COLS
        printLine = printLine & "," & ws.Cells(rowNo, i).Value
    Next i
    
End Function

Public Sub importData()
' データをインポート
'
    Dim openFileName As String
    Dim ws As Worksheet
    Dim line As String
    Dim arr As Variant
    Dim rowNo As Integer: rowNo = 2
    Dim columnNo As Integer

    ' インポートファイルを指定
    ChDir ThisWorkbook.Path
    openFileName = Application.GetOpenFilename("模擬データ,*.txt", , "インポートするデータファイルを指定")
    
    ' キャンセル処理
    If openFileName = "False" Then
        Exit Sub
    End If
    
    ' 入力対象シート
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
    
    Call openMsgBox("データをインポートしました", , vbInformation)
    
End Sub

Public Sub deleteData()
' データを削除する
'
    ' メッセージ表示
    Dim response As Integer
    response = openMsgBox("データを完全削除しますか？", , vbYesNo + vbDefaultButton2)
    
    If response <> 6 Then
        End
    End If
    
    response = openMsgBox("本当によろしいですか？", , vbYesNo + vbDefaultButton2)
    
    If response <> 6 Then
        End
    End If
    
    Sheets(DATA).Range(Cells(2, 1), Cells(getLastRowNo, DATA_COLS)).ClearContents
    
    ' テーブルのサイズ変更
    Sheets(DATA).ListObjects(DATA_TABLE_NAME).Resize Range(Cells(DATA_ROW_HEADER, DATA_COL_REGIST_KEY), Cells(DATA_ROW_HEADER + 1, DATA_COLS))
End Sub
