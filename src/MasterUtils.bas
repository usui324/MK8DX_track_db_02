Attribute VB_Name = "MasterUtils"
Option Explicit
' マスタの制約
' - 1列目がPrimary Keyであること
' - PKが非空文字列であること
' - レコードは空行を挟まないこと
'

Public Function getRecord(tableName As String, pk As String) As Variant
' PKから取得したレコードを一つ返す
' @param tableName: テーブル名
' @param pk: pk
' @return レコード情報の1×n配列
'
    ' キーを探索
    Dim keyList As Variant: keyList = getRecordKeyList(tableName)
    Dim rowNo As Long: rowNo = 0
    Dim i As Long
    For i = 1 To getRecordNum(tableName)
        If keyList(i, 1) = pk Then
            rowNo = i + 1
            Exit For
        End If
    Next i
    
    ' キーが見つからない場合
    If rowNo = 0 Then
        getRecord = Empty
        Exit Function
    End If
    
    ' カラム数
    Dim columnNum As Long: columnNum = getColumnNum(tableName)
    
    ' レコードを返す
    getRecord = Sheets(tableName).Range(Sheets(tableName).Cells(rowNo, 1), Sheets(tableName).Cells(rowNo, columnNum)).Value

End Function

Public Function getRecords(tableName As String, key As String, keyColumnName As String) As Variant
' キーから取得した1つ以上のレコードを返す
' @param tableName: テーブル名
' @param key: キー
' @param keyColumnName: キーのカラム名
' @return レコード情報のm×n配列
'
    Dim i As Long
    Dim j As Long
    
    ' キーを探索するカラムを取得
    Dim keyList As Variant: keyList = getColumn(tableName, keyColumnName)
    
    ' カラムが見つからない場合
    If IsEmpty(keyList) Then
        getRecords = Empty
        Exit Function
    End If
    
    ' レコード数を取得
    Dim recordNum As Long: recordNum = getRecordNum(tableName)
    
    ' 該当するレコード数
    Dim targetRecordNum As Long: targetRecordNum = 0
    ' 該当するレコードのインデックスリスト
    Dim targetIndexList As Variant
    ReDim targetIndexList(recordNum, 1)
    For i = 1 To recordNum
        If keyList(i, 1) = key Then
            targetRecordNum = targetRecordNum + 1
            targetIndexList(targetRecordNum, 1) = i
        End If
    Next i
    
    ' レコードが見つからない場合
    If targetRecordNum = 0 Then
        getRecords = Empty
        Exit Function
    End If
    
    ' 格納用Variant配列を作成
    Dim targetRecords As Variant
    ReDim targetRecords(1 To targetRecordNum, 1 To getColumnNum(tableName))
    
    ' テーブルを取得
    Dim table As Variant: table = getTable(tableName)
    
    For i = 1 To targetRecordNum
        For j = 1 To getColumnNum(tableName)
            targetRecords(i, j) = table(targetIndexList(i, 1), j)
        Next j
    Next i
    
    getRecords = targetRecords

End Function

Public Function getColumn(tableName As String, columnName As String) As Variant
' カラム名から取得したカラムを一列返す
' @param tableName: テーブル名
' @param columnName: カラム名
' @return カラム情報のn×1配列
'
    ' カラム名のリストを取得
    Dim columnNameLIst As Variant: columnNameLIst = getColumnList(tableName)
    ' カラム名を探索
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnNameLIst(1, i) = columnName Then
            columnNo = i
            Exit For
        End If
    Next i
    
    ' 見つからない場合
    If columnNo = 0 Then
        getColumn = Empty
        Exit Function
    End If
    
    ' レコード数を取得
    Dim recordNum As Long: recordNum = getRecordNum(tableName)
    
    ' カラムを返す
    getColumn = Sheets(tableName).Range(Sheets(tableName).Cells(2, columnNo), Sheets(tableName).Cells(recordNum + 1, columnNo)).Value

End Function

Public Function getData(tableName As String, pk As String, columnName As String) As Variant
' 特定レコードの特定のカラムを取得
' @param tableName: テーブル名
' @param pk: pk
' @param columnName: カラム名
' @return 取得データの1×1配列
'
    ' レコードを取得
    Dim record As Variant: record = getRecord(tableName, pk)
    ' レコードが見つからない場合
    If IsEmpty(record) Then
        getData = Empty
        Exit Function
    End If
    
    ' カラム一覧を取得
    Dim columnList As Variant: columnList = getColumnList(tableName)
    ' カラム名を探索
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnList(1, i) = columnName Then
            columnNo = i
            Exit For
        End If
    Next i
    ' カラムが見つからない場合
    If columnNo = 0 Then
        getData = Empty
        Exit Function
    End If
    
    ' 特定のデータを返す
    getData = record(1, i)

End Function

Public Function getDatas(tableName As String, key As String, keyColumnName As String, targetColumnName) As Variant
' キーから取得した1つ以上のレコードの特定のカラムを返す
' @param tableName: テーブル名
' @param key: キー
' @param keyColumnName: キーのカラム名
' @param targetColumnName: 取得するカラム名
' @return 取得データのn×1配列
'
    ' レコードを取得
    Dim records As Variant: records = getRecords(tableName, key, keyColumnName)
    ' レコードが見つからない場合
    If IsEmpty(records) Then
        getDatas = Empty
        Exit Function
    End If
    
    ' カラム一覧を取得
    Dim columnList As Variant: columnList = getColumnList(tableName)
    ' カラム名を探索
    Dim columnNo As Long: columnNo = 0
    Dim i As Long
    For i = 1 To getColumnNum(tableName)
        If columnList(1, i) = targetColumnName Then
            columnNo = i
            Exit For
        End If
    Next i
    ' カラムが見つからない場合
    If columnNo = 0 Then
        getDatas = Empty
        Exit Function
    End If
    
    ' 特定のデータを返す
    Dim targetDatas As Variant
    ReDim targetDatas(1 To UBound(records, 1), 1 To 1)
    For i = 1 To UBound(targetDatas, 1)
        targetDatas(i, 1) = records(i, columnNo)
    Next i
    
    getDatas = targetDatas

End Function
Public Function getColumnNum(tableName As String) As Long
' 指定したテーブルのカラム数を取得
' @param tableName: テーブル名
'
    getColumnNum = Sheets(tableName).Range("A1").End(xlToRight).column

End Function

Public Function getRecordNum(tableName As String) As Long
' 指定したテーブルのレコード数を取得
' @param tableName: テーブル名
'
    getRecordNum = Sheets(tableName).Range("A1").End(xlDown).Row - 1

End Function

Public Function getColumnList(tableName As String) As Variant
' 指定したテーブルのカラム名のリストを取得
' @param tableName: テーブル名
' @return カラム名を格納した1×nの二次元配列
'
    getColumnList = Sheets(tableName).Range(Sheets(tableName).Range("A1"), Sheets(tableName).Range("A1").End(xlToRight)).Value

End Function

Public Function getRecordKeyList(tableName As String) As Variant
' 指定したテーブルのpkのリストを取得
' @param tableName: テーブル名
' @return pkを格納したn×1の二次元配列
'
    getRecordKeyList = Sheets(tableName).Range(Sheets(tableName).Range("A2"), Sheets(tableName).Range("A1").End(xlDown)).Value

End Function

Public Function getTable(tableName As String) As Variant
' 指定したテーブルを取得
' @paran tableName: テーブル名
' @return テーブルの全レコード
'
    getTable = Sheets(tableName).Range(Sheets(tableName).Range("A2"), Sheets(tableName).Range("A2").End(xlToRight).End(xlDown)).Value

End Function


