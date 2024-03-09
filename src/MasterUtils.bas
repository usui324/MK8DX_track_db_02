Attribute VB_Name = "MasterUtils"
Option Explicit

' マスタ関連サービスクラス
' マスタ制約
' - 1列目がPKであること
' - PKは一意であること
' - PKが空文字でないこと
' - レコードは空行を挟まず定義されていること
'

Public Function getMasterRecord(masterName As String, key As Variant) As Range
' マスタからレコードを取得
' @param masterName: マスタテーブル名
' @param key: 取得するレコードキー

    ' シートを選択
    Call selectSheet(masterName)
    
    ' キーを探索
    Dim keysColumn As Range: Set keysColumn = ActiveSheet.Range("A2", Range("A2").End(xlDown))
    Dim findKey As Range: Set findKey = findWholeMatch(keysColumn, key)
    
    ' キーが見つからない場合
    If findKey Is Nothing Then
        Set getMasterRecord = Nothing
        Exit Function
    End If
    
    ' 取得するレコードの行番号
    Dim recordRowNo As Long: recordRowNo = findKey.Row
    ' カラム数
    Dim columnNum As Long: columnNum = getMasterColumnNum(masterName)
    
    ' レコードを返す
    Set getMasterRecord = ActiveSheet.Range(Cells(recordRowNo, 1), Cells(recordRowNo, columnNum))

End Function

Public Function getMasterColumn(masterName As String, column As String) As Range
' マスタからカラムを取得
' @param masterName: マスタテーブル名
' @param key: 取得するカラム名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' カラムを探索
    Dim columnList As Range: Set columnList = getMasterColumnList(masterName)
    Dim findColumn As Range: Set findColumn = findWholeMatch(columnList, column)
    
    ' カラムが見つからない場合
    If findColumn Is Nothing Then
        Debug.Print column; masterName
        Set getMasterColumn = Nothing
    End If
        
    ' 取得するカラムの列番号
    Dim columnNo As Long: columnNo = findColumn.column
    ' レコード数
    Dim recordNum As Long: recordNum = getMasterRecordRowNo(masterName)
    
    ' 各レコードの取得カラムを返す
    Set getMasterColumn = ActiveSheet.Range(Cells(2, columnNo), Cells(recordNum, columnNo))

End Function

Function getMasterData(masterName As String, key As String, column As String) As Range
' マスタからレコードの特定のデータを取得
' @param masterName: マスタテーブル名
' @param key: 取得するレコードキー
' @param column: 取得するカラム名

    ' 取得するカラムの列番号
    Dim columnNo As Long: columnNo = findWholeMatch(getMasterColumnList(masterName), column).column
    ' 取得するレコードの行番号
    Dim rowNo As Long: rowNo = getMasterRecord(masterName, key).Row

    Set getMasterData = Cells(rowNo, columnNo)
    
End Function

Public Function getMasterColumnList(masterName As String) As Range
' マスタのカラム名リストを取得
' @param masterName: マスタテーブル名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' カラムリストを返す
    Set getMasterColumnList = ActiveSheet.Range("A1", Range("A1").End(xlToRight))

End Function

Public Function getMasterKeyList(masterName As String) As Range
' マスタのキーリストを取得
' @param masterName: マスタテーブル名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' キーリストを返す
    Set getMasterKeyList = ActiveSheet.Range("A1", Range("A1").End(xlDown))

End Function

Public Function getMasterColumnNum(masterName As String) As Long
' マスタのカラム数を取得
' @param masterName: マスタテーブル名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' カラム数を返す
     getMasterColumnNum = ActiveSheet.Range("A1").End(xlToRight).column

End Function

Public Function getMasterRecordRowNo(masterName As String) As Long
' マスタの最終レコードの行番号を返す
' @param masterName: マスタテーブル名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' 最終キーの行番号を返す
    getMasterRecordRowNo = ActiveSheet.Range("A1").End(xlDown).Row

End Function

Public Function getMasterRecords(masterName As String, key As Variant, keyColumnName As String) As Range
' 非PKから複数レコードを取得する
' @param masterName: マスタテーブル名
' @param key: 取得するレコードキー
' @param keyColumnName: キーの列名

    ' シートを選択
    Call selectSheet(masterName)
    
    ' キーの列番号
    Dim keyColumnNo As Long: keyColumnNo = findWholeMatch(getMasterColumnList(masterName), keyColumnName).column
    
    ' キーを探索
    Dim keysColumn As Range: Set keysColumn = _
        ActiveSheet.Range(Cells(2, keyColumnNo), Cells(2, keyColumnNo).End(xlDown))
    Dim findKeys As Range: Set findKeys = findAllWholeMatch(keysColumn, key)
    
    ' キーが見つからない場合
    If findKeys Is Nothing Then
        Set getMasterRecords = Nothing
        Exit Function
    End If
    
    ' 取得するレコードの行番号
    Dim recordRowNoList() As Long: ReDim recordRowNoList(findKeys.Count)
    Dim i As Long, c As Long: c = 0
    For i = 1 To findKeys.Count
        recordRowNoList(c) = findKeys(i).Row
        c = c + 1
    Next i
    
    ' カラム数
    Dim columnNum As Long: columnNum = getMasterColumnNum(masterName)
    
    ' レコードを返す
    Set getMasterRecords = ActiveSheet.Range(Cells(recordRowNoList(0), 1), Cells(recordRowNoList(0), columnNum))
    For i = 1 To findKeys.Count - 1
        Set getMasterRecords = Union(getMasterRecords, _
            ActiveSheet.Range(Cells(recordRowNoList(i), 1), Cells(recordRowNoList(i), columnNum)))
    Next i

End Function

Public Function getMasterDatas(masterName As String, key As Variant, keyColumnName As String, column As String) As Range
' 非PKから複数のレコードの特定のデータを取得
' @param masterName: マスタテーブル名
' @param key: 取得するレコードキー
' @param keyColumnName: キーの列名
' @param column: 取得するカラム名

    ' 取得するカラムの列番号
    Dim columnNo As Long: columnNo = findWholeMatch(getMasterColumnList(masterName), column).column
    ' 取得するレコードリスト
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
