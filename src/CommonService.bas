Attribute VB_Name = "CommonService"
' 汎用サービスクラス
'

Public Sub selectSheet(sheetName As String)
' シートを選択する
' @param sheetName: シート名
    
On Error GoTo Exception
    
    Range("A1").Select
    
    Sheets(sheetName).Select
    
    Range("A1").Select
    
    Exit Sub
Exception:

    Call openErrorMsgBox("invalid sheetName: " & sheetName)

End Sub

Public Sub openErrorMsgBox(message As String)
' エラーメッセージを表示する
' @param message: メッセージ
    
   Call openMsgBox(message, "Error")
   End

End Sub

Public Function openMsgBox(message As String, Optional title As String = "MK8DX Track DB", Optional style As VbMsgBoxStyle = VbMsgBoxStyle.vbOKOnly) As VbMsgBoxResult

' メッセージボックスを表示する
' @param message: メッセージ内容
' @param title: タイトル
' @param style: メッセージボックスのスタイル

    openMsgBox = MsgBox(message, style, title)

End Function

Public Function findWholeMatch(r As Range, Target As Variant) As Range
' Rangeオブジェクトから完全一致するオブジェクトを探索する
' @param r 探索元Rangeオブジェクト
' @param target 探索対象文字列

    Set findWholeMatch = r.Find(Target, LookAt:=xlWhole, MatchCase:=True)
    
End Function

Public Sub goToRegistDataSheet()
' データ登録シートへ移動
    Application.ScreenUpdating = False
    
    ' シート選択
    selectSheet (REGIST_DATA)
    ' ウィンドウサイズの調整
    Application.WindowState = xlNormal
    ActiveWindow.Width = 470
    ActiveWindow.Height = 720
    ' セルの選択
    Range(INIT_SELECT_REGIST_DATA).Select
    

    Application.ScreenUpdating = True
End Sub

Public Sub goToDataSheet()
' データシートへ移動
'
    Application.ScreenUpdating = False
    
    ' シート選択
    selectSheet (DATA)
    ' ウィンドウサイズの調整
    Application.WindowState = xlMaximized
    ' セルの選択
    Range(INIT_SELECT_DATA).Select
    
    Application.ScreenUpdating = True
End Sub

Public Sub goToGraphSheet()
' グラフシートへ移動
'
    Application.ScreenUpdating = False
    
    ' シート選択
    selectSheet (GRAPH)
    ' ウィンドウサイズの調整
    Application.WindowState = xlMaximized
    ' セルの選択
    Range(INIT_SELECT_GRAPH).Select

    Application.ScreenUpdating = True
End Sub

Public Sub goToSettingsSheet()
' 設定シートへ移動
'
    Application.ScreenUpdating = False
    
    ' シート選択
    selectSheet (SETTINGS)
    ' ウィンドウサイズの調整
    Application.WindowState = xlMaximized
    ' セルの選択
    Range(INIT_SELECT_SETTINGS).Select

    Application.ScreenUpdating = True
End Sub

Public Function isCorrectArray(ByVal arrs As Variant) As Boolean
' 配列が正常か判定する
'
    isCorrectArray = True
    
    ' 最大インデックスを取得
    Dim a As Long
    On Error GoTo err
    a = UBound(arrs)
    
    ' インデックスが負数ならFalse
    If a < 0 Then
        isCorrectArray = False
    End If
    
err:
    'エラーが生じたときエラー番号で9か13の場合はFalse
    If err.Number = 9 Or err.Number = 13 Then
        isCorrectArray = False
    End If
    
End Function

Public Function convertLongToStr(longNum As Long, strSize As Integer) As String
' 数値を文字列に変換する
'
    Dim l As Integer: l = Len(CStr(longNum))
    
    If l >= strSize Then
        convertLongToStr = CStr(longNum)
        Exit Function
    End If
    
    Dim i As Integer
    For i = l + 1 To strSize
        convertLongToStr = convertLongToStr + "0"
    Next i
    
    convertLongToStr = convertLongToStr + CStr(longNum)
End Function

Public Sub saveBook()
' 保存処理
'
    ThisWorkbook.Save
End Sub
