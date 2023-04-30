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

Public Function findWholeMatch(r As Range, target As Variant) As Range
' Rangeオブジェクトから完全一致するオブジェクトを探索する
' @param r 探索元Rangeオブジェクト
' @param target 探索対象文字列

    Set findWholeMatch = r.Find(target, LookAt:=xlWhole, MatchCase:=True)
    
End Function

Public Sub goToRegistDataSheet()
' データ登録シートへ移動
    Application.ScreenUpdating = False
    
    ' シート選択
    selectSheet (REGIST_DATA)
    ' ウィンドウサイズの調整
    Application.WindowState = xlNormal
    ActiveWindow.Width = 430
    ActiveWindow.Height = 720

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

    Application.ScreenUpdating = True
End Sub
