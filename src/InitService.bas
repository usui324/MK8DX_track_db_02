Attribute VB_Name = "InitService"
' 初期化に関するサービス
'
Public Sub init()
' 初期化処理
'
    ' メッセージの表示
    Dim res As VbMsgBoxResult: res = openMsgBox("入力が初期化されます。よろしいでしょうか？", , vbOKCancel)
    If res = VbMsgBoxResult.vbCancel Then
        Exit Sub
    End If
    
    ' 言語設定
    Dim language As String: language = getLanguage

    ' Storageの初期化
    initStorage (language)
    
    ' RegistDataの初期化
    initInputData
    
End Sub

