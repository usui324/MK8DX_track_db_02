Attribute VB_Name = "ImgService"
Option Explicit

Public Sub removeTrackImg()
' 呼び出し元の画像を削除する
'
    ActiveSheet.Shapes(Application.Caller).Delete

End Sub

Public Sub removeAllTrackImg()
' コース画像を全て削除する
'
    Dim img As Shape
    For Each img In Sheets(REGIST_DATA).Shapes
        If img.Type = msoPicture Then
            img.Delete
        End If
    Next

End Sub

Public Sub addTrackImg(trackKey As String)
' コース画像を追加する
'
    Dim path As String: path = ThisWorkbook.path & TRACK_IMG_DIR
    Dim fileName As String: fileName = path & getTrackImgFileName(trackKey)
    Dim img As Shape
    
    ' TODO: MasterUtilsのリファクタ、ActiveSheet,SelectedSheetの撲滅
    Sheets(REGIST_DATA).Select
    
    ' 画像を追加
    Set img = ActiveSheet.Shapes.AddPicture( _
        fileName:=fileName, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=TRACK_IMG_LEFT, _
        Top:=TRACK_IMG_TOP, _
        Width:=TRACK_IMG_WIDTH, _
        Height:=TRACK_IMG_HEIGHT _
    )
    ' 画像にマクロを登録
    img.OnAction = "removeTrackImg"
    ' 画像に効果を付与
    With img.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent1
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6
        .Radius = 5
    End With
    
End Sub
