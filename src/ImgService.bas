Attribute VB_Name = "ImgService"
Option Explicit

Public Sub removeTrackImg()
' �Ăяo�����̉摜���폜����
'
    ActiveSheet.Shapes(Application.Caller).Delete

End Sub

Public Sub removeAllTrackImg()
' �R�[�X�摜��S�č폜����
'
    Dim img As Shape
    For Each img In Sheets(REGIST_DATA).Shapes
        If img.Type = msoPicture Then
            img.Delete
        End If
    Next

End Sub

Public Sub addTrackImg(trackKey As String)
' �R�[�X�摜��ǉ�����
'
    Dim path As String: path = ThisWorkbook.path & TRACK_IMG_DIR
    Dim fileName As String: fileName = path & getTrackImgFileName(trackKey)
    Dim img As Shape
    
    ' TODO: MasterUtils�̃��t�@�N�^�AActiveSheet,SelectedSheet�̖o��
    Sheets(REGIST_DATA).Select
    
    ' �摜��ǉ�
    Set img = ActiveSheet.Shapes.AddPicture( _
        fileName:=fileName, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=TRACK_IMG_LEFT, _
        Top:=TRACK_IMG_TOP, _
        Width:=TRACK_IMG_WIDTH, _
        Height:=TRACK_IMG_HEIGHT _
    )
    ' �摜�Ƀ}�N����o�^
    img.OnAction = "removeTrackImg"
    ' �摜�Ɍ��ʂ�t�^
    With img.Glow
        .Color.ObjectThemeColor = msoThemeColorAccent1
        .Color.TintAndShade = 0
        .Color.Brightness = 0
        .Transparency = 0.6
        .Radius = 5
    End With
    
End Sub
