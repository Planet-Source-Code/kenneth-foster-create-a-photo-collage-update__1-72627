Attribute VB_Name = "modResizePic"
'QuickScale by Harley Neal

Option Explicit


Public Sub ReszPic(PicBox As Object, _
                   ByVal ActualPic As StdPicture, _
                   ByVal MaxHeight As Integer, _
                   ByVal MaxWidth As Integer)

   'MaxHeight is max. image height allowed
   'MaxWidth is max. picture width allowed
  Dim NewH As Integer 'New Height
  Dim NewW As Integer 'New Width

   'set starting var.
   NewH = ActualPic.Height 'actual image height
   NewW = ActualPic.Width 'actual image width
   'do logic

   If NewH > MaxHeight Or NewW > MaxWidth Then 'picture is too large

      If NewH > NewW Then 'height is greater than width
         NewW = Fix((NewW / NewH) * MaxHeight) 'rescale height
         NewH = MaxHeight 'set max height
       ElseIf NewW > NewH Then 'width is greater than height
         NewH = Fix((NewH / NewW) * MaxWidth) 'rescale width
         NewW = MaxHeight 'set max width
         Debug.Print "Width>"
       Else 'image is perfect square
         NewH = MaxHeight
         NewW = MaxWidth
      End If

   End If

   'Draw newly scaled picture

   With PicBox
      .AutoRedraw = True 'set needed properties
      .Cls 'clear picture box
      .Width = NewW
      .Height = NewH
      .PaintPicture ActualPic, 0, 0, NewW, NewH 'paint new picture size in picturebox
   End With

End Sub

