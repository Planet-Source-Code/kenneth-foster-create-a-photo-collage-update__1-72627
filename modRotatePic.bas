Attribute VB_Name = "modRotatePic"
'this code is from http://www.vbforums.com/showthread.php?t=587739

Option Explicit

Public Declare Function PlgBlt Lib "GDI32.dll" ( _
      ByVal hDCDest As Long, _
      ByRef lpPoint As PointAPI, _
      ByVal hdcSrc As Long, _
      ByVal nXSrc As Long, _
      ByVal nYSrc As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hbmMask As Long, _
      ByVal xMask As Long, _
      ByVal yMask As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32.dll" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "GDI32.dll" ( _
      ByVal hDC As Long, _
      ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "GDI32.dll" (ByVal hDC As Long) As Long

Public Type PointAPI
   X As Long
   Y As Long
End Type


Public Function DrawStdPictureRot(frm As Object, _
                                  ByVal inDC As Long, _
                                  ByVal inX As Long, _
                                  ByVal inY As Long, _
                                  ByVal inAngle As Single, _
                                  ByRef inPicture As StdPicture) As Long

  Dim hDC As Long
  Dim hOldBMP As Long
  Dim PlgPts(0 To 4) As PointAPI
  Dim PicWidth As Long
  Dim  PicHeight As Long
  Dim HalfWidth As Single
  Dim  HalfHeight As Single
  Dim AngleRad As Single

  Const Pi As Single = 3.14159
  Const HalfPi As Single = Pi * 0.5

   ' Validate input picture
   If (inPicture Is Nothing) Then Exit Function
   If (inPicture.Type <> vbPicTypeBitmap) Then Exit Function

   ' Get picture size
   PicWidth = frm.ScaleX(inPicture.Width, vbHimetric, vbPixels)
   PicHeight = frm.ScaleY(inPicture.Height, vbHimetric, vbPixels)

   ' Get half picture size and angle in radians
   HalfWidth = PicWidth / 2
   HalfHeight = PicHeight / 2
   AngleRad = (inAngle / 180) * Pi

   ' Create temporary DC and select input picture into it
   hDC = CreateCompatibleDC(0&)
   hOldBMP = SelectObject(hDC, inPicture.Handle)

   If (hOldBMP) Then ' Get angle vectors for width and height
      PlgPts(0).X = Cos(AngleRad) * HalfWidth
      PlgPts(0).Y = Sin(AngleRad) * HalfWidth
      PlgPts(1).X = Cos(AngleRad + HalfPi) * HalfHeight
      PlgPts(1).Y = Sin(AngleRad + HalfPi) * HalfHeight

      ' Project parallelogram points for rotated area
      PlgPts(2).X = HalfWidth + inX - PlgPts(0).X - PlgPts(1).X
      PlgPts(2).Y = HalfHeight + inY - PlgPts(0).Y - PlgPts(1).Y
      PlgPts(3).X = HalfWidth + inX - PlgPts(1).X + PlgPts(0).X
      PlgPts(3).Y = HalfHeight + inY - PlgPts(1).Y + PlgPts(0).Y
      PlgPts(4).X = HalfWidth + inX - PlgPts(0).X + PlgPts(1).X
      PlgPts(4).Y = HalfHeight + inY - PlgPts(0).Y + PlgPts(1).Y

      ' Draw rotated image
      DrawStdPictureRot = PlgBlt(inDC, PlgPts(2), hDC, 0, 0, PicWidth, PicHeight, 0&, 0, 0)

      ' De-select Bitmap from DC
      Call SelectObject(hDC, hOldBMP)
   End If

   ' Destroy temporary DC
   Call DeleteDC(hDC)

End Function

