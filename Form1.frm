VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Create A Photo Collage"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10395
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Text to picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   60
      TabIndex        =   22
      Top             =   2880
      Width           =   3975
      Begin VB.TextBox txtAddText 
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   3735
      End
      Begin VB.CommandButton cmdFont 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Font"
         Height          =   690
         Left            =   1410
         Picture         =   "Form1.frx":74F2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton cmdTextColor 
         Caption         =   "Text Color"
         Height          =   690
         Left            =   90
         Picture         =   "Form1.frx":7C34
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Do same to add text to main page (right), but use Undo for Corrections."
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   120
         TabIndex        =   31
         Top             =   1710
         Width           =   3780
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Left click picture (below) to place text. Move mouse and repeat if not in correct place."
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   3765
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MS Sans Serif"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2610
         TabIndex        =   28
         Top             =   1095
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regular"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2610
         TabIndex        =   27
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FontSize  8"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2610
         TabIndex        =   26
         Top             =   570
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Basic Info"
      Height          =   1275
      Left            =   -30
      TabIndex        =   13
      Top             =   8340
      Width           =   4065
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "6. Press ""New"" to clear and        start over."
         Height          =   390
         Left            =   1950
         TabIndex        =   19
         Top             =   795
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "5. When Finished ""Save"""
         Height          =   255
         Left            =   1950
         TabIndex        =   18
         Top             =   540
         Width           =   2010
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "4. Right Click to Set. Repeat         steps."
         Height          =   390
         Left            =   1965
         TabIndex        =   17
         Top             =   165
         Width           =   2220
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "3. Left click on picture and drag into positon."
         Height          =   465
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   1680
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Make Adjustments in       Control Panel."
         Height          =   390
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   1740
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Load a Picture"
         Height          =   255
         Left            =   135
         TabIndex        =   14
         Top             =   210
         Width           =   1320
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9555
      Left            =   4065
      ScaleHeight     =   635
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   8
      Top             =   60
      Width           =   6285
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2100
         Index           =   0
         Left            =   1800
         ScaleHeight     =   140
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   9
         Top             =   1245
         Width           =   2100
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   3975
      Begin VB.CommandButton cmdShape 
         Caption         =   "Load Shape"
         Height          =   540
         Left            =   75
         TabIndex        =   33
         Top             =   1935
         Width           =   840
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Exit"
         Height          =   540
         Left            =   3090
         Picture         =   "Form1.frx":7F3E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1935
         Width           =   795
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Undo"
         Height          =   525
         Left            =   3090
         Picture         =   "Form1.frx":84C8
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1350
         Width           =   795
      End
      Begin VB.CommandButton cmdSaveJpeg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   525
         Left            =   2070
         Picture         =   "Form1.frx":8852
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1350
         Width           =   840
      End
      Begin VB.CommandButton cmdTexture 
         Caption         =   "Load Texture"
         Height          =   540
         Left            =   1080
         TabIndex        =   20
         Top             =   1935
         Width           =   810
      End
      Begin VB.CommandButton cmdReLoad 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ReLoad Last Pix"
         Height          =   540
         Left            =   1095
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1335
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Save"
         Height          =   540
         Left            =   2070
         Picture         =   "Form1.frx":8DDC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1935
         Width           =   840
      End
      Begin VB.CommandButton cmdLoad 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Load Pix"
         Height          =   540
         Left            =   75
         Picture         =   "Form1.frx":9366
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1335
         Width           =   840
      End
      Begin VB.HScrollBar HS3 
         Height          =   345
         Left            =   45
         Max             =   270
         Min             =   50
         TabIndex        =   5
         Top             =   915
         Value           =   100
         Width           =   3090
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   360
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   3090
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Busy Please Wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2310
         TabIndex        =   37
         Top             =   60
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Use arrow keys for fine adj."
         Height          =   225
         Left            =   1920
         TabIndex        =   36
         Top             =   2520
         Width           =   1965
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y: 39"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1065
         TabIndex        =   35
         Top             =   2505
         Width           =   810
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X: 21"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   75
         TabIndex        =   34
         Top             =   2505
         Width           =   825
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3165
         TabIndex        =   7
         Top             =   975
         Width           =   555
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Resize Picture"
         Height          =   285
         Left            =   1185
         TabIndex        =   6
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3165
         TabIndex        =   4
         Top             =   435
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rotate Picture"
         Height          =   210
         Left            =   1185
         TabIndex        =   3
         Top             =   165
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2685
      Left            =   45
      ScaleHeight     =   2685
      ScaleWidth      =   3150
      TabIndex        =   0
      Top             =   5085
      Width           =   3150
   End
   Begin VB.Image imgUndo 
      Height          =   510
      Index           =   0
      Left            =   135
      Stretch         =   -1  'True
      Top             =   7590
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   1110
      Left            =   180
      Picture         =   "Form1.frx":98F0
      Stretch         =   -1  'True
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   240
      Stretch         =   -1  'True
      Top             =   165
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Menu title 
      Caption         =   "Options"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPix 
         Caption         =   "Load Picture"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Last Picture"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadShape 
         Caption         =   "Load Shape"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBkgdCol 
         Caption         =   "Background Color"
      End
      Begin VB.Menu dash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveBitmap 
         Caption         =   "Save Bitmap"
      End
      Begin VB.Menu mnuSaveJpeg 
         Caption         =   "Save Jpeg"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Ken Foster except for some module codes

'To use T.O.P for a search , highlight just the word or words you want to search for.
'Press Ctrl + F3
' To continue search , just use F3 until all uses are found.

'***************** Table of Procedures *************
'   Private Sub Form_Load
'   Private Sub Form_Unload
'   Private Sub ReDraw
'   Private Sub cmdColor_Click
'   Private Sub cmdExit_Click
'   Private Sub cmdNew_Click
'   Private Sub cmdFont_Click
'   Private Sub cmdLoad_Click
'   Private Sub cmdSave_Click
'   Private Sub cmdSaveJpeg_Click
'   Private Sub cmdShape_Click
'   Private Sub cmdTextColor_Click
'   Private Sub cmdTexture_Click
'   Private Sub cmdUndo_Click
'   Private Sub mnuBkgdCol_Click
'   Private Sub mnuNew_Click
'   Private Sub mnuLoadPix_Click
'   Private Sub mnuSaveBitmap_Click
'   Private Sub mnuSaveJpeg_Click
'   Private Sub mnuLoadShape_Click
'   Private Sub mnuExit_Click
'   Private Sub picItem_KeyDown
'   Private Sub picItem_MouseDown
'   Private Sub picItem_MouseMove
'   Private Sub picMain_MouseDown
'   Private Sub Picture1_MouseDown
'   Private Sub HS3_Change
'   Private Sub HS3_Scroll
'   Private Sub HScroll1_Change
'   Private Sub HScroll1_Scroll
'   Private Sub JpegSave
'   Public Function SaveJPEG
'***************** End of Table ********************

Option Explicit
'used to make picture background transparent
Const RGN_AND = 1
Const RGN_COPY = 5
Const RGN_DIFF = 4
Const RGN_OR = 2
Const RGN_XOR = 3
Private Declare Sub CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long)
Private Declare Sub SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean)
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'----------------

Private Declare Function TransparentBlt Lib "msimg32.dll" ( _
      ByVal hdc As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal nSrcWidth As Long, _
      ByVal nSrcHeight As Long, _
      ByVal crTransparent As Long) As Boolean

Private CurFont    As New StdFont
Private CF         As CFDialog
Private cc         As CFDialog
Private JPEGclass  As cJpeg
Private OldX       As Integer   'used to move picbox
Private OldY       As Integer   'used to move picbox
Private tPosX      As Single    'X position to paint picture onto main pixbox
Private tPosY      As Single    'Y position to paint picture onto main pixbox
Private sText      As Boolean   'add text to picture
Private sText1     As Boolean   'add text to main picture box
Private nxt        As Integer   'undo counter
Private SetShape   As Boolean   'adding a shape and not a picture
Private shptranscol As Long     'shape transparency color
Private reload         As String

Private Sub Form_Load()

   Image1.Picture = Image2.Picture           'load default picture "NoPicture"
   'set up some property values
   Me.AutoRedraw = True
   Me.ScaleMode = vbPixels
   picItem(0).AutoRedraw = True
   picItem(0).ScaleMode = vbPixels
   Picture1.AutoRedraw = True
   Picture1.AutoSize = True
   Picture1.ScaleMode = vbPixels
   HScroll1.Max = 179
   HScroll1.Min = -179
   nxt = 0
   shptranscol = vbWhite  'default tranparency color
   ReDraw
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set JPEGclass = Nothing
   Set CurFont = Nothing
   Set CF = Nothing
   Set cc = Nothing
   Unload Me
End Sub

Private Sub cmdColor_Click()

  Dim TheColor As Long

   Set CF = New CFDialog
   If CF.VBChooseColor(TheColor, , , , Me.hWnd) Then
      picMain.BackColor = TheColor
   End If

   nxt = nxt + 1
   Load imgUndo(nxt)
   imgUndo(nxt).Picture = picMain.Image
   Set CF = Nothing
End Sub

Private Sub cmdExit_Click()

   Set JPEGclass = Nothing
   Set CurFont = Nothing
   Set CF = Nothing
   Set cc = Nothing
   Unload Me
End Sub

Private Sub cmdFont_Click()

   Set cc = New CFDialog

   If cc.VBChooseFont(CurFont, , Me.hWnd) Then

      With txtAddText
         .FontName = CurFont.Name
         '.FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .FontStrikethru = False
         .FontUnderline = False
      End With

      With Picture1
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .FontStrikethru = False
         .FontUnderline = False
      End With

      With picMain
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .FontStrikethru = False
         .FontUnderline = False
      End With

   End If
   Label4.Caption = "FontSize  " & Round(CurFont.Size)

   If CurFont.Bold = True Then
      If CurFont.Italic = True Then
         Label5.Caption = "Italic Bold"
       Else
         Label5.Caption = "Bold"
      End If

    ElseIf CurFont.Italic = True Then
      Label5.Caption = "Italic"
    Else
      Label5.Caption = "Regular"
   End If

   Label7.Caption = CurFont.Name
   Set cc = Nothing
End Sub

Private Sub cmdLoad_Click()

  Dim Filename As String

   Picture1.Cls
   sText = False
   sText1 = False
   txtAddText.Text = ""
   picItem(0).Visible = True            'make image visible
   Filename = FileDialog(Me, False, "Open Picture File", "*.jpg;*.jpeg;*.gif;*.bmp")
   If Filename = "" Then Exit Sub       'if cancel is pressed
   reload = Filename
   Image1.Picture = LoadPicture(Filename)
   ReDraw
End Sub

Private Sub cmdReLoad_Click()

   If reload = "" Then Exit Sub       'if cancel is pressed
   Image1.Picture = LoadPicture(reload)
   picItem(0).Visible = True
   ReDraw
End Sub

Private Sub cmdSaveJpeg_Click()
   JpegSave picMain
End Sub

Private Sub cmdSave_Click()

  Dim Filename As String

   picMain.Picture = picMain.Image
   Filename = FileDialog(Me, True, "Save Picture File", "*.bmp", , "bmp")
   If Filename = "" Then Exit Sub        'if Cancel is pressed
   SavePicture picMain.Picture, Filename
   MsgBox "Picture saved in " & Filename
End Sub

Private Sub cmdShape_Click()

  Dim Filename As String

   SetShape = True
   Filename = FileDialog(Me, False, "Open Picture File", "*.jpg;*.jpeg;*.gif;*.bmp", , , App.Path & _
      "\Shapes\")
   If Filename = "" Then Exit Sub
   reload = Filename
   picItem(0).Visible = True                     'make image visible
   Picture1.Picture = LoadPicture(Filename)
   Image1.Picture = LoadPicture(Filename)
   'get transparency color
   shptranscol = CLng(Picture1.Point(0, 0))   'upper left corner
   picItem(0).BackColor = shptranscol

   ReDraw
End Sub

Private Sub cmdTextColor_Click()

  Dim TheColor As Long

   Set CF = New CFDialog
   If CF.VBChooseColor(TheColor, , , , Me.hWnd) Then
      cmdTextColor.BackColor = TheColor
      Picture1.ForeColor = TheColor
      picMain.ForeColor = TheColor
   End If

   Set CF = Nothing
End Sub

Private Sub cmdTexture_Click()

  Dim Filename As String
  Dim x As Integer
  Dim y As Integer

   Filename = FileDialog(Me, False, "Open Picture File", "*.jpg;*.jpeg;*.gif;*.bmp", , , App.Path & _
      "\Textures\")
   If Filename = "" Then Exit Sub                'if cancel is pressed
   picMain.Picture = LoadPicture(Filename)
   'tile background

   For x = 0 To picMain.Width Step picMain.ScaleX(picMain.Picture.Width)
      For y = 0 To picMain.Height Step picMain.ScaleY(picMain.Picture.Height)
         picMain.PaintPicture picMain.Picture, x, y
      Next

   Next
   'update undo's
   nxt = nxt + 1
   Load imgUndo(nxt)
   imgUndo(nxt).Picture = picMain.Image
End Sub

Private Sub cmdUndo_Click()
   If nxt = 0 Then Exit Sub
   Unload imgUndo(nxt)
   nxt = nxt - 1
   picMain.Picture = imgUndo(nxt).Picture
End Sub

Private Sub HS3_Change()
   HS3_Scroll
End Sub

Private Sub HS3_Scroll()
   ReDraw
   Label10.Caption = HS3.Value
End Sub

Private Sub HScroll1_Change()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
   ReDraw
   Label6.Caption = HScroll1.Value
End Sub

Private Sub JpegSave(pic As PictureBox)

  Dim Filename As String

   pic.Picture = pic.Image
   On Error GoTo JpgErr
   Filename = FileDialog(Me, True, "Save Picture File", "*.jpeg", , "Jpeg")
   If Filename = "" Then Exit Sub                           'if Cancel is pressed
   If SaveJPEG(Filename, pic, Form1, True, 90) = True Then  ' save pic as Jpeg
      MsgBox "Jpeg saved in folder " & Filename
   End If

JpgErr:
   Exit Sub
End Sub

Private Sub mnuBkgdCol_Click()

  Dim TheColor As Long

   Set CF = New CFDialog
   If CF.VBChooseColor(TheColor, , , , Me.hWnd) Then
      picMain.BackColor = TheColor
   End If

   nxt = nxt + 1
   Load imgUndo(nxt)
   imgUndo(nxt).Picture = picMain.Image
   Set CF = Nothing
End Sub

Private Sub mnuExit_Click()

   Set JPEGclass = Nothing
   Set CurFont = Nothing
   Set CF = Nothing
   Set cc = Nothing
   Unload Me
End Sub

Private Sub mnuLoadPix_Click()
   cmdLoad_Click
End Sub

Private Sub mnuLoadShape_Click()
   cmdShape_Click
End Sub

Private Sub mnuNew_Click()

   picMain.Picture = LoadPicture()
   Image1.Picture = Image2.Picture         'load default picture "NoPicture"
   picItem(0).Visible = True
   HScroll1.Value = 0                      'set rotate angle to zero
   HS3.Value = 100
   ReDraw
End Sub

Private Sub mnuSaveBitmap_Click()
   cmdSave_Click
End Sub

Private Sub mnuSaveJpeg_Click()
   cmdSaveJpeg_Click
End Sub

Private Sub picItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   'arrow key pressed
   If KeyCode = 39 Then picItem(0).Left = picItem(0).Left + 1  'Right arrow
   If KeyCode = 37 Then picItem(0).Left = picItem(0).Left - 1  'Left arrow
   If KeyCode = 38 Then picItem(0).Top = picItem(0).Top - 1    'Up arrow
   If KeyCode = 40 Then picItem(0).Top = picItem(0).Top + 1    'Down arrow
   'update labels
   Label16.Caption = "X: " & picItem(0).Left
   Label17.Caption = "Y: " & picItem(0).Top
End Sub

Private Sub picItem_MouseDown(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)

   If Button = 1 Then
      OldX = x
      OldY = y
      TransParentPic picItem(0), picItem(0).Point(0, 0)    'make picture background transparent
      Label19.Visible = True
   End If

   If Button = 2 Then
      If SetShape = False Then
         If Picture1.Width > Picture1.Height Then
            DrawStdPictureRot picMain, picMain.hdc, picItem(0).Left + picItem(0).Width / 8, _
               picItem(0).Top + picItem(0).Height / 4, HScroll1.Value, Picture1.Picture
          Else
            DrawStdPictureRot picMain, picMain.hdc, picItem(0).Left + picItem(0).Width / 4.5, _
               picItem(0).Top + picItem(0).Height / 8, HScroll1.Value, Picture1.Picture
         End If

       Else
         TransparentBlt picMain.hdc, picItem(0).Left, picItem(0).Top, picItem(0).Width, _
            picItem(0).Height, picItem(0).hdc, 0, 0, picItem(0).Width, picItem(0).Height, _
            shptranscol
         SetShape = False
      End If

      picItem(0).Visible = False
      nxt = nxt + 1
      Load imgUndo(nxt)
      imgUndo(nxt).Picture = picMain.Image
      txtAddText.Text = ""
   End If
End Sub

Private Sub picItem_MouseMove(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)

   If Button = 1 Then
      picItem(0).Left = picItem(0).Left + (x - OldX)
      picItem(0).Top = picItem(0).Top + (y - OldY)
      Label16.Caption = "X: " & picItem(0).Left
      Label17.Caption = "Y: " & picItem(0).Top
   End If
End Sub

Private Sub picItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 1 Then
      TransParentPic picItem(0), vbBlue             'make picture background visible again
      Label19.Visible = False
   End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 1 Then
      If txtAddText.Text = "" Then Exit Sub
      tPosX = x
      tPosY = y
      sText1 = True
      nxt = nxt + 1
      Load imgUndo(nxt)
      imgUndo(nxt).Picture = picMain.Image
   End If

   ReDraw
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 1 Then
      If txtAddText.Text = "" Then Exit Sub
      tPosX = x
      tPosY = y
      sText = True
   End If
   ReDraw
End Sub

Private Sub ReDraw()

   ReszPic Picture1, Image1.Picture, HS3.Value, HS3.Value
   If sText = True Then                          'text on picture
      Picture1.CurrentX = tPosX
      Picture1.CurrentY = tPosY
      Picture1.Print txtAddText.Text
   End If

   If sText1 = True Then                         'text on main page
      picMain.CurrentX = tPosX
      picMain.CurrentY = tPosY
      picMain.Print txtAddText.Text
   End If

   Picture1.Picture = Picture1.Image             'render picture so it can be used
   picItem(0).Cls
   If Picture1.Width = Picture1.Height Then
      picItem(0).Width = Picture1.Width * 1.35
      picItem(0).Height = Picture1.Width * 1.35
      DrawStdPictureRot picItem(0), picItem(0).hdc, picItem(0).Width / 8, picItem(0).Width / 8, _
         HScroll1.Value, Picture1.Picture
         Exit Sub
   End If
   If Picture1.Width > Picture1.Height Then
      picItem(0).Width = Picture1.Width * 1.35    'make receiving box larger than image to allow for
                                                 '   rotating and not cut corners
      picItem(0).Height = Picture1.Width * 1.35  'make receiving box larger than image to allow for
                                                 '   rotating and not cut corners
      DrawStdPictureRot picItem(0), picItem(0).hdc, picItem(0).Width / 8, picItem(0).Height / 4, _
         HScroll1.Value, Picture1.Picture
    Else
      picItem(0).Width = Picture1.Height * 1.35     'make receiving box larger than image to allow
                                                  '   for rotating and not cut corners
      picItem(0).Height = Picture1.Height * 1.35   'make receiving box larger than image to allow
                                                  '   for rotating and not cut corners
      DrawStdPictureRot picItem(0), picItem(0).hdc, picItem(0).Width / 4.5, picItem(0).Height / 8, _
         HScroll1.Value, Picture1.Picture
   End If

   picMain.Refresh
   TransParentPic picItem(0), vbBlue
End Sub

Public Function SaveJPEG(ByVal Filename As String, _
                         pic As PictureBox, _
                         PForm As Form, _
                         Optional ByVal Overwrite As Boolean = True, _
                         Optional ByVal Quality As Byte = 90) As Boolean

  Dim m_Picture As IPictureDisp
  Dim m_DC As Long
  Dim m_Millimeter As Single

   m_Millimeter = PForm.ScaleX(100, vbPixels, vbMillimeters)
   Set m_Picture = pic
   m_DC = pic.hdc
   'this is not my code....from PSC
   'initialize class
   Set JPEGclass = New cJpeg
   'check there is image to save and the filename string is not empty

   If m_DC <> 0 And LenB(Filename) > 0 Then
      'check for valid quality
      If Quality < 1 Then Quality = 1
      If Quality > 100 Then Quality = 100
      'set quality
      JPEGclass.Quality = Quality
      'save in full color
      JPEGclass.SetSamplingFrequencies 1, 1, 1, 1, 1, 1
      'copy image from hDC

      If JPEGclass.SampleHDC(m_DC, CLng(m_Picture.Width / m_Millimeter), CLng(m_Picture.Height / _
         m_Millimeter)) = 0 Then
         'if overwrite is set and file exists, delete the file
         If Overwrite And LenB(Dir$(Filename)) > 0 Then Kill Filename
         'save file and return True if success
         SaveJPEG = JPEGclass.SaveFile(Filename) = 0
      End If

   End If
End Function

Private Sub TransParentPic(frm As PictureBox, Col As Long)
    Dim g, x, y, Rgn, Rgn1
    'TransParentPic Me, Me.Point(0, 0)
    'Point(0, 0) gets the left/top color of
    '     the form
    
    'Create a Main Region
    Rgn = CreateRectRgn(0, 0, 0, 0)
    
    For y = 0 To frm.ScaleHeight
        For x = 0 To frm.ScaleWidth
            'If color doesnt = Col then we will star
            '     t to create a line
            If frm.Point(x, y) <> Col Then
                g = x
                Do
                    x = x + 1
                Loop Until frm.Point(x, y) = Col Or x = frm.ScaleWidth + 1
                'Create a Second Region to add to the Ma
                '     in
                Rgn1 = CreateRectRgn(g, y, x, y + 1)
                'combined them
                Call CombineRgn(Rgn, Rgn, Rgn1, RGN_OR)
                'NOTE: IF YOU DO NOT DELETE THE REGION
                'IT WILL ERROR WINDOWS
                Call DeleteObject(Rgn1)
            End If
        Next
    Next
    'Set the New Region
    Call SetWindowRgn(frm.hWnd, Rgn, True)
    'NOTE: IF YOU DO NOT DELETE THE REGION
    'IT WILL ERROR WINDOWS
    Call DeleteObject(Rgn)
End Sub

