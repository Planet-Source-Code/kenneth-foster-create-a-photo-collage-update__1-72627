Attribute VB_Name = "modOpenSave"
Option Explicit

Private Type OPENFILENAME
   lStructSize       As Long
   hwndOwner         As Long
   hInstance         As Long
   lpstrFilter       As String
   lpstrCustomFilter As String
   nMaxCustFilter    As Long
   nFilterIndex      As Long
   lpstrFile         As String
   nMaxFile          As Long
   lpstrFileTitle    As String
   nMaxFileTitle     As Long
   lpstrInitialDir   As String
   lpstrTitle        As String
   flags             As Long
   nFileOffset       As Integer
   nFileExtension    As Integer
   lpstrDefExt       As String
   lCustData         As Long
   lpfnHook          As Long
   lpTemplateName    As String
End Type

Private Type BrowseInfo
   hwndOwner         As Long
   pIDLRoot          As Long
   pszDisplayName    As Long
   lpszTitle         As Long
   ulFlags           As Long
   lpfnCallback      As Long
   lParam            As Long
   iImage            As Long
End Type

Private Const OFN_READONLY             As Long = &H1
Private Const OFN_OVERWRITEPROMPT      As Long = &H2
Private Const OFN_HIDEREADONLY         As Long = &H4
Private Const OFN_NOCHANGEDIR          As Long = &H8
Private Const OFN_SHOWHELP             As Long = &H10
Private Const OFN_ENABLEHOOK           As Long = &H20
Private Const OFN_ENABLETEMPLATE       As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_NOVALIDATE           As Long = &H100
Private Const OFN_ALLOWMULTISELECT     As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT   As Long = &H400
Private Const OFN_PATHMUSTEXIST        As Long = &H800
Private Const OFN_FILEMUSTEXIST        As Long = &H1000
Private Const OFN_CREATEPROMPT         As Long = &H2000
Private Const OFN_SHAREAWARE           As Long = &H4000
Private Const OFN_NOREADONLYRETURN     As Long = &H8000
Private Const OFN_NOTESTFILECREATE     As Long = &H10000
Private Const OFN_NONETWORKBUTTON      As Long = &H20000
Private Const OFN_NOLONGNAMES          As Long = &H40000
Private Const OFN_EXPLORER             As Long = &H80000
Private Const OFN_NODEREFERENCELINKS   As Long = &H100000
Private Const OFN_LONGNAMES            As Long = &H200000

Private Const OFN_SHAREFALLTHROUGH     As Long = 2
Private Const OFN_SHARENOWARN          As Long = 1
Private Const OFN_SHAREWARN            As Long = 0

Private Const MAX_PATH                 As Long = 260

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
      Alias "GetOpenFileNameA" ( _
      pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
      Alias "GetSaveFileNameA" ( _
      pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long


Public Function FileDialog(FormObject As Form, _
                           SaveDialog As Boolean, _
                           ByVal Title As String, _
                           ByVal Filter As String, _
                           Optional ByVal FileName As String, _
                           Optional ByVal Extention As String, _
                           Optional ByVal InitDir As String) As String

  Dim OFN   As OPENFILENAME
  Dim r     As Long
  Dim l As Long

   If Len(FileName) > MAX_PATH Then Call MsgBox("Filename Length Overflow", vbExclamation, _
      App.Title + " - FileDialog Function"): Exit Function

   FormObject.Enabled = False
   FileName = FileName + String(MAX_PATH - Len(FileName), 0)

   With OFN
      .lStructSize = Len(OFN)
      .hwndOwner = FormObject.hWnd
      .hInstance = App.hInstance
      .lpstrFilter = Replace(Filter, "|", vbNullChar)
      .lpstrFile = FileName
      .nMaxFile = MAX_PATH
      .lpstrFileTitle = Space$(MAX_PATH - 1)
      .nMaxFileTitle = MAX_PATH
      .lpstrInitialDir = InitDir
      .lpstrTitle = Title
      .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
      .lpstrDefExt = Extention
   End With

   l = GetTickCount

   If SaveDialog Then r = GetSaveFileName(OFN) Else r = GetOpenFileName(OFN)

   If GetTickCount - l < 20 Then
      OFN.lpstrFile = ""
      If SaveDialog Then r = GetSaveFileName(OFN) Else r = GetOpenFileName(OFN)
   End If

   If r = 1 Then FileDialog = Left$(OFN.lpstrFile, InStr(1, OFN.lpstrFile + vbNullChar, vbNullChar) _
      - 1)
   FormObject.Enabled = True

End Function

