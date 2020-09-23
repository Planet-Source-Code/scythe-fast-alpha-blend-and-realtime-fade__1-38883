VERSION 5.00
Begin VB.Form FastAlpha 
   Caption         =   "Fast Alpha Blend and Realtime Fade"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   803
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Fade Out"
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   7
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fade In"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Autoblend left to right"
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Autoblend left to right"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2400
      Max             =   100
      TabIndex        =   3
      Top             =   5760
      Value           =   50
      Width           =   7335
   End
   Begin VB.PictureBox Result 
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   5625
      Left            =   4080
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   2
      Top             =   0
      Width           =   3960
   End
   Begin VB.PictureBox PicRight 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   5625
      Left            =   8040
      Picture         =   "Alpha.frx":0000
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   1
      Top             =   0
      Width           =   3960
   End
   Begin VB.PictureBox PicLeft 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      Height          =   5625
      Left            =   120
      Picture         =   "Alpha.frx":4B57
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   0
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "FastAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fast Alpha
'A simple and fast Alphablend routine
'© Scythe
'Scythe@cablenet.de

'Needed for Speedtest
'Private Declare Function GetTickCount Lib "kernel32" () As Long

'Needed for DIB
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbgreen As Byte
 rgbred As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

'2 Lookup Tables for faster blending
Dim LookUpPercent(100, 255) As Byte  'Holds the percent calculation
Dim LookUpColorAdd(255, 255) As Byte  'Holds the Add Colors math


Private Sub Form_Load()
 Dim i As Integer
 Dim f As Integer
 Dim a As Long

 If InIde = True Then
  MsgBox "Compile me to see the real speed", vbInformation, "Fast Alpha"
 End If

 'Fill the tables
 For i = 0 To 255
  For f = 0 To 255
   a = i + f
   If a > 255 Then a = 255
   LookUpColorAdd(i, f) = a
  Next f
 Next i

 For i = 0 To 100
  For f = 0 To 255
   LookUpPercent(i, f) = f / 100 * i
  Next f
 Next i

 HScroll1_Change

End Sub

Private Sub Command1_Click(Index As Integer)
 Dim buf1() As RGBQUAD
 Dim buf2() As RGBQUAD
 Dim buf3() As RGBQUAD
 Dim LeftPercent As Long
 Dim RightPercent As Long
 Dim x As Long
 Dim y As Long
 Dim i As Long
 Dim lwidht As Long
 Dim lheight As Long
 
 

 'Get the 2 Pictures
 Pic2Array PicLeft, buf1()
 Pic2Array PicRight, buf2()



 'Create the buffer for our new Picture
 ReDim buf3(0 To Result.ScaleWidth - 1, 0 To Result.ScaleHeight - 1)

 Me.MousePointer = 11

 'For Autoblend i only show every second %
 For i = 0 To 100 Step 2
  'Get the Percent for each picture
  If Index = 0 Then
   LeftPercent = 100 - i
   RightPercent = i
  Else
   LeftPercent = i
   RightPercent = 100 - i
  End If

  'Create the new picture
  For x = 0 To Result.ScaleWidth - 1
   For y = 0 To Result.ScaleHeight - 1
    buf3(x, y).rgbBlue = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbBlue), LookUpPercent(RightPercent, buf2(x, y).rgbBlue))
    buf3(x, y).rgbgreen = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbgreen), LookUpPercent(RightPercent, buf2(x, y).rgbgreen))
    buf3(x, y).rgbred = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbred), LookUpPercent(RightPercent, buf2(x, y).rgbred))
   Next y
  Next x

  Array2Pic Result, buf3()
  Result.Refresh

 Next i

 Me.MousePointer = 0
End Sub



Private Sub HScroll1_Change()
 Dim buf1() As RGBQUAD
 Dim buf2() As RGBQUAD
 Dim buf3() As RGBQUAD
 Dim LeftPercent As Byte
 Dim RightPercent As Byte
 Dim x As Long
 Dim y As Long

 'Get the 2 Pictures
 Pic2Array PicLeft, buf1()
 Pic2Array PicRight, buf2()

 'Create the buffer for our new Picture
 ReDim buf3(0 To Result.ScaleWidth - 1, 0 To Result.ScaleHeight - 1)

 'Get the Percent for each picture
 LeftPercent = 100 - HScroll1.Value
 RightPercent = HScroll1.Value

 'Create the new picture
 For x = 0 To Result.ScaleWidth - 1
  For y = 0 To Result.ScaleHeight - 1
   buf3(x, y).rgbBlue = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbBlue), LookUpPercent(RightPercent, buf2(x, y).rgbBlue))
   buf3(x, y).rgbgreen = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbgreen), LookUpPercent(RightPercent, buf2(x, y).rgbgreen))
   buf3(x, y).rgbred = LookUpColorAdd(LookUpPercent(LeftPercent, buf1(x, y).rgbred), LookUpPercent(RightPercent, buf2(x, y).rgbred))
  Next y
 Next x
 Array2Pic Result, buf3()
 Result.Refresh
End Sub

'Fade in/Out
Private Sub Command2_Click(Index As Integer)
Dim buf1() As RGBQUAD
 
 Dim buf3() As RGBQUAD
 Dim LeftPercent As Long
 Dim RightPercent As Long
 Dim x As Long
 Dim y As Long
 Dim i As Long
 Dim lwidht As Long
 Dim lheight As Long
 Dim StartPercent As Long
 Dim EndPercent As Long
 Dim FadeStep As Long
 
 
 'Create the buffer for our new Picture
 ReDim buf3(0 To Result.ScaleWidth - 1, 0 To Result.ScaleHeight - 1)

 Me.MousePointer = 11
  
  'Fade In or Out
  'thats the Question :o)
  If Index = 0 Then
   StartPercent = 0
   EndPercent = 100
   FadeStep = 2
   'Get the Picture
   Pic2Array PicLeft, buf1()
  Else
   StartPercent = 100
   EndPercent = 0
   FadeStep = -2
   'Get the Picture
   Pic2Array PicRight, buf1()
  End If

 'For Autoblend i only show every second %
 For i = StartPercent To EndPercent Step FadeStep

  'Create the new picture
  'For FadeIn / Out we oly ned the
  'LookUpPercent so its much faster than Alphablend
  For x = 0 To Result.ScaleWidth - 1
   For y = 0 To Result.ScaleHeight - 1
    buf3(x, y).rgbBlue = LookUpPercent(i, buf1(x, y).rgbBlue)
    buf3(x, y).rgbgreen = LookUpPercent(i, buf1(x, y).rgbgreen)
    buf3(x, y).rgbred = LookUpPercent(i, buf1(x, y).rgbred)
   Next y
  Next x

  Array2Pic Result, buf3()
  Result.Refresh

 Next i

 Me.MousePointer = 0
End Sub


'Convert Picture to Array
Private Sub Pic2Array(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 ReDim PicArray(0 To PicBox.ScaleWidth - 1, 0 To PicBox.ScaleHeight - 1)
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 'Now get the Picture
 GetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS
End Sub

'Convert Array to Picture
Private Sub Array2Pic(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = PicBox.ScaleWidth
 .biHeight = PicBox.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = PicBox.ScaleWidth * PicBox.ScaleHeight
 End With
 SetDIBits PicBox.hdc, PicBox.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicArray(0, 0), Binfo, DIB_RGB_COLORS
End Sub

'Test if we are in ide or compiled mode
Private Function InIde() As Boolean
 On Error GoTo DivideError
 Debug.Print 1 / 0
 Exit Function
DivideError:
 InIde = True
End Function

