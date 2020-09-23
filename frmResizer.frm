VERSION 5.00
Begin VB.Form frmResizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resize Picture"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Half Size"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   360
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4170
      Left            =   120
      Picture         =   "frmResizer.frx":0000
      ScaleHeight     =   274
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   418
      TabIndex        =   0
      Top             =   120
      Width           =   6330
   End
End
Attribute VB_Name = "frmResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Requires 2 PictureBoxes - one of which is set to Visible = False
'both to Scalemode = 3
'API declares
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Constant for copying
Const SRCCOPY = &HCC0020

Private Sub ImgResize(src As PictureBox, TmpPic As PictureBox, mWidth As Long, mHeight As Long)
src.AutoRedraw = False 'add a line at the end of sub to return to True if you wish
TmpPic.AutoRedraw = True
TmpPic.Height = mHeight 'set our hidden picturebox
TmpPic.Width = mWidth 'to the desired size
'stretch our original picture to fit
StretchBlt TmpPic.hdc, 0, 0, mWidth, mHeight, src.hdc, 0, 0, src.Width, src.Height, SRCCOPY
'save the stretched picture to a tempfile
'or to a renamed file in which case reload to
'original picturebox but dont delete
SavePicture TmpPic.Image, App.Path + "\tempimg.bmp"
'reload the original picturebox
src.AutoSize = True
src.Picture = LoadPicture(App.Path + "\tempimg.bmp")
'remove the temp file
Kill App.Path + "\tempimg.bmp"
'clear our hidden picturebox
TmpPic.Picture = LoadPicture()
'example call to halve the size
'ImgResize Picture1, Picture2, Picture1.Width / 2, Picture1.Height / 2
End Sub


Private Sub Command1_Click()
ImgResize Picture1, Picture2, Picture1.Width / 2, Picture1.Height / 2
End Sub

