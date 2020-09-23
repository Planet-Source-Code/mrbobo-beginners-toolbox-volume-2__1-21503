VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "All bugs and error messages provided at no extra cost."
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BOBO ENTERPRISES   2001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
