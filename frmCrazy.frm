VERSION 5.00
Begin VB.Form frmCrazy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobo Crazy Mouse - Ctl-Q to exit"
   ClientHeight    =   30
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3840
   Icon            =   "frmCrazy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   840
   End
End
Attribute VB_Name = "frmCrazy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'hotkey api
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
'moving cursor api
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Dim P As POINTAPI
Dim mx As Integer
Dim my As Integer
Dim curx As Integer
Dim cury As Integer
Dim tempx As Integer
Dim tempy As Integer
Dim temp1x As Integer
Dim temp1y As Integer
Dim xachieved As Boolean
Dim z As Long



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
bCancel = True
Call UnregisterHotKey(Me.hWnd, &HBFFF&)

End Sub

Private Sub Timer1_Timer()
z = GetCursorPos(P)
If xachieved = True Then
    NextRandPos
Else
    SetChangeFactor
        tempx1 = P.x + tempx
        If tempx1 = mx Then xachieved = True
        If tempx1 + tempx = mx Then xachieved = True
        If tempx1 - tempx = mx Then xachieved = True
        tempy1 = P.y + tempy
        If tempy1 = my Then xachieved = True
        If tempy1 + tempy = my Then xachieved = True
        If tempy1 - tempy = my Then xachieved = True
        SetCursorPos tempx1, tempy1
End If
End Sub
Function RandomNumber(Max As Double, min As Double)
On Error GoTo error
Randomize Timer
RandomNumber = Int((Max - min + 1) * Rnd + min)
Exit Function
error:  MsgBox Err.Description, vbExclamation, "Error"
End Function


Public Sub NextRandPos()
mx = RandomNumber(Screen.Width / Screen.TwipsPerPixelX, 1)
my = RandomNumber(Screen.Height / Screen.TwipsPerPixelY, 1)
SetChangeFactor
End Sub

Public Sub SetChangeFactor()
z = GetCursorPos(P)
curx = P.x
cury = P.y
xachieved = False
tempx = Int((mx - curx) / 20)
tempy = Int((my - cury) / 20)
If tempx = 0 Or tempy = 0 Then xachieved = True
End Sub
Private Sub ProcessMessages()
    Dim Message As Msg
    Do While Not bCancel
        WaitMessage
        If PeekMessage(Message, Me.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            Timer1.Enabled = False
            Unload Me
            Exit Do
        End If
        DoEvents
    Loop
End Sub


Public Sub letsGo()
Dim ret As Long
bCancel = False
ret = RegisterHotKey(Me.hWnd, &HBFFF&, MOD_CONTROL, vbKeyQ)
NextRandPos
Timer1.Enabled = True
ProcessMessages

End Sub
