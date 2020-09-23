VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Starter Kit Volume 2"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   4290
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3413
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   2e6
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Listboxes"
      TabPicture(0)   =   "Form1.frx":00C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Scrollbars"
      TabPicture(1)   =   "Form1.frx":00E4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "TreeViews"
      TabPicture(2)   =   "Form1.frx":0100
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "KeyCodes"
      TabPicture(3)   =   "Form1.frx":011C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Examples"
      TabPicture(4)   =   "Form1.frx":0138
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   6615
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   34
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label10 
            Caption         =   $"Form1.frx":0154
            Height          =   1815
            Left            =   360
            TabIndex        =   40
            Top             =   1440
            Width           =   5895
         End
         Begin VB.Line Line1 
            X1              =   360
            X2              =   6360
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000009&
            X1              =   360
            X2              =   6360
            Y1              =   1095
            Y2              =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "ASCII ="
            Height          =   195
            Left            =   2280
            TabIndex        =   39
            Top             =   540
            Width           =   540
         End
         Begin VB.Label lblASCII 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3000
            TabIndex        =   38
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Shift ="
            Height          =   195
            Left            =   4200
            TabIndex        =   37
            Top             =   540
            Width           =   450
         End
         Begin VB.Label lblShift 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4800
            TabIndex        =   36
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Enter Character"
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   9
         Top             =   480
         Width           =   6615
         Begin VB.VScrollBar VScroll1 
            Height          =   2295
            Left            =   3960
            TabIndex        =   24
            Top             =   600
            Width           =   255
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   3000
            Width           =   3375
         End
         Begin VB.PictureBox Picture1 
            Height          =   2295
            Left            =   480
            ScaleHeight     =   2235
            ScaleWidth      =   3315
            TabIndex        =   21
            Top             =   600
            Width           =   3375
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   10000
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   10000
               _ExtentX        =   17648
               _ExtentY        =   17648
               _Version        =   393217
               Enabled         =   -1  'True
               Appearance      =   0
               TextRTF         =   $"Form1.frx":0316
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Label Label4 
            Caption         =   $"Form1.frx":03DE
            Height          =   2295
            Left            =   4440
            TabIndex        =   25
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Height          =   3495
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   6615
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   4200
            TabIndex        =   18
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Reset"
            Height          =   375
            Index           =   5
            Left            =   2400
            TabIndex        =   17
            Top             =   2880
            Width           =   1575
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Remove Dupes"
            Height          =   375
            Index           =   4
            Left            =   2400
            TabIndex        =   16
            Top             =   2400
            Width           =   1575
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Move Down"
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   15
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Move Up"
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Remove"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   13
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdList 
            Caption         =   "Add Item"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
         Begin VB.ListBox List1 
            Height          =   2790
            ItemData        =   "Form1.frx":049C
            Left            =   360
            List            =   "Form1.frx":049E
            TabIndex        =   11
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "The Listbox is one of the most simple controls but is also one of the most useful."
            Height          =   2175
            Left            =   4200
            TabIndex        =   20
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Item to Add :"
            Height          =   255
            Left            =   4200
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton Command5 
            Caption         =   "Picture Resizer"
            Height          =   375
            Left            =   480
            TabIndex        =   43
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Random Mouse Mover"
            Height          =   375
            Left            =   480
            TabIndex        =   41
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Resize a Picture correctly so you can save as a smaller picture."
            Height          =   375
            Left            =   2880
            TabIndex        =   44
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label Label11 
            Caption         =   "Press Ctl Q to stop - This example shows control of the cursor pos and registering Hotkeys."
            Height          =   495
            Left            =   2880
            TabIndex        =   42
            Top             =   600
            Width           =   3495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdTV 
            Caption         =   "Rename"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   4800
            TabIndex        =   32
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdTV 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   3000
            TabIndex        =   29
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton cmdTV 
            Caption         =   "Add PSC Link"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   3000
            TabIndex        =   28
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdTV 
            Caption         =   "Add Folder"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   4800
            TabIndex        =   27
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton cmdTV 
            Caption         =   "Load Favorites"
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   26
            Top             =   1440
            Width           =   1935
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   5280
            Top             =   1440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":04A0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":0A3A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":0FD4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":3788
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":5F3C
                  Key             =   "Add"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":64D8
                  Key             =   "Organise"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   2535
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   4471
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   441
            LabelEdit       =   1
            Style           =   1
            HotTracking     =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
         Begin VB.Label Label6 
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   3000
            Width           =   6135
         End
         Begin VB.Label Label5 
            Caption         =   $"Form1.frx":6A74
            Height          =   975
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   3375
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Code :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfile As String
Dim beforerename As String
Dim renameroot As String
Dim ImAdir As Boolean
Dim mShift As Integer
Dim mAscii As Integer
Private Sub cmdList_Click(Index As Integer)
Dim ListPos As Long, ListPos2 As Long
Dim nItem As Integer
LockWindowUpdate List1.hWnd
Select Case Index
    Case 0
        If Text1.Text <> "" Then List1.AddItem Text1.Text
        List1.Selected(List1.ListCount - 1) = True
        RTF.Text = "If Text1.Text <> " + Chr(34) + Chr(34) + " Then List1.AddItem Text1.Text" + vbCrLf + _
        "List1.Selected(List1.ListCount - 1) = True"
    Case 1
        If List1.ListIndex > 0 Then
            nItem = List1.ListIndex - 1
        Else
            nItem = 0
        End If
        List1.RemoveItem List1.ListIndex
        If List1.ListCount > 0 Then List1.Selected(nItem) = True
        RTF.Text = "List1.RemoveItem List1.ListIndex"
    Case 2
        With List1
          If .ListIndex < 0 Then Exit Sub
          nItem = .ListIndex
          If nItem = 0 Then Exit Sub
          .AddItem .Text, nItem - 1
          .RemoveItem nItem + 1
          .Selected(nItem - 1) = True
        End With
        RTF.Text = "With List1" + vbCrLf + _
        "  If .ListIndex < 0 Then Exit Sub" + vbCrLf + _
        "  nItem = .ListIndex" + vbCrLf + _
        "  If nItem = 0 Then Exit Sub" + vbCrLf + _
        "  .AddItem .Text, nItem - 1" + vbCrLf + _
        "  .RemoveItem nItem + 1" + vbCrLf + _
        "  .Selected(nItem - 1) = True" + vbCrLf + _
        "End With"
    Case 3
        With List1
          If .ListIndex < 0 Then Exit Sub
          nItem = .ListIndex
          If nItem = .ListCount - 1 Then Exit Sub
          .AddItem .Text, nItem + 2
          .RemoveItem nItem
          .Selected(nItem + 1) = True
        End With
        RTF.Text = "With List1" + vbCrLf + _
        "  If .ListIndex < 0 Then Exit Sub" + vbCrLf + _
        "  nItem = .ListIndex" + vbCrLf + _
        "  If nItem = .ListCount - 1 Then Exit Sub" + vbCrLf + _
        "  .AddItem .Text, nItem + 2" + vbCrLf + _
        "  .RemoveItem nItem" + vbCrLf + _
        "  .Selected(nItem + 1) = True" + vbCrLf + _
        "End With"
    Case 4
        For x = 0 To List1.ListCount - 1
            For y = 0 To List1.ListCount - 1
                ListPos = SendMessageByString(List1.hWnd, LB_FINDSTRINGEXACT, 0, List1.List(x))
                If ListPos <> x And ListPos <> -1 Then List1.RemoveItem ListPos
            Next y
        Next x
        RTF.Text = "For x = 0 To List1.ListCount - 1" + vbCrLf + _
        "    For y = 0 To List1.ListCount - 1" + vbCrLf + _
        "        ListPos = SendMessageByString(List1.hwnd, LB_FINDSTRINGEXACT, 0, List1.List(x))" + vbCrLf + _
        "        If ListPos <> x And ListPos <> -1 Then List1.RemoveItem ListPos" + vbCrLf + _
        "    Next y" + vbCrLf + _
        "Next x"
    Case 5
        ReloadList
End Select
LockWindowUpdate 0
End Sub

Private Sub cmdTV_Click(Index As Integer)
Select Case Index
    Case 0
        If TreeFilled = False Then FillTree TreeView1
        cmdTV(0).Enabled = False
        cmdTV(1).Enabled = True
        cmdTV(2).Enabled = True
        cmdTV(3).Enabled = True
        cmdTV(4).Enabled = True
    Case 1
        NewFolder TreeView1
    Case 2
        AddLink TreeView1
    Case 3
        DeleteLink TreeView1
    Case 4
        If TreeView1.SelectedItem.Key <> SpecialFolder(6) + "\" Then TreeView1.StartLabelEdit
End Select
End Sub

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText RTF.Text
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
frmCrazy.letsGo

End Sub

Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Command4_Click()
Form2.Show vbModal
End Sub

Private Sub Command5_Click()
frmResizer.Show
End Sub

Private Sub Form_Load()
ReloadList
VScroll1.Max = RichTextBox1.Height - Picture1.Height
HScroll1.Max = RichTextBox1.Width - Picture1.Width
RichTextBox1.Text = "This RichTextbox width and height properties are set" + vbCrLf + "to 10000. It has no scrollbars set (Scrollbars = 0)" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "" + vbCrLf + "See, you can use your own scrollbars to control the position of any control"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Public Sub ReloadList()
List1.Clear
For x = 0 To 19
    List1.AddItem "Listitem" + Trim(Str(x + 1))
Next x
For x = 0 To 19 Step 2
    List1.AddItem "Listitem" + Trim(Str(x + 1))
Next x
List1.Selected(0) = True
End Sub

Private Sub HScroll1_Change()
    RichTextBox1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
    RichTextBox1.Left = -HScroll1.Value
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 1
        RTF.Text = "'Assumes a Picturebox(Picture1), a RichTextbox(RichTextBox1) and a vertical(VScroll1) and horizontal(HScroll1) scrollbar" + vbCrLf + _
        "Private Sub Form_Load()" + vbCrLf + _
        "    VScroll1.Max = RichTextBox1.Height - Picture1.Height" + vbCrLf + _
        "    HScroll1.Max = RichTextBox1.Width - Picture1.Width" + vbCrLf + _
        "End Sub" + vbCrLf + vbCrLf + _
        "Private Sub HScroll1_Change()" + vbCrLf + _
        "    RichTextBox1.Left = -HScroll1.Value" + vbCrLf + _
        "End Sub" + vbCrLf + vbCrLf + _
        "Private Sub HScroll1_Scroll()" + vbCrLf + _
        "    RichTextBox1.Left = -HScroll1.Value" + vbCrLf + _
        "End Sub" + vbCrLf + vbCrLf + _
        "Private Sub VScroll1_Change()" + vbCrLf + _
        "    RichTextBox1.Top = -VScroll1.Value" + vbCrLf + _
        "End Sub" + vbCrLf + vbCrLf + _
        "Private Sub VScroll1_Scroll()" + vbCrLf + _
        "    RichTextBox1.Top = -VScroll1.Value" + vbCrLf + _
        "End Sub"
    Case 2
        RTF.Text = "Most of the code for the Treeview is found" + vbCrLf + _
        "in the module ModTV. It also calls functions" + vbCrLf + _
        "in ModGP. Use these functions in conjunction" + vbCrLf + _
        "with the Treeview1 entries on Form1."
    Case 3
        Text2.SetFocus
        RTF.Text = "Private Sub Text2_KeyPress(KeyAscii As Integer)" + vbCrLf + _
        "    mAscii = KeyAscii" + vbCrLf + _
        "End Sub"
    
    Case Else
        RTF.Text = ""
End Select
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
    cmdList(0).Enabled = True
Else
    cmdList(0).Enabled = False
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Text2.SelStart = 0
Text2.SelLength = 1
mShift = Shift
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
mAscii = KeyAscii
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
lblASCII = mAscii
lblShift = mShift
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim temp As String
Set fred = TreeView1.HitTest(x, y)
If fred Is Nothing Then Exit Sub
Set LastSelected = fred

End Sub
Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
If ImAdir = False Then 'ImAdir variable was set in the sub TreeView1_BeforeLabelEdit
    If FileExists(beforerename + ".url") Then
        temp = SafeSave(renameroot + NewString + ".url")
        NewString = ChangeExt(safesavename)
        Name beforerename + ".url" As temp
        TreeView1.SelectedItem.Key = temp
    End If
Else
    If FileExists(beforerename) Then
        temp = SafeSave(renameroot + NewString)
        NewString = safesavename
        Name beforerename As temp
        TreeView1.SelectedItem.Key = temp + "\"
    End If
End If
dontselect = False
End Sub
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
dontselect = True
If TreeView1.SelectedItem.Key = "IE" Then
Cancel = 1
Exit Sub
End If
If Right(TreeView1.SelectedItem.Key, 1) = "\" Then
    ImAdir = True
Else
    ImAdir = False
End If
For x = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes.Item(x).Selected = True Then
        beforerename = SpecialFolder(6) + "\" + Right(TreeView1.Nodes.Item(x).FullPath, Len(TreeView1.Nodes.Item(x).FullPath) - 10)
        renameroot = Left(beforerename, Len(beforerename) - Len(TreeView1.Nodes.Item(x).Text))
        Exit For
    End If
Next x
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim temp As String
temp = Node.Key
If Right(temp, 1) = "\" Or temp = "IE" Then Exit Sub
Label6 = ReadINI(temp, "InternetShortcut", "URL")

End Sub

Private Sub VScroll1_Change()
    RichTextBox1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
    RichTextBox1.Top = -VScroll1.Value
End Sub



