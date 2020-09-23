Attribute VB_Name = "ModTV"
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public TreeFilled As Boolean
Public LastSelected As Node

Function SpecialFolder(ByVal CSIDL As Long) As String
Dim R As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
R = SHGetSpecialFolderLocation(Form1.hwnd, CSIDL, IDL)
If R = NOERROR Then
    sPath = Space$(MAX_LENGTH)
    R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    If R Then
        SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End If
End Function
Public Sub FillTree(TV As TreeView)
TV.Nodes.add , , "IE", "Favorites", 3, 3
ListSubDirs TV, SpecialFolder(6) + "\", "IE"
ListFiles TV, SpecialFolder(6) + "\", "IE"
TV.Nodes(1).Expanded = True
TreeFilled = True
End Sub
Private Sub ListFiles(TV As TreeView, Path, parent)
On Error Resume Next
Dim Count, D(), i, DirName
DirName = Dir(Path, 6)
Do While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
        If LCase(Right(DirName, 3)) = "url" Then
            TV.Nodes.add parent, tvwChild, Path & DirName, Left(DirName, Len(DirName) - 4), 4, 4
        End If
    End If
    DirName = Dir
Loop
End Sub
Private Sub ListSubDirs(TV As TreeView, Path, parent)
On Error Resume Next
Dim Count, D(), i, DirName
DirName = Dir(Path, 16)
Do While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
        If GetAttr(Path + DirName) = 16 Then
            If (Count Mod 10) = 0 Then
                ReDim Preserve D(Count + 10)
            End If
            Count = Count + 1
            D(Count) = DirName
        End If
    End If
    DirName = Dir
Loop
For i = 1 To Count
    TV.Nodes.add parent, tvwChild, Path & D(i) & "\", D(i), 1, 2
    ListFiles TV, Path & D(i) & "\", Path & D(i) & "\"
    ListSubDirs TV, Path & D(i) & "\", Path & D(i) & "\"
Next
DoEvents
End Sub

Public Sub AddLink(TV As TreeView)
On Error Resume Next
Dim temp As String, temp1 As String
LockWindowUpdate TV.hwnd
temp1 = ""
temp1 = LastSelected.Key
If Right(temp1, 1) = "\" Then
    temp = SafeSave(temp1 + "PSC.url")
    Open temp For Binary Access Write As #1
    Put #1, , "[InternetShortcut]" & vbNewLine & "URL=http://www.planetsourcecode.com/vb/"
    Close #1
    Screen.MousePointer = 11
    ListSubDirs TV, SpecialFolder(6) + "\", "IE"
    ListFiles TV, SpecialFolder(6) + "\", "IE"
    TV.Refresh
    TV.Nodes(1).Expanded = True
    LastSelected.Expanded = True
    Screen.MousePointer = 0
ElseIf LCase(Right(temp1, 4)) = ".url" Then
    m = 0
    For y = Len(temp1) To 1 Step -1
    m = m + 1
    If Mid(temp1, y, 1) = "\" Then
        temp1 = Left(temp1, Len(temp1) - (m - 1))
        Exit For
    End If
    Next y
    temp = SafeSave(temp1 + "PSC.url")
    Open temp For Binary Access Write As #1
    Put #1, , "[InternetShortcut]" & vbNewLine & "URL=http://www.planetsourcecode.com/vb/"
    Close #1
    Screen.MousePointer = 11
    ListSubDirs TV, SpecialFolder(6) + "\", "IE"
    ListFiles TV, SpecialFolder(6) + "\", "IE"
    TV.Refresh
    TV.Nodes(1).Expanded = True
    LastSelected.Expanded = True
    Screen.MousePointer = 0
Else
    temp = SafeSave(SpecialFolder(6) + "\" + "PSC.url")
    Open temp For Binary Access Write As #1
    Put #1, , "[InternetShortcut]" & vbNewLine & "URL=http://www.planetsourcecode.com/vb/"
    Close #1
    Screen.MousePointer = 11
    ListFiles TV, SpecialFolder(6) + "\", "IE"
    TV.Refresh
    TV.Nodes(1).Expanded = True
    Screen.MousePointer = 0
End If
    For x = 1 To TV.Nodes.Count
        If TV.Nodes.Item(x).Key = temp Then
            TV.Nodes.Item(x).Selected = True
            Exit For
        End If
    Next x
LockWindowUpdate 0
End Sub

Public Sub NewFolder(TV As TreeView)
On Error Resume Next
temp1 = TV.SelectedItem.Key
LockWindowUpdate TV.hwnd
If temp1 = "IE" Then temp1 = SpecialFolder(6) + "\"
If Right(temp1, 1) = "\" Then
    temp = SafeSave(temp1 + "New Folder")
    temp2 = safesavename
    dontselect = True
    MkDir temp1 + temp2
    Screen.MousePointer = 11
    ListSubDirs TV, SpecialFolder(6) + "\", "IE"
    ListFiles TV, SpecialFolder(6) + "\", "IE"
    TV.Refresh
    Screen.MousePointer = 0
    LockWindowUpdate 0
    For x = 1 To TV.Nodes.Count
        If TV.Nodes.Item(x).Key = temp1 + temp2 + "\" Then
            TV.Nodes.Item(x).Selected = True
            Exit For
        End If
    Next x
    TV.StartLabelEdit
ElseIf LCase(Right(temp1, 4)) = ".url" Then
    m = 0
    For y = Len(temp1) To 1 Step -1
    m = m + 1
    If Mid(temp1, y, 1) = "\" Then
        temp1 = Left(temp1, Len(temp1) - (m - 1))
        Exit For
    End If
    Next y
    temp = SafeSave(temp1 + "New Folder")
    temp2 = safesavename
    dontselect = True
    MkDir temp1 + temp2
    Screen.MousePointer = 11
    ListSubDirs TV, SpecialFolder(6) + "\", "IE"
    ListFiles TV, SpecialFolder(6) + "\", "IE"
    TV.Refresh
    Screen.MousePointer = 0
    LockWindowUpdate 0
    For x = 1 To TV.Nodes.Count
        If TV.Nodes.Item(x).Key = temp1 + temp2 + "\" Then
            TV.Nodes.Item(x).Selected = True
            Exit For
        End If
    Next x
    TV.StartLabelEdit
End If
LockWindowUpdate 0
End Sub

Public Sub DeleteLink(TV As TreeView)
Dim temp As String
For x = 1 To TV.Nodes.Count
    If TV.Nodes.Item(x).Selected = True Then
        If TV.Nodes.Item(x).FullPath = "Favorites" Then
            MsgBox "You cannot delete this."
            Exit Sub
        End If
        temp = SpecialFolder(6) + "\" + Right(TV.Nodes.Item(x).FullPath, Len(TV.Nodes.Item(x).FullPath) - 10)
        If FileExists(temp + ".url") Then
            If MsgBox("Are you sure you wish to delete this link ?", vbYesNo, "Bobo Enterprises") = vbYes Then
                LockWindowUpdate TV.hwnd
                Kill temp + ".url"
                TV.Nodes.Remove x
                LockWindowUpdate 0
            Else
                Exit Sub
            End If
       ElseIf FileExists(temp) Then
            If GetAttr(temp) = 16 Then
            If MsgBox("Are you sure you wish to delete this folder ?", vbYesNo, "Bobo Enterprises") = vbYes Then
                LockWindowUpdate TV.hwnd
                ShellDeleteOne temp
                TV.Nodes.Remove x
                LockWindowUpdate 0
            Else
               Exit Sub
            End If
            End If
        End If
    Exit For
    End If
Next x
 LockWindowUpdate 0
End Sub

