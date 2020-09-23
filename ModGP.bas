Attribute VB_Name = "ModGP"
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_ALLOWUNDO = &H40
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4&
Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim FO_FUNC As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long
Public Const LB_FINDSTRINGEXACT = &H1A2
Public safesavename As String
Public ret As String
Public Retlen As String
Public Function SafeSave(Path As String) As String
Dim mPath As String, mTemp As String, mFile As String, mExt As String, m As Integer
On Error Resume Next
mPath = Mid$(Path, 1, InStrRev(Path, "\")) 'Path only
mname = Mid$(Path, InStrRev(Path, "\") + 1) 'File only
mFile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1) 'File only - no extension
If mFile = "" Then mFile = mname
mExt = Mid$(mname, InStrRev(mname, ".")) 'Extension only
mTemp = ""
Do
    If Not FileExists(mPath + mFile + mTemp + mExt) Then
        SafeSave = mPath + mFile + mTemp + mExt
        safesavename = mFile + mTemp + mExt
        Exit Do
    End If
    m = m + 1
    mTemp = Right(Str(m), Len(Str(m)) - 1)
Loop
End Function
Function FileExists(ByVal Filename As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(Filename)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & error, MB_OK, "Error"
                End
            End If
    End Select
End Function
Public Sub ShellDeleteOne(sfile As String)
     On Error Resume Next
   Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim R As Long
    FOF_FLAGS = FOF_SILENT Or FOF_NOCONFIRMATION
sfile = sfile & Chr$(0)
With SHFileOp
  .wFunc = FO_DELETE
  .pFrom = sfile
  .fFlags = FOF_FLAGS
End With
R = SHFileOperation(SHFileOp)
End Sub


Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
Dim temp As String
temp = Mid$(filepath, 1, InStrRev(filepath, "."))
temp = Left(temp, Len(temp) - 1)
If newext <> "" Then newext = "." + newext
ChangeExt = temp + newext
End Function

Public Function ReadINI(Filename As String, Section As String, Key As String)
ret = Space$(255)
Retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), Filename)
ret = Left$(ret, Retlen)
ReadINI = ret
End Function

