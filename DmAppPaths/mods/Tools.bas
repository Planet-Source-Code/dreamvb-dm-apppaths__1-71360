Attribute VB_Name = "Tools"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Listview Consts
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Public ExeName As String
Public ExePath As String
Public EditOp As Integer
Public ButtonPress As VbMsgBoxResult

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    If Len(lzFileName) = 0 Then Exit Function
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Function GetPathFormFile(ByVal lFilePath As String) As String
Dim sPos As Integer
    sPos = InStrRev(lFilePath, "\", Len(lFilePath), vbBinaryCompare)
    If (sPos > 0) Then
        GetPathFormFile = Left$(lFilePath, sPos)
    Else
        GetPathFormFile = lFilePath
    End If
End Function

Public Sub lvSizeColumns(lv As ListView)
Dim Counter As Long
    'Resizes Listview Column Headers.
    For Counter = 0 To (lv.ColumnHeaders.Count - 1)
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, Counter, _
        ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next Counter
End Sub

Public Function RunApp(iHwnd As Long, OpenOp As String, FileName As String) As Long
    RunApp = ShellExecute(iHwnd, OpenOp, FileName, "", "", 1)
End Function
