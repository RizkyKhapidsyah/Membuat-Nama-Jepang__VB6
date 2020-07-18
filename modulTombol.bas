Attribute VB_Name = "modulTomboldanLink"
Option Explicit
Private Type tagInitCommonControlsEx
  Ukuran As Long
  icece As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iceks As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function KontrolXP() As Boolean
On Error Resume Next
Dim iceks As tagInitCommonControlsEx
With iceks
  .Ukuran = Len(iceks)
  .icece = ICC_USEREX_CLASSES
End With
InitCommonControlsEx iceks
KontrolXP = CBool(Err = 0)
End Function
