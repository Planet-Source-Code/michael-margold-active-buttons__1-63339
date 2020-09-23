Attribute VB_Name = "HyperLink"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function Link()
  Link = ShellExecute(0&, vbNullString, "http://www.soft-collection.com", vbNullString, vbNullString, vbNormalFocus)
End Function
