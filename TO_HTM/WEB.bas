Attribute VB_Name = "WEB"
Option Explicit

Private Declare Function GetActiveWindow Lib "USER32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub GoToUrl(Url As String, Optional Subject As String)
    If IsMissing(Subject) Then
        Subject = ""
    Else
        Subject = "?subject=" & Subject
    End If
    If InStr(1, Url, "@") <> 0 Then
      Url = "mailto:" & Url & Subject
    Else
      If InStr(1, UCase(Url), "HTTP://", 1) = 0 Then
        Url = Url
      End If
    End If
    ShellExecute GetActiveWindow(), "Open", Url, "", 0&, 1
End Sub


