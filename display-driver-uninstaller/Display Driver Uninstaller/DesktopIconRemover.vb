Option Strict Off

Imports System.IO

Public Class DesktopIconRemover
    Public Shared Function GetShortcutTargetPath(shortcutPath As String) As String
        If System.IO.File.Exists(shortcutPath) Then
            Dim shell As Object = CreateObject("WScript.Shell")
            Dim shortcut As Object = shell.CreateShortcut(shortcutPath)
            Return CStr(shortcut.TargetPath)
        Else
            Throw New FileNotFoundException("Shortcut file not found.", shortcutPath)
        End If
    End Function
End Class