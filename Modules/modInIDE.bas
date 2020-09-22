Attribute VB_Name = "modInIDE"
Option Explicit

Private Declare Function GetModuleFileName Lib "kernel32" _
    Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long _
    ) As Long

Public Function InIDE() As Boolean
    InIDE = False
    Dim strFileName As String
    Dim lngCount As Long

    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)

    InIDE = (UCase(Right(strFileName, 7)) Like "VB?.EXE")
End Function

