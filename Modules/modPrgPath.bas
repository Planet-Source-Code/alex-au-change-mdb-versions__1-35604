Attribute VB_Name = "modPrgPath"
Option Explicit

Const MAX_PATH = 206

Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Function GetPrgPath() As String
    'App.Path is not always reliable, especially on a network.
    'Try using GetModuleFileName API call to avoid the problem

    Dim lngFileHandle   As Long
    Dim lngReturn       As Long
    Dim strFilePath     As String

    strFilePath = Space$(MAX_PATH)
    lngFileHandle = GetModuleHandle(App.EXEName)
    lngReturn = GetModuleFileName(lngFileHandle, strFilePath, MAX_PATH)
    GetPrgPath = strFilePath
End Function

