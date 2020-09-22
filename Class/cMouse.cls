VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim OldPointer As Variant

Sub SetPointer(Optional NewPointer As MousePointerConstants = vbHourglass)
    ' Enforce a new mouse pointer.
    ' Store the original pointer only once.
    If IsEmpty(OldPointer) Then OldPointer = Screen.MousePointer
    Screen.MousePointer = NewPointer
End Sub

Private Sub Class_Terminate()
    ' Restore the original pointer when the object goes out of scope.
    ' Only if SetPointer had been actually called
    If Not IsEmpty(OldPointer) Then Screen.MousePointer = OldPointer
End Sub


