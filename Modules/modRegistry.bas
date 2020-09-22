Attribute VB_Name = "modRegistry"
'Copyright Tretyakov Konstantin (kt_ee@yahoo.com)
'You may use this code for free, if you give me some credit
'At least remember, thet it is not fair to put your name on what you didn't do

'And I would surely appreciate, if you mail me the program (or link to it)
'you created, using this code, (or if you somehow modified this one)

Option Explicit
'General Declares
Public NeedRestart As Boolean


'Constants
Public Const BmpView = "SOFTWARE\Classes\Paint.Picture\DefaultIcon" 'in local machine( and in classes root\paint.pi...)
Public Const MinAnim = "Control Panel\desktop\WindowMetrics" 'in  currentuser(users\.default\contr...)
Public Const MenuDelay = "Control Panel\desktop" 'in cur user (in users\.Default\contr...)
Public Const WinInfo = "SOFTWARE\Microsoft\Windows\CurrentVersion" 'in loc machine
Public Const RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"
Public Const ControlPanel = "{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Public Const PrintersReg = "{2227A280-3AEA-1069-A2DE-08002B30309D}" 'all in local machine and in clasesroot\clsid
Public Const DialUp = "{992CFFA0-F557-101A-88EC-00DD010CCC48}"
Public Const MainRoot = "SOFTWARE\Classes\CLSID\"
Public Const SecRoot = "Software\Classes\CLSID\"
Public Const CLSID = "CLSID\"
'Public Const IE = "Software\Microsoft\Internet Explorer\Main"
'Public Const OL = "Identities\{CBB57CC0-1395-11D3-9505-444553540001}\Software\Microsoft\Outlook Express\5.0"
Public Const Key = "SOFTWARE\Microsoft\Windows\CurrentVersion\"
Public Const Setup = "SOFTWARE\Microsoft\Windows\CurrentVersion\setup"


Public Const CPSN = "\Control Panel."
Public Const PRNSN = "\Printers."
Public Const DUPNSN = "\Dial-Up Networking."
Public Const RBSN = "\Recycle Bin."

    Public SysRoot As String

'Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'*******************Registry Access********************************************
'Functions here from ADVAPI32.DLL

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lValueLen As Long) As Long
'Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long           ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


Public Enum APIRegistryRoots
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

'*********Reg. Value Types*********
Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

Public Enum APIRegistryValueTypes
    NoType = REG_NONE   ' No value type
    StringType = REG_SZ ' Unicode nul terminated string
    ExpandedStringType = REG_EXPAND_SZ  ' Unicode nul terminated string
    BinaryType = REG_BINARY ' Free form binary
    DwordType = REG_DWORD ' 32-bit number
End Enum

'***********************************

Private Const ERROR_SUCCESS = 0

Private Const KEY_ALL_ACCESS = 983103


Private Const regErrOpenKey = 1
Private Const regErrQueryKey = 2
Private Const regErrCreateKey = 3
Private Const regErrSetKey = 4



''******Working with INI files***********
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$) As Integer
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpValue$, ByVal lpFileName$) As Integer
'
'
'Public Sub SaveINIString(ByVal INIFileName As String, ByVal Section As String, ByVal Key As String, Optional ByVal Value As String = "")
'    WritePrivateProfileString Section, Key, Value, INIFileName
'End Sub

'Public Function GetINIString(ByVal INIFileName As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String) As String
'    Dim lRes As Long, sBuf As String
'    Const lBufLen = 255
'    sBuf = String(lBufLen, 0)
'    lRes = GetPrivateProfileString(Section, Key, Default, sBuf, lBufLen, INIFileName)
'    If lRes Then
'        sBuf = Left(sBuf, InStr(sBuf, Chr(0)) - 1)
'        GetINIString = sBuf
'    End If
'End Function

'Public Function EnumStringKeys(ByVal hKey As APIRegistryRoots, ByVal SubKey As String, Optional KeyCount As Long) As Variant
'    Dim lRes As Long, lHandle As Long, lSubKeyNo As Long
'    Dim sRetVal As String, lRetLen As Long
'    Dim vResult() As String, NulCharPos As Long
'    lRes = RegOpenKey(hKey, SubKey, lHandle)
'    If lRes <> ERROR_SUCCESS Then
'        Err.Raise vbObjectError + regErrOpenKey, "EnumStringKeys", "Win32 API function failed to open registry key '" & SubKey & "'."
'        Exit Function
'    End If
'    lSubKeyNo = 0
'    Do Until lRes <> ERROR_SUCCESS
'        lRetLen = 1028
'        sRetVal = Space(lRetLen)
'        lRes = RegEnumKey(lHandle, lSubKeyNo, sRetVal, lRetLen)
'        NulCharPos = InStr(sRetVal, Chr(0))
'        If NulCharPos > 0 Then
'            ReDim Preserve vResult(lSubKeyNo)
'            sRetVal = Left(sRetVal, NulCharPos - 1)
'            vResult(lSubKeyNo) = sRetVal
'        End If
'        lSubKeyNo = lSubKeyNo + 1
'    Loop
'    KeyCount = lSubKeyNo - 1
'    EnumStringKeys = vResult
'End Function

Function GetStringKey(ByVal hKey As APIRegistryRoots, ByVal SubKey As String, Optional ByVal ValueName As String = "") As String
    On Error GoTo ErrFound
    Dim lRes As Long, lRetLen As Long, sRetVal As String
    Dim lHandle As Long
    lRes = RegOpenKey(hKey, SubKey, lHandle)
    If lRes <> ERROR_SUCCESS Then
        'Oops
        Err.Raise vbObjectError + regErrOpenKey, "GetStringKey", "Win 32 API function failed to open registry key '" & SubKey & "'."
        Exit Function
    Else
        lRetLen = 400
        sRetVal = Space(lRetLen)
        lRes = RegQueryValueEx(lHandle, ValueName, 0, REG_SZ, ByVal sRetVal, lRetLen)
        If lRes <> ERROR_SUCCESS Then
            'Oops
            Err.Raise vbObjectError + regErrQueryKey, "GetStringKey", "Win32 API function failed to get registry key value: '" & SubKey & "'."
            RegCloseKey lHandle
            Exit Function
        Else
            'OK...
            sRetVal = IIf(InStr(sRetVal, Chr(0)) <> 0, Left(sRetVal, InStr(sRetVal, Chr(0)) - 1), RTrim(sRetVal))
        End If
        RegCloseKey lHandle
    End If
    GetStringKey = sRetVal
    Exit Function
ErrFound:
    GetStringKey = ""
End Function

'Function GetDwordKey&(ByVal hKey As APIRegistryRoots, ByVal SubKey As String, Optional ByVal ValueName As String = "")
'    Dim lRes As Long, lRetLen As Long, lRetVal As Long, lHandle As Long
'    lRes = RegOpenKey(hKey, SubKey, lHandle)
'    If lRes <> ERROR_SUCCESS Then
'        'Error !
'        Err.Raise vbObjectError + regErrOpenKey, "GetDwordKey", "Win 32 API function failed to open registry key '" & SubKey & "'."
'        Exit Function
'    Else
'        'Opened succesfully !
'        lRetLen = 4  'Bytes
'        If RegQueryValueEx(lHandle, ValueName, 0, REG_DWORD, ByVal lRetVal&, lRetLen) = ERROR_SUCCESS Then
'            GetDwordKey = lRetVal
'        Else
'            Err.Raise vbObjectError + regErrQueryKey, "GetDwordKey", "Win32 API function failed to get registry key value: " & SubKey & "'."
'            RegCloseKey lHandle
'            Exit Function
'        End If
'        RegCloseKey lHandle
'    End If
'End Function

Sub SetStringKey(ByVal hKey As APIRegistryRoots, ByVal SubKey As String, Optional ByVal ValueName As String = "", Optional ByVal Setting As String = "")
    Dim hNewHandle&, lpdwDisposition&, Temp As SECURITY_ATTRIBUTES
    If RegCreateKeyEx(hKey, SubKey, 0, ValueName, 0, KEY_ALL_ACCESS, Temp, hNewHandle, lpdwDisposition) = ERROR_SUCCESS Then
        If RegSetValueEx(hNewHandle, ValueName, 0, REG_SZ, ByVal Setting, Len(Setting)) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + regErrSetKey, "SetStringKey", "Win32 API function failed to set key value: '" & SubKey & "'."
            Exit Sub
        End If
    Else
        Err.Raise vbObjectError + regErrCreateKey, "SetStringKey", "Win32 API function failed to create key: '" & SubKey & "'."
        Exit Sub
    End If
    RegCloseKey hNewHandle
End Sub

'Sub SetDWordKey(ByVal hKey As APIRegistryRoots, ByVal SubKey$, Optional ByVal ValueName$ = "", Optional ByVal Setting& = 0)
'    Dim hNewHandle&, lpdwDisposition&, TempSec As SECURITY_ATTRIBUTES
'    If RegCreateKeyEx(hKey, SubKey, 0, ValueName, 0, KEY_ALL_ACCESS, TempSec, hNewHandle, lpdwDisposition) = ERROR_SUCCESS Then
'        If RegSetValueEx(hNewHandle, ValueName, 0, REG_DWORD, Setting, 4) <> ERROR_SUCCESS Then
'            Err.Raise vbObjectError + regErrSetKey, "SetDWordKey", "Win32 API function failed to set registry key value: '" & SubKey & "'."
'            Exit Sub
'        End If
'    Else
'        Err.Raise vbObjectError + regErrCreateKey, "SetDWordKey", "Win32 API function failed to create registry key: '" & SubKey & "'."
'        Exit Sub
'    End If
'    RegCloseKey hNewHandle
'End Sub

'I wanted this one to be here, 'cause some functions need the CPU to be restarted
'Function RestartComputer() As Boolean
'    Dim lFlags As Long
'    'If 1 the not restart, if 2 then restart, if 4 then no ask, close all programs and show explorer
'    'If Restart Then
'        lFlags = 2
'    'Else
'    '    lFlags = 1
'    'End If
'    'If Critical Then lFlags = lFlags + 4
'    RestartComputer = ExitWindowsEx(lFlags, 1)
'End Function

