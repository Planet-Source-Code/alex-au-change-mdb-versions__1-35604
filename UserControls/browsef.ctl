VERSION 5.00
Begin VB.UserControl FolderBrowser 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   KeyPreview      =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   2535
   Begin VB.TextBox txtFolder 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   60
      Width           =   1680
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   270
      Left            =   1830
      Picture         =   "browsef.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   270
   End
End
Attribute VB_Name = "FolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'Event Declarations:
Event Change() 'MappingInfo=txtFolder,txtFolder,-1,Change

Private Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    'Set the properties of the folder dialog
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    'Show the Browse For Folder dialog
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
      'Trim off the null chars ending the path
      'and display the returned folder
      Offset = InStr(RetPath, Chr$(0))
      GetFolder = Left$(RetPath, Offset - 1)
    End If

End Function

Private Sub cmdBrowse_Click()
    Dim sFolder As String
    
    txtFolder.SetFocus 'so we dont see the Focus Rectangle
    'Show the browse for folder dialog
    sFolder = GetFolder(UserControl.hWnd, "Select a folder")
    If sFolder <> "" Then
        txtFolder.Text = sFolder
    End If
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the Browse dialog to show
    'Alt and down arrow.
    If KeyCode = vbKeyDown And Shift = 4 Then
        cmdBrowse_Click
    End If
End Sub

Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Line (cmdBrowse.Left - Screen.TwipsPerPixelX, 0)-(cmdBrowse.Left - Screen.TwipsPerPixelX, UserControl.ScaleHeight), vb3DLight, BF
End Sub

Private Sub UserControl_Resize()
    'Position the constituent controls
    cmdBrowse.Move UserControl.ScaleWidth - cmdBrowse.Width, 0, cmdBrowse.Width, UserControl.ScaleHeight
    txtFolder.Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY, UserControl.ScaleWidth - (cmdBrowse.Width + (3 * Screen.TwipsPerPixelX)), UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    'Get the tooltip
    txtFolder.ToolTipText = UserControl.Extender.ToolTipText
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtFolder.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtFolder.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtFolder.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtFolder.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtFolder.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtFolder.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtFolder.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtFolder.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtFolder.Refresh
End Sub

Private Sub txtFolder_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtFolder.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtFolder.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtFolder.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtFolder.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtFolder.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtFolder.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtFolder.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtFolder.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtFolder,txtFolder,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtFolder.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtFolder.Text() = New_Text
    PropertyChanged "Text"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtFolder.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtFolder.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtFolder.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtFolder.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtFolder.Locked = PropBag.ReadProperty("Locked", False)
    txtFolder.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtFolder.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtFolder.SelText = PropBag.ReadProperty("SelText", "")
    txtFolder.Text = PropBag.ReadProperty("Text", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtFolder.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtFolder.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtFolder.Enabled, True)
    Call PropBag.WriteProperty("Font", txtFolder.Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", txtFolder.Locked, False)
    Call PropBag.WriteProperty("SelLength", txtFolder.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtFolder.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtFolder.SelText, "")
    Call PropBag.WriteProperty("Text", txtFolder.Text, "")
End Sub

