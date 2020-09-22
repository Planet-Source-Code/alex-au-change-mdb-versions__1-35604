VERSION 5.00
Begin VB.Form frmChangeMDBVersion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change MDB Version"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmChangeMDBVersion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin ChangeMDBVersion.ButtonEx btnexCmd 
      Default         =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   2625
      TabIndex        =   6
      Top             =   945
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   953
      BorderStyle     =   4
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "Access 2000 &To Access 97"
      Height          =   250
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   1260
      Width           =   2500
   End
   Begin VB.OptionButton optVersion 
      Caption         =   "&Access 97 To Access 2000"
      Height          =   250
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   945
      Value           =   -1  'True
      Width           =   2500
   End
   Begin ChangeMDBVersion.FolderBrowser fldrDest 
      Height          =   330
      Left            =   1995
      TabIndex        =   3
      Top             =   420
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChangeMDBVersion.FolderBrowser fldrSource 
      Height          =   330
      Left            =   1995
      TabIndex        =   1
      Top             =   105
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChangeMDBVersion.ButtonEx btnexCmd 
      Cancel          =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   3780
      TabIndex        =   7
      Top             =   945
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   953
      BorderStyle     =   4
      Caption         =   "E&xit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   1575
      Width           =   4950
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Destination Data Path:"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   525
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Source Data Path:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1800
   End
   Begin VB.Label lblProgress 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   1575
      Width           =   4950
   End
End
Attribute VB_Name = "frmChangeMDBVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Windows 98 and 2000 DAO registry Key
'HKEY_LOCAL_MACHINE\Software\CLASSES\DAO.DBEngine.35
'HKEY_LOCAL_MACHINE\Software\CLASSES\DAO.DBEngine.36

Dim objDAO As Variant           'Late bound Data Object
Dim bInIDE As Boolean           ' Check whether in IDE mode

Private Sub btnexCmd_Click(Index As Integer)
    Select Case Index
        Case 0              ' OK
            PerformMigration
        Case 1              ' Exit
            Unload Me
        Case Else
            MsgBox "Missing btnexCmd_Click(" & Trim(Str(Index)) & ")!"
    End Select
End Sub

Private Sub fldrDest_GotFocus()
    Me.lblStatus.Caption = "Ready"
    fldrDest.SelStart = 0
    fldrDest.SelLength = Len(fldrDest.Text)
End Sub

Private Sub fldrSource_Change()
    fldrDest.Text = fldrSource.Text & "\NewVer"
End Sub

Private Sub fldrSource_GotFocus()
    Me.lblStatus.Caption = "Ready"
    fldrSource.SelStart = 0
    fldrSource.SelLength = Len(fldrSource.Text)
End Sub

Private Sub Form_Click()
    Me.lblStatus.Caption = "Ready"
End Sub

Private Sub Form_Load()
    bInIDE = InIDE()
    
    Me.Caption = Me.Caption & " (Build " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    
    If Not bInIDE Then On Error Resume Next
    
    Dim sDAO As String
    sDAO = GetStringKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\DAO.DBEngine.36", "")
    If Len(Trim(sDAO)) <= 0 Then GoTo ErrLoaded
    
    Set objDAO = CreateObject("DAO.DBEngine.36")
    If UCase(Trim(TypeName(objDAO))) <> "DBENGINE" Then GoTo ErrLoaded

    fldrSource.Text = App.Path
    fldrDest.Text = App.Path & "\NewVer"

    With lblStatus
        .Left = 0
        .Width = Me.ScaleWidth
    End With

    With lblProgress
        .Visible = False
        .Left = 0
        .Width = 0
    End With
    
    Exit Sub

ErrLoaded:
    btnexCmd(0).Enabled = False
    MsgBox "No suitable DAO database engine can be loaded. This tool requires DAO 3.6.", vbCritical + vbOKOnly, "Change MDB Version Error"
    
End Sub

Private Sub PerformMigration()
    If Not bInIDE Then On Error GoTo Err_Handler

    Dim m As New cMouse
    m.SetPointer vbHourglass          ' Or any other pointer shape
    
    Dim sMdb As String                              'Database to be used
    Dim sSource() As String                         'Source Database Array
    Dim nCount As Long                              'Count number of databases
    Dim sSourcePath As String
    Dim sDestinationPath As String
    Dim sCurMDB As String
    Dim sRet As Long
    Dim sAppPath As String
    Dim sTemp As String
    Dim i As Long
    
    sAppPath = Trim(App.Path)
    sSourcePath = Trim(fldrSource.Text)
    sDestinationPath = Trim(fldrDest.Text)
    
    If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    If Right(sSourcePath, 1) <> "\" Then sSourcePath = sSourcePath & "\"
    If Right(sDestinationPath, 1) <> "\" Then sDestinationPath = sDestinationPath & "\"
    
    If optVersion(0).Value Then                     ' 97 to 2000
        sMdb = "2K"
    ElseIf optVersion(1).Value Then                 ' 2000 to 97
        sMdb = "97"
    End If
    
    sTemp = Dir(sSourcePath & "*.MDB")
    
    If Len(Trim(sTemp)) < 2 Then
        MsgBox "No MDB files found in " & sSourcePath, vbExclamation + vbOKOnly, "Change MDB Version."
        Exit Sub
    End If
    
    nCount = 1
    ReDim sSource(1 To nCount) As String
    sSource(nCount) = sTemp
    
    While Len(Trim(sTemp)) > 2
        sTemp = Dir
        If Len(Trim(sTemp)) > 2 Then
            nCount = nCount + 1
            ReDim Preserve sSource(1 To nCount) As String
            sSource(nCount) = sTemp
        End If
    Wend
    
    If Len(Trim(Dir(sDestinationPath & "NUL"))) < 2 Then CreateDirectoryStruct sDestinationPath
    
    If Len(Trim(Dir(sDestinationPath & "*.MDB"))) > 2 Then
        sRet = MsgBox("'" & sDestinationPath & "' contains some Access database files. " & _
               "Same file name in the path will be overwritten. " & _
               "Do you want to continue?", vbYesNo + vbCritical, "Change MDB Version")
        Select Case sRet
            Case vbYes
            Case vbNo
                Exit Sub
        End Select
    End If
    
    With lblProgress
        .Left = 0
        .Width = 0
        .Visible = True
    End With

    For i = 1 To nCount
        If Len(Trim(Dir(sDestinationPath & sSource(i)))) > 2 Then Kill sDestinationPath & sSource(i)
        Me.lblStatus.Caption = "(" & Trim(Str(i)) & "/" & Trim(Str(nCount)) & ") Creating " & sSource(i)
        Me.lblStatus.Refresh
        CreateDataJetDatabase sDestinationPath & sSource(i), sMdb
        Me.lblStatus.Caption = "(" & Trim(Str(i)) & "/" & Trim(Str(nCount)) & ") Building Structure for " & sSource(i)
        Me.lblStatus.Refresh
        MigrateDB sSourcePath & sSource(i), sDestinationPath & sSource(i)
        Me.lblProgress.Width = i / nCount * lblStatus.Width
        Me.lblProgress.Refresh
    Next i

    Me.lblStatus.Caption = "Version Changed!"
    With lblProgress
        .Visible = False
        .Left = 0
        .Width = 0
    End With

Exit_Sub:
    Exit Sub
    
Err_Handler:
    MsgBox "(" & Str(Err.Number) & ") " & Err.Description
End Sub

Private Sub CreateDataJetDatabase(ByVal sDestinationDB As String, ByVal sVer As String)
    Dim db As DAO.Database
    Select Case sVer
        Case "97"
            Set db = Workspaces(0).CreateDatabase(sDestinationDB, dbLangGeneral, dbVersion30)
        Case "2K"
            Set db = Workspaces(0).CreateDatabase(sDestinationDB, dbLangGeneral, dbVersion40)
    End Select
    db.Close
    Set db = Nothing
End Sub

Private Sub MigrateDB(ByVal sSourceDB As String, ByVal sDestinationDB As String)
    Dim dbSource As DAO.Database
    Dim dbDest As DAO.Database
    
    Dim tdfLoop As DAO.TableDef
    
    Dim qryLoop As DAO.QueryDef
    Dim qryNew As DAO.QueryDef
    
    Dim idxFrom As Index
    Dim idxTo As Index
    
    Dim fldFrom As Field
    Dim tdfDest As DAO.TableDef
    Dim rsSource As DAO.Recordset
    Dim rsDest As DAO.Recordset
    Dim sTemp As String
    
    Dim sConnect As String

    Set dbSource = Workspaces(0).OpenDatabase(sSourceDB)
    Set dbDest = Workspaces(0).OpenDatabase(sDestinationDB)
    
    With dbSource
        For Each tdfLoop In .TableDefs
            If UCase(Left(Trim(tdfLoop.Name), 4)) <> "MSYS" Then
                sConnect = "SELECT * INTO [;database=" & sDestinationDB & "].[" & tdfLoop.Name & "] from [" & tdfLoop.Name & "]"
                dbSource.Execute sConnect
                dbDest.TableDefs.Refresh
                For Each idxFrom In dbSource.TableDefs(tdfLoop.Name).Indexes
                    Set idxTo = dbDest.TableDefs(tdfLoop.Name).CreateIndex(idxFrom.Name)
                    idxTo.Fields = idxFrom.Fields
                    idxTo.Unique = idxFrom.Unique
                    idxTo.Primary = idxFrom.Primary
                    dbDest.TableDefs(tdfLoop.Name).Indexes.Append idxTo
                Next idxFrom
                
                Set rsSource = dbSource.OpenRecordset(tdfLoop.Name, dbOpenDynaset)
                Set tdfDest = dbDest.TableDefs(tdfLoop.Name)
                For Each fldFrom In rsSource.Fields
                    If Not (IsNull(fldFrom.DefaultValue) Or Len(Trim(fldFrom.DefaultValue)) < 1) Then
                        tdfDest.Fields(fldFrom.Name).DefaultValue = fldFrom.DefaultValue
                    End If
                Next fldFrom
            End If
        Next tdfLoop
        For Each qryLoop In .QueryDefs
            Set qryNew = dbDest.CreateQueryDef(qryLoop.Name, qryLoop.SQL)
        Next qryLoop
    End With

End Sub

Private Sub CreateDirectoryStruct(CreateThisPath As String)
    'do initial check
    Dim ret As Boolean, Temp$, ComputerName As String, IntoItCount As Integer, X%, WakeString As String
    Dim MadeIt As Integer
    If Dir$(CreateThisPath, vbDirectory) <> "" Then Exit Sub
    'is this a network path?


    If Left$(CreateThisPath, 2) = "\\" Then ' this is a UNC NetworkPath
        'must extract the machine name first, th
        '     en get to the first folder
        IntoItCount = 3
        ComputerName = Mid$(CreateThisPath, IntoItCount, InStr(IntoItCount, CreateThisPath, "\") - IntoItCount)
        IntoItCount = IntoItCount + Len(ComputerName) + 1
        IntoItCount = InStr(IntoItCount, CreateThisPath, "\") + 1
        'temp = Mid$(CreateThisPath, IntoItCount
        '     , x)
    Else ' this is a regular path
        IntoItCount = 4
    End If
    WakeString = Left$(CreateThisPath, IntoItCount - 1)
    'start a loop through the CreateThisPath
    '     string


    Do
        X = InStr(IntoItCount, CreateThisPath, "\")


        If X <> 0 Then
            X = X - IntoItCount
            Temp = Mid$(CreateThisPath, IntoItCount, X)
        Else
            Temp = Mid$(CreateThisPath, IntoItCount)
        End If
        IntoItCount = IntoItCount + Len(Temp) + 1
        Temp = WakeString + Temp
        'Create a directory if it doesn't alread
        '     y exist
        ret = (Dir$(Temp, vbDirectory) <> "")


        If Not ret Then
            'ret& = CreateDirectory(temp, Security)
            MkDir Temp
        End If
        IntoItCount = IntoItCount 'track where we are in the String
        WakeString = Left$(CreateThisPath, IntoItCount - 1)
    Loop While WakeString <> CreateThisPath
End Sub

Private Sub optVersion_Click(Index As Integer)
    Me.lblStatus.Caption = "Ready"
End Sub
