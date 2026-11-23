VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4260
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   4328.71
      ScaleMode       =   0  'User
      ScaleWidth      =   7686.12
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   720
      Width           =   7605
      Begin VB.Frame fraSecurityOptions 
         Caption         =   "Database Location"
         Height          =   2295
         Index           =   3
         Left            =   -19789
         TabIndex        =   48
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtSecurityTableName 
            Height          =   285
            Left            =   1200
            TabIndex        =   51
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CommandButton cmdSecurityBrowseForDB 
            Caption         =   "Browse"
            Height          =   375
            Left            =   3240
            TabIndex        =   50
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtSecurityDBFile 
            Height          =   525
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lblSecurity 
            Caption         =   "Table Name:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblSecurity 
            AutoSize        =   -1  'True
            Caption         =   "Location:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame fraSecurityOptions 
         Height          =   3735
         Index           =   1
         Left            =   -19789
         TabIndex        =   43
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdSecurityDeleteUser 
            Caption         =   "Delete User"
            Height          =   375
            Left            =   2400
            TabIndex        =   45
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CommandButton cmdSecurityAddUser 
            Caption         =   "Add User"
            Height          =   375
            Left            =   360
            TabIndex        =   44
            Top             =   3240
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid grdSecurityUsers 
            Height          =   2535
            Left            =   360
            TabIndex        =   46
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   4471
            _Version        =   393216
            Rows            =   11
            FixedCols       =   0
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
            HighLight       =   2
            ScrollBars      =   2
         End
         Begin VB.Label lblSecurity 
            Alignment       =   2  'Center
            Caption         =   "NOTE: the Administrator user has access by default."
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   47
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Frame fraSecurityOptions 
         Caption         =   "Connection Values"
         Height          =   3255
         Index           =   2
         Left            =   396
         TabIndex        =   29
         Top             =   240
         Width           =   4230
         Begin VB.CommandButton cmdSecurityTestConn 
            Caption         =   "Test Connection"
            Height          =   495
            Left            =   1200
            TabIndex        =   36
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtSecurityServer 
            Height          =   330
            HideSelection   =   0   'False
            Left            =   1125
            TabIndex        =   30
            Top             =   375
            Width           =   3015
         End
         Begin VB.TextBox txtSecurityDatabase 
            Height          =   300
            Left            =   1125
            TabIndex        =   31
            Top             =   780
            Width           =   3015
         End
         Begin VB.TextBox txtSecurityPWD 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1125
            PasswordChar    =   "*"
            TabIndex        =   33
            Top             =   1530
            Width           =   3015
         End
         Begin VB.TextBox txtSecurityUID 
            Height          =   300
            Left            =   1125
            TabIndex        =   32
            Top             =   1155
            Width           =   3015
         End
         Begin VB.TextBox txtSecurityTable 
            Height          =   300
            Left            =   1110
            TabIndex        =   35
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtSecurityRepeatPassword 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1110
            PasswordChar    =   "*"
            TabIndex        =   34
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Server:"
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   42
            Top             =   450
            Width           =   510
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "Data&base:"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   41
            Top             =   825
            Width           =   735
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Password:"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   40
            Top             =   1575
            Width           =   735
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&UID:"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   39
            Top             =   1200
            Width           =   330
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Table:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   2325
            Width           =   450
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Repeat Pass:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   1965
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4155
      Index           =   0
      Left            =   -20000
      ScaleHeight     =   4222.017
      ScaleMode       =   0  'User
      ScaleWidth      =   7670.961
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   825
      Width           =   7590
      Begin VB.Frame fraSecurityOptions 
         Caption         =   "Connection Values"
         Height          =   3780
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   150
         Width           =   7455
         Begin VB.TextBox txtRepeatPassword 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1110
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   2445
            Width           =   3015
         End
         Begin VB.TextBox txtUID 
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   18
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txtPWD 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1125
            PasswordChar    =   "*"
            TabIndex        =   19
            Top             =   2055
            Width           =   3015
         End
         Begin VB.TextBox txtDatabase 
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   17
            Top             =   1305
            Width           =   3015
         End
         Begin VB.TextBox txtServer 
            Height          =   330
            HideSelection   =   0   'False
            Index           =   0
            Left            =   1125
            TabIndex        =   16
            Top             =   900
            Width           =   3015
         End
         Begin VB.CommandButton cmdTestConn 
            Caption         =   "Test Connection"
            Height          =   495
            Index           =   0
            Left            =   1725
            TabIndex        =   15
            Top             =   2925
            Width           =   1935
         End
         Begin VB.Frame fraResultsConn 
            BorderStyle     =   0  'None
            Height          =   3165
            Left            =   4200
            TabIndex        =   7
            Top             =   300
            Width           =   3165
            Begin VB.TextBox txtRepeatPassword 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   0
               PasswordChar    =   "*"
               TabIndex        =   13
               Top             =   2145
               Width           =   3015
            End
            Begin VB.TextBox txtUID 
               Height          =   300
               Index           =   1
               Left            =   15
               TabIndex        =   11
               Top             =   1380
               Width           =   3015
            End
            Begin VB.TextBox txtPWD 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   15
               PasswordChar    =   "*"
               TabIndex        =   12
               Top             =   1755
               Width           =   3015
            End
            Begin VB.TextBox txtDatabase 
               Height          =   300
               Index           =   1
               Left            =   15
               TabIndex        =   10
               Top             =   1005
               Width           =   3015
            End
            Begin VB.TextBox txtServer 
               Height          =   330
               HideSelection   =   0   'False
               Index           =   1
               Left            =   15
               TabIndex        =   9
               Top             =   600
               Width           =   3015
            End
            Begin VB.CommandButton cmdTestConn 
               Caption         =   "Test Connection"
               Height          =   495
               Index           =   1
               Left            =   600
               TabIndex        =   8
               Top             =   2625
               Width           =   1935
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Results"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Index           =   1
               Left            =   450
               TabIndex        =   14
               Top             =   0
               Width           =   1965
            End
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Repeat Pass:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   2490
            Width           =   960
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&UID:"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   25
            Top             =   1725
            Width           =   330
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Password:"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   24
            Top             =   2100
            Width           =   735
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "Data&base:"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   23
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lblStep3 
            AutoSize        =   -1  'True
            Caption         =   "&Server:"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   22
            Top             =   975
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   1650
            TabIndex        =   21
            Top             =   300
            Width           =   1965
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Tag             =   "Cancel"
      Top             =   5415
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Tag             =   "Apply"
      Top             =   5415
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4260
      Index           =   1
      Left            =   420
      ScaleHeight     =   4328.71
      ScaleMode       =   0  'User
      ScaleWidth      =   7686.12
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   7605
      Begin VB.CheckBox chkAllowGorE 
         Caption         =   "Allow E for G disk types"
         Height          =   240
         Left            =   720
         TabIndex        =   54
         Top             =   840
         Width           =   3390
      End
      Begin VB.CheckBox chkEnableAutoPrint 
         Caption         =   "Enable Automatic Printing of Unit Report"
         Height          =   240
         Left            =   720
         TabIndex        =   27
         Top             =   450
         Width           =   3390
      End
   End
   Begin ComctlLib.TabStrip tbsOptions 
      Height          =   4965
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8758
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Databases"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Select database locations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc. "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Security"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bShowResults As Boolean
Dim m_SecurityTab As Integer


Private Sub cmdApply_Click()
    SaveOptions
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim bResult As Boolean
    
    bResult = SaveOptions
    If bResult = True Then
        Unload Me
    End If
    
End Sub

Private Sub cmdSecurityTestConn_Click()
    '----------------------------------------------------------------------------------------------
    '   Created:   03/15/01    By: Ray Raf
    '   Purpose:    Allows the user to check the connection properties
    '                    to a SQL Server based security system.
    '----------------------------------------------------------------------------------------------
    Dim bResult As Boolean
    
    Me.MousePointer = vbHourglass
    bResult = CheckSecurityFrame2
    If bResult = True Then
        Me.MousePointer = vbDefault
        MsgBox "Connection info is good!"
    End If
    
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '--------------------------Default code created by application wizard
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub


Private Sub Form_Load()
    m_SecurityTab = 3
    picOptions(0).Left = tbsOptions.Left + 50
    bShowResults = True
    LoadOptions             'will load saved values of options onto options form
End Sub


Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = tbsOptions.Left + 50
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

Private Function SaveOptions() As Boolean
    '----------------------------------------------------------------------------------------------
    'Created:   02/09/00    By: Ray Raf
    'Purpose:   Save all option values to registry
    'Returns:   True if all values acceptable and seved to registry, False otherwise
    '----------------------------------------------------------------------------------------------
    Dim bResult As Boolean
    
    SaveOptions = False
    'check for valid data before saving anything
    'check impulse current limits 1st
    
    'tab 1
    bResult = CheckDBConnInfo(0)
    If bResult = False Then
        SaveOptions = False
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        tbsOptions_Click
        Exit Function
    End If
    
    If bShowResults = True Then
        'don't try to check results db connection if results db info not visible
        bResult = CheckDBConnInfo(1)
        If bResult = False Then
            SaveOptions = False
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
            tbsOptions_Click
            Exit Function
        End If
    End If
        
    'check security settings
    bResult = CheckSecurityFrame2
    If bResult = False Then
        Exit Function
    End If
    
        
    'if we get this far all values are good, go ahead and save them
    'tab 1 settings
    SaveDBOptions
    'tab 1 public variables
    ConnectDatabases
    
    SaveSet AppName:=App.Title, Section:="Settings", _
                    Key:="AutoPrint", Setting:=chkEnableAutoPrint.Value
    globAutoPrint = chkEnableAutoPrint.Value
    
    SaveSet AppName:=App.Title, Section:="Settings", _
                    Key:="GandESubs", Setting:=chkAllowGorE.Value
    globGandESubs = chkAllowGorE.Value
    
    SaveSecurityFrame2
    
    'return that everything was good
    SaveOptions = True
    
End Function

Private Sub LoadOptions()
    '----------------------------------------------------------------------------------------------
    'Created:   02/09/00    By: Ray Raf
    'Purpose:   Load current settings into proper controls on this form
    '----------------------------------------------------------------------------------------------
    
    'tab 1 options
    LoadDBOptions
    
    'tab 2 options
    chkEnableAutoPrint.Value = GetSet(AppName:=App.Title, Section:="Settings", _
                                                        Key:="AutoPrint", Default:="0")
    chkAllowGorE.Value = GetSet(AppName:=App.Title, Section:="Settings", _
                                                        Key:="GandESubs", Default:="0")
                                                        
    'tab 3 options
    LoadSecurityFrame2
    
    
End Sub


Private Sub LoadDBOptions()
    '----------------------------------------------------------------------------------------------
    'Created:   07/19/02    By: Ray Raf
    'Purpose:   Load Database Settings into controls on this form
    '----------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    Dim strDBName As String
    
    'this routine assumes that all SQL Server database connections to be
    'saved have their data stored in an array of textboxes with all of the
    'data for a connection stored in boxes with the same index
    'ie txtDatabase(0), txtServer(0), etc.
    For iLoop = 0 To 1
        'load a name to prefix the stored settings
        Select Case iLoop
            Case 0
                strDBName = "SetupReadOnly"
            Case 1
                strDBName = "Results"
        End Select
        
        If bShowResults = True Or iLoop <> 1 Then
            'don't try to load results db connection if results db info not visible
            'load data for each connection
            txtServer(iLoop).Text = GetSet(AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_Server", Default:="")
            txtDatabase(iLoop).Text = GetSet(AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_Database", Default:="")
            txtUID(iLoop).Text = GetSet(AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_UID", Default:="")
            txtPWD(iLoop).Text = GetSet(AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_PWD", Default:="")
            txtRepeatPassword(iLoop).Text = txtPWD(iLoop).Text
        End If
    Next iLoop

End Sub


Private Sub SaveDBOptions()
    '----------------------------------------------------------------------------------------------
    'Created:   07/19/02    By: Ray Raf
    'Purpose:   Save Database Settings into registry/ini file
    '----------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    Dim strDBName As String
    
    'this routine assumes that all SQL Server database connections to be
    'saved have their data stored in an array of textboxes with all of the
    'data for a connection stored in boxes with the same index
    'ie txtDatabase(0), txtServer(0), etc.
    For iLoop = 0 To 1
        'load a name to prefix the stored settings
        Select Case iLoop
            Case 0
                strDBName = "SetupReadOnly"
            Case 1
                strDBName = "Results"
        End Select
        
        If bShowResults = True Or iLoop <> 1 Then
            'don't try to save results db connection if results db info not visible
            'load data for each connection
            SaveSet AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_Server", Setting:=txtServer(iLoop).Text
            SaveSet AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_Database", Setting:=txtDatabase(iLoop).Text
            SaveSet AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_UID", Setting:=txtUID(iLoop).Text
            SaveSet AppName:=App.Title, Section:="DB", _
                Key:=strDBName & "_PWD", Setting:=txtPWD(iLoop).Text
        End If
        
    Next iLoop

End Sub

Private Sub cmdTestConn_Click(Index As Integer)
    '----------------------------------------------------------------------------------------------
    '   Created:   03/15/01    By: Ray Raf
    '   Purpose:    Allows the user to check the connection properties
    '                    to a SQL Server based security system.
    '----------------------------------------------------------------------------------------------
    Dim bResult As Boolean
    
    Me.MousePointer = vbHourglass
    bResult = CheckDBConnInfo(Index)
    If bResult = True Then
        Me.MousePointer = vbDefault
        MsgBox "Connection info is good!"
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Function CheckDBConnInfo(Index As Integer) As Boolean
    '----------------------------------------------------------------------------------------------
    'Created:   03/15/01    By: Ray Raf
    'Purpose:   checks the entries in the SQL Server selection frame.
    '----------------------------------------------------------------------------------------------
    Dim dbcon As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim ConnectionString As String
    Dim strTableName As String
    
    If Trim(txtServer(Index).Text) = "" Then
        MsgBox "You must enter the server name which contains the " & vbCrLf & _
                    "database where the info resides."
        'make value in question active
        txtServer(Index).SetFocus
        CheckDBConnInfo = False
        Exit Function
    End If
    
    If Trim(txtDatabase(Index).Text) = "" Then
        MsgBox "You must enter the database name which contains the " & vbCrLf & _
                    "table where the info resides."
        'make value in question active
        txtDatabase(Index).SetFocus
        CheckDBConnInfo = False
        Exit Function
    End If
    
    If Trim(txtUID(Index).Text) = "" Then
        MsgBox "You must enter the user ID to use to gain access" & vbCrLf & _
                    "to the database where the info resides."
        'make value in question active
        txtUID(Index).SetFocus
        CheckDBConnInfo = False
        Exit Function
    End If
    
    If txtPWD(Index).Text <> txtRepeatPassword(Index).Text Then
        'passwords don't match
        MsgBox "The passwords don't match!"
        'make value in question active
        txtPWD(Index).SetFocus
        CheckDBConnInfo = False
        Exit Function
    End If
    
    'check connection to SQL Database
    ConnectionString = "driver={SQL Server};" & _
        "server=" & txtServer(Index).Text & ";" & _
        "uid=" & txtUID(Index).Text & ";" & _
        "pwd=" & txtPWD(Index).Text & ";" & _
        "database=" & txtDatabase(Index).Text & ";"
    
    'make sure connection isn't already open
    If dbcon.State = 1 Then
        dbcon.Close
    End If
    
    'open the connection to the database
    With dbcon
        .ConnectionString = ConnectionString
        On Error Resume Next
        .Open
        If Err.Number <> 0 Then
            'had a problem making a connection.
            MsgBox "Could not make a connection to the SQL Server database specified. " & vbCrLf & _
                        "Returned message was: " & Err.Description
            Debug.Print ConnectionString
            CheckDBConnInfo = False
            Err.Clear
            Exit Function
        End If
    End With
    
    'pick a table that is appropriate for the function being performed.
    Select Case Index
        Case 0
            strTableName = "MCOV_Porc_Rev1"   ' Changed table name from MCOV_Porc to MCOV_Porc_Rev1 DLimburg 21FEb2011
        Case 1
            strTableName = "WO_Info"
    End Select
    
    'check to make sure table can be opened
    rst.Open strTableName, dbcon, adOpenDynamic, adLockReadOnly
    If Err.Number <> 0 Then
        MsgBox "Could not open table specified." & vbCrLf & _
                        "Returned message was: " & Err.Description
        CheckDBConnInfo = False
        Err.Clear
        Exit Function
    End If
    
    CheckDBConnInfo = True
    On Error GoTo 0
    rst.Close
    dbcon.Close
    

End Function

Private Sub LoadSecurityFrame2()
    '----------------------------------------------------------------------------------------------
    '   Created:   01/31/01    By: Ray Raf
    '   Purpose:    Loads the option values for fraSecurityOptions(2)
    '                   (Using SQL Server Database)
    '----------------------------------------------------------------------------------------------
            
    txtSecurityUID.Text = GetSet(App.Title, "Security", "UID", " ")
    txtSecurityPWD.Text = GetSet(App.Title, "Security", "Password", " ")
    txtSecurityRepeatPassword.Text = txtSecurityPWD.Text
    txtSecurityDatabase.Text = GetSet(App.Title, "Security", "Database", " ")
    txtSecurityServer.Text = GetSet(App.Title, "Security", "Server", " ")
    txtSecurityTable.Text = GetSet(App.Title, "Security", "Table", " ")
    
End Sub

Private Sub SaveSecurityFrame2()
    '----------------------------------------------------------------------------------------------
    '   Created:   02/02/01    By: Ray Raf
    '   Purpose:    Saves the option values for fraSecurityOptions(2)
    '                   (Using SQL Server Database)
    '----------------------------------------------------------------------------------------------
    
    SaveSet AppName:=App.Title, Section:="Security", _
                    Key:="UID", Setting:=txtSecurityUID.Text
    SaveSet AppName:=App.Title, Section:="Security", _
                    Key:="Password", Setting:=txtSecurityPWD.Text
    SaveSet AppName:=App.Title, Section:="Security", _
                    Key:="Database", Setting:=txtSecurityDatabase.Text
    SaveSet AppName:=App.Title, Section:="Security", _
                    Key:="Server", Setting:=txtSecurityServer.Text
    SaveSet AppName:=App.Title, Section:="Security", _
                    Key:="Table", Setting:=txtSecurityTable.Text
     
End Sub

Private Function CheckSecurityFrame2() As Boolean
    '----------------------------------------------------------------------------------------------
    'Created:   03/15/01    By: Ray Raf
    'Purpose:   checks the entries in the SQL Server security table
    '                selection frame.
    '----------------------------------------------------------------------------------------------
    Dim dbcon As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim ConnectionString As String
    
    If Trim(txtSecurityServer.Text) = "" Then
        MsgBox "You must enter the server name which contains the " & vbCrLf & _
                    "database where the security info resides."
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityServer.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
    
    If Trim(txtSecurityDatabase.Text) = "" Then
        MsgBox "You must enter the database name which contains the " & vbCrLf & _
                    "table where the security info resides."
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityDatabase.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
    
    If Trim(txtSecurityUID.Text) = "" Then
        MsgBox "You must enter the user ID to use to gain access" & vbCrLf & _
                    "to the database where the security info resides."
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityUID.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
    
    If Trim(txtSecurityPWD.Text) = "" Then
        MsgBox "You must enter the password for the given User ID" & vbCrLf & _
                    "in the database where the security info resides."
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityPWD.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
    
    If txtSecurityPWD.Text <> txtSecurityRepeatPassword.Text Then
        'passwords don't match
        MsgBox "The passwords don't match!"
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityPWD.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
        
    
    If Trim(txtSecurityTable.Text) = "" Then
        MsgBox "You must enter the table name which " & vbCrLf & _
                    "contains the security info."
        'make tab with security stuff active
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        tbsOptions_Click
        'make value in question active
        txtSecurityTable.SetFocus
        CheckSecurityFrame2 = False
        Exit Function
    End If
    
    
    'check connection to SQL Database
    ConnectionString = "driver={SQL Server};" & _
        "server=" & txtSecurityServer.Text & ";" & _
        "uid=" & txtSecurityUID.Text & ";" & _
        "pwd=" & txtSecurityPWD.Text & ";" & _
        "database=" & txtSecurityDatabase.Text & ";"
    
    'make sure connection isn't already open
    If dbcon.State = 1 Then
        dbcon.Close
    End If
    
    'open the connection to the database
    With dbcon
        .ConnectionString = ConnectionString
        On Error Resume Next
        .Open
        If Err.Number <> 0 Then
            'had a problem making a connection.
            MsgBox "Could not make a connection to the SQL Server database specified. " & vbCrLf & _
                        "Returned message was: " & Err.Description
            CheckSecurityFrame2 = False
            Err.Clear
            Exit Function
        End If
    End With
    
    'check to make sure table can be opened
    rst.Open txtSecurityTable.Text, dbcon, adOpenDynamic, adLockReadOnly
    If Err.Number <> 0 Then
        MsgBox "Could not open table specified." & vbCrLf & _
                        "Returned message was: " & Err.Description
        CheckSecurityFrame2 = False
        Err.Clear
        Exit Function
    End If
    
    CheckSecurityFrame2 = True
    On Error GoTo 0
    rst.Close
    dbcon.Close
    
End Function


