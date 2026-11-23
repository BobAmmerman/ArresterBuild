VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Porcelain and Polymer Arrester Build"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10275
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Print the Current Form"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Runtime"
            Object.ToolTipText     =   "Goto Runtime Screen"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Setup"
            Object.ToolTipText     =   "Goto Setup Screen"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "AutoCal"
            Object.ToolTipText     =   "Goto Auto Calibration Results"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Comms"
            Object.ToolTipText     =   "Show Comms Screen"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stats"
            Object.ToolTipText     =   "Show Stats Screen"
            Object.Tag             =   ""
            ImageIndex      =   21
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "AlarmReset"
            Object.ToolTipText     =   "Reset any/all active machine alarms"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop Machine at end of cycle or Purge the dial"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BatchName"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   2
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2880
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   5175
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   688
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "frmMain.frx":0442
            Text            =   "User:"
            TextSave        =   "User:"
            Key             =   "Access"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Logged In User"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "WO: "
            TextSave        =   "WO: "
            Key             =   "WO"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Work Order"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "S/N: "
            TextSave        =   "S/N: "
            Key             =   "SN"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Serial Number"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Unit: "
            TextSave        =   "Unit: "
            Key             =   "Unit"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Unit"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1905
            MinWidth        =   882
            TextSave        =   "05/31/2006"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1508
            MinWidth        =   882
            TextSave        =   "7:15 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   2880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   21
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1152
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":17F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1E9A
            Key             =   "Comms"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":24CE
            Key             =   "Setup"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":27E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2E1C
            Key             =   "RunTime"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3136
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3450
            Key             =   "Exclaimation"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":376A
            Key             =   "Lightning"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3A84
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3D9E
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":40B8
            Key             =   "Unsecure"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":43D2
            Key             =   "Secure"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":46EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileDeleteWO 
         Caption         =   "Delete &Work Order"
      End
      Begin VB.Menu mnuFileUnbuild 
         Caption         =   "&Unbuild Unit"
      End
      Begin VB.Menu mnuFilePrintUnit 
         Caption         =   "P&rint Unit Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFilePrintMultiple 
         Caption         =   "Print &Multiple Reports"
      End
      Begin VB.Menu mnuFileLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnuFileLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewError 
         Caption         =   "View &Error Log"
      End
      Begin VB.Menu mnuViewMenuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Arrester Build"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Dim iFormsShown As Integer   'used to indicate which forms should be shown
Public bStartup As Boolean         'Flag to indicate that system is starting up and
                                             'some events should be ignored
                                             

Private Sub MDIForm_Load()
    'Dim frmTemplate As frmTemplate
    Dim iResult As Integer
    Dim iLoop As Integer
    
    Set globFormMain = Me
    bStartup = True

    Me.Left = GetSet(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSet(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSet(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSet(App.Title, "Settings", "MainHeight", 6500)
    
    globFormMain.sbStatusBar.Panels(1).Picture = imlIcons.ListImages("Stop").Picture
    
    globCurrentAccessLevel = VerifyUser
    If globCurrentAccessLevel >= 0 Then
        'good login , proceed normally
        frmTemplate.Show
        frmTemplate.Width = Me.ScaleWidth
        frmTemplate.Height = Me.ScaleHeight
        SetAccess
    Else
        'don't let user do anything if not logged in properly
        MsgBox "Login Failed! Exiting program."
        'hard stop program
        End
    End If
    bStartup = False
    
    MDIForm_Resize
    
End Sub




Private Sub MDIForm_Resize()
    '*******************************************************************************************
    'Created 3/30/00    By:Ray Raf
    'Purpose:   Size child forms to current main form size
    '*******************************************************************************************
    Dim iFormLoop As Integer
    Dim iPossibleTop As Integer, iPossibleHeight As Integer
    Dim iHeightStorage As Integer
    Dim i2ndScreenHeight As Integer
    
    If bStartup = True Then Exit Sub  ' don't size things yet
    
    frmTemplate.Width = Me.ScaleWidth
    frmTemplate.Height = Me.ScaleHeight
    
    'GetFormIndexes
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSet App.Title, "Settings", "MainLeft", Me.Left
        SaveSet App.Title, "Settings", "MainTop", Me.Top
        SaveSet App.Title, "Settings", "MainWidth", Me.Width
        SaveSet App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    'cleanup and close up shop before stopping
    'Shutdown
End Sub





Private Sub mnuFileDeleteWO_Click()
    Dim strWO As String
    
    strWO = InputBox("Enter Work Order Number to delete")
 ' Change by DLimburg on 5/30/06
 ' Change length test to accept a work order with a length of either 7 or 8
    If (Len(Trim(strWO)) < 7) Or (Len(Trim(strWO)) > 8) Then
        MsgBox "The work order must be exactly 7 or 8 digits.", vbInformation, "Error"
        Exit Sub
    End If
    
    If Not IsNumeric(strWO) Then
        MsgBox "The work order only contains digits.", vbInformation, "Error"
        Exit Sub
    End If
    
    DeleteWO Trim(strWO)
    
End Sub

Private Sub mnuFileLogin_Click()
    'user has requested login to system
    Dim iResult As Integer
    
    iResult = VerifyUser
    globCurrentAccessLevel = iResult
    If iResult < 0 Then
        'login was not validated
        MsgBox "User login Failed.  Program is locked until valid user logs in!"
        globCurrentUser = ""
    Else
        'new/existing user logged in sucessfully.
        globLastActivity = Timer
        SetAccess
    End If

End Sub

Private Sub mnuFileLogout_Click()
    'logout current user
    
    globCurrentAccessLevel = -1
    globCurrentUser = ""
    globLastActivity = Timer - conLoginExpireTime - 2   'make sure login time shows as expired
    globFormMain.sbStatusBar.Panels(1).Text = "User:"
    globFormMain.sbStatusBar.Panels(1).Picture = globFormMain.imlIcons.ListImages("Stop").Picture

    MsgBox "Program is locked until valid user logs in!"
    SetAccess

End Sub

Private Sub mnuFilePrintMultiple_Click()
    'this will bring up the form to allow the user to print a unit report.
    'A unit report should only be able to be accessed if frmWorkOrder is displayed
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
    frmPrintMultiple.Show vbModal

End Sub

Private Sub mnuFilePrintUnit_Click()
    'this will bring up the form to allow the user to print a unit report.
    'A unit report should only be able to be accessed if frmWorkOrder is displayed
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
    frmManualReport.Show vbModal
    
End Sub

Private Sub mnuFileUnbuild_Click()
    frmUnbuildSelection.Show vbModal
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show vbModal, Me
End Sub

Private Sub mnuViewError_Click()
     frmViewErrorLog.Show vbModeless
End Sub

Private Sub mnuViewOptions_Click()
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    If globCurrentAccessLevel <> 1 Then
        MsgBox "Your access sevel is not sufficient to access this feature"
    Else
        frmOptions.Show vbModal, Me
    End If
    
End Sub




Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub


