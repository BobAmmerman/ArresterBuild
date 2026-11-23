VERSION 5.00
Begin VB.Form frmWorkOrder 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4275
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "&Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5190
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1050
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6300
      TabIndex        =   1
      Top             =   4260
      Width           =   1050
   End
   Begin VB.CommandButton btnPreviousWO 
      Caption         =   "Use &Previous Work Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3375
      TabIndex        =   0
      Top             =   615
      Width           =   1785
   End
   Begin VB.TextBox txtWorkOrder 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1380
      MaxLength       =   12
      TabIndex        =   3
      Top             =   195
      Width           =   1785
   End
   Begin VB.Label lblPrevWorkOrder 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   750
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Previous Work Order:"
      Height          =   390
      Left            =   225
      TabIndex        =   5
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Work Order:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   225
      TabIndex        =   2
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "frmWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnNext_Click()

    Dim rstWO As New ADODB.Recordset
    Dim SQL As String
    Dim iSNChecked As Integer
    Dim existingSNs As Boolean
    Dim strLastWO As String
    Dim bResult As Boolean
    
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
 ' Change by DLimburg on 5/30/06
 ' Change length test to accept a work order with a length of either 7 or 8
    If (Len(Trim(txtWorkOrder.Text)) < 7) Or (Len(Trim(txtWorkOrder.Text)) > 8) Then
        MsgBox "A work order must be exactly 7 or 8 characters", vbInformation, "Error"
        txtWorkOrder.SetFocus
        Exit Sub
    End If
    
    If adoWorkOrderInDatabase(txtWorkOrder.Text) = True Then  'also sets cat number, buildpath and type
        'find out if serial numbers were checked.  This only has to be done the first time the work order is access
        SQL = "SELECT * FROM WO_Info WHERE WO_Number = '" & globCurrentWO.strWO & "' "
        rstWO.Open SQL, dbconResultsData, adOpenDynamic, adLockPessimistic
        iSNChecked = rstWO!SerialNumbersChecked
        If iSNChecked = 0 Then
            '--procedure 1.6, verify catalog number
            If VerifyCatalogNumber(globCurrentWO.strCatNumber) = False Then
                Set rstWO = Nothing     'close recordset
                Exit Sub
            End If
            
            'good catalog number, check to see if all serial numbers available
            existingSNs = adoCheckForExistingSNsByWO(globCurrentWO.strWO)
            If existingSNs = False Then
                'no built units found for any serial numbers to be built in this work order.
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastWO_Used", Setting:=txtWorkOrder.Text
                'mark the fact that the serial numbers were checked.
                SQL = "SELECT * FROM WO_Info WHERE WO_Number = '" & globCurrentWO.strWO & "' "
                If rstWO.State > 0 Then
                    rstWO.Close
                End If
                rstWO.Open SQL, dbconResultsData, adOpenDynamic, adLockPessimistic
                rstWO!SerialNumbersChecked = 1
                rstWO.Update
                rstWO.Close
                GoTo nextForm
            Else
                MsgBox "Conflict, notifying user with list (option to print), moving data between tables."
                Exit Sub
            End If
        Else
            'don't need to check serial numbers or catalog number
            'see if this WO is the same as the last one used
            strLastWO = GetSet(AppName:=App.Title, Section:="Memory", _
                    Key:="LastWO_Used", Default:="")
            If Trim(strLastWO) <> Trim(txtWorkOrder.Text) Then
                'Set it so user has no default for num disks to use
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastNumDisksToUse", Setting:="0"
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastUnit", Setting:="0"
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastWO_Used", Setting:=txtWorkOrder.Text
                'Set it so user has no default for num 3kV disks to use
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastNum3kVDisksToUse", Setting:="0"
                'Set it so user has no default for num 5kV disks to use
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastNum5kVDisksToUse", Setting:="0"
            End If
            GoTo nextForm
        End If
    Else
        'WO doesn't exist, allow user to add it.
        globButtonResult = False
        ' add 14Apr2011 DLImburg Porc calcined disk add
        frmAddWO.picCalcineDiskBuild.Visible = False
        frmAddWO.opCalcinedNo = False
        frmAddWO.opCalcinedYes = False
        frmAddWO.Show vbModal
        If globButtonResult <> vbOK Then
            'work order not created
            Exit Sub
        Else
            'Work order created OK
            GoTo nextForm
        End If
    End If
    
    
    MsgBox "At end of procedure."
    Exit Sub
nextForm:
    'unit infomation common in WO
    ParseCatNumToUnit globCurrentWO.strCatNumber, globCurrentUnit
    
    'show the work order number in the status bar
    globFormMain.sbStatusBar.Panels("WO").Text = "WO: " & globCurrentWO.strWO
    Load frmWOOptions
    frmWOOptions.Show
    frmWOOptions.ZOrder 0
    resizeForm
    Unload Me
End Sub

Private Sub btnPreviousWO_Click()
    Dim bResult As Boolean
    
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    ' if the work order number has already been entered, then have
    ' the Enter key call the Next button click event
    If txtWorkOrder.Text = lblPrevWorkOrder.Caption Then
        Call btnNext_Click
        Exit Sub
    End If
    
    txtWorkOrder.Text = lblPrevWorkOrder.Caption
End Sub

Private Sub Form_Load()
    Call setDefaultLocation(frmWorkOrder)
    Call setCaption("What is the work order?")
    Call setProgress(0)
    
    lblPrevWorkOrder.Caption = GetSet(AppName:=App.Title, Section:="Memory", _
            Key:="LastWO_Used", Default:="")
    With globFormMain.sbStatusBar
        .Panels("WO").Text = "WO:"
        .Panels("SN") = "S/N:"
        .Panels("Unit") = "Unit:"
    End With
    
    'allow user to select manual report printing
    globFormMain.mnuFilePrintUnit.Enabled = True
    globFormMain.mnuFileLogout.Enabled = True
    SetAccess
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'make sure that when this form is unloaded the user cannot print a unit report
    globFormMain.mnuFilePrintUnit.Enabled = False
    globFormMain.mnuFileUnbuild.Enabled = False
    globFormMain.mnuFileLogout.Enabled = False
    globFormMain.mnuFileDeleteWO.Enabled = False
    globFormMain.mnuFilePrintMultiple.Enabled = False
    globFormMain.mnuViewOptions.Enabled = False
    
End Sub
