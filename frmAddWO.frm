VERSION 5.00
Begin VB.Form frmAddWO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Work Order"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCalcineDiskBuild 
      Height          =   375
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   3315
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton opCalcinedNo 
         Caption         =   "No"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   45
         Width           =   735
      End
      Begin VB.OptionButton opCalcinedYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   45
         Width           =   735
      End
      Begin VB.Label lbPorcCalcineQuestion 
         Caption         =   "Are calcined disks used?"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   45
         Width           =   1935
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2355
      TabIndex        =   12
      Top             =   5175
      Width           =   1545
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3930
      TabIndex        =   13
      Top             =   5175
      Width           =   1545
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate"
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
      Height          =   345
      Left            =   6075
      TabIndex        =   9
      Top             =   1920
      Width           =   1005
   End
   Begin VB.TextBox txtAutoAmount 
      Enabled         =   0   'False
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
      Left            =   5265
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1920
      Width           =   720
   End
   Begin VB.TextBox txtAutoSN 
      Enabled         =   0   'False
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
      Left            =   5265
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1515
      Width           =   1830
   End
   Begin VB.CommandButton btnManualAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2700
      TabIndex        =   4
      Top             =   1515
      Width           =   600
   End
   Begin VB.TextBox txtManualSN 
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
      Left            =   765
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1515
      Width           =   1830
   End
   Begin VB.CommandButton btnAutoClear 
      Caption         =   "Clear"
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
      Height          =   315
      Left            =   6120
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton btnManualClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   255
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.ListBox lstAutoSN 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmAddWO.frx":0000
      Left            =   4365
      List            =   "frmAddWO.frx":0002
      TabIndex        =   10
      Top             =   2310
      Width           =   2730
   End
   Begin VB.ListBox lstManualSN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "frmAddWO.frx":0004
      Left            =   240
      List            =   "frmAddWO.frx":0006
      TabIndex        =   5
      Top             =   1890
      Width           =   3060
   End
   Begin VB.OptionButton chkSN 
      Caption         =   "Automatic serial numbers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4260
      TabIndex        =   2
      Top             =   1125
      Width           =   3255
   End
   Begin VB.OptionButton chkSN 
      Caption         =   "Manually enter serial numbers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   1125
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.TextBox txtCatNum 
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
      Left            =   1860
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2130
   End
   Begin VB.TextBox txtWO 
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
      Left            =   1860
      MaxLength       =   12
      TabIndex        =   20
      Top             =   165
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "How Many:"
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
      Index           =   5
      Left            =   4320
      TabIndex        =   19
      Top             =   1965
      Width           =   885
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "First SN:"
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
      Index           =   4
      Left            =   4305
      TabIndex        =   18
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "SN:"
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
      Index           =   3
      Left            =   270
      TabIndex        =   17
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "- or -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3615
      TabIndex        =   16
      Top             =   1185
      Width           =   390
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   " Catalog Number:"
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
      Index           =   1
      Left            =   210
      TabIndex        =   15
      Top             =   645
      Width           =   1455
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
      Index           =   0
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   1020
   End
End
Attribute VB_Name = "frmAddWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAutoClear_Click()
    lstAutoSN.Clear
End Sub

Private Sub btnCancel_Click()
    'let calling routine know that WO creation was cancelled
    globButtonResult = vbCancel
    Unload Me
End Sub

Private Sub btnGenerate_Click()
    Dim iLoop  As Integer
    Dim lngUpperHalf As Long
    Dim lngLowerHalf As Long
    Dim strSN As String
    Dim lngNewLower As Long
    Dim lngNewUpper As Long
    Dim strHoldLeftMostDigit As String
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
    If Len(Trim(txtAutoSN.Text)) <> 11 Then
        MsgBox "A serial number must be 11 digits.", vbInformation, "Error"
        txtAutoSN.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtAutoSN.Text) Then
        MsgBox "A serial number consists of digits.", vbInformation, "Error"
        txtAutoSN.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtAutoAmount.Text) = False Or Val(txtAutoAmount.Text) <= 0 Then
        MsgBox "You must specify how many serial numbers you want to generate.", vbInformation, "Error"
        txtAutoAmount.SetFocus
        Exit Sub
    End If
    
    lstAutoSN.Clear
    
    'with serial numbers being 10 digits, we can easily exceed the limits of a long type
    'and other types will not hold 10 digits of resolution.  we must break the number
    'into 2 parts and work with the parts individually
    'lngUpperHalf = CLng(Left(txtAutoSN.Text, 5))
    ' 10/3/05 added to capture left most character
    strHoldLeftMostDigit = Left(txtAutoSN.Text, 1)
    '10/3/05 modified to select digits 10 - 6
    'lngUpperHalf = CLng(Left(txtAutoSN.Text, 5))
    lngUpperHalf = CLng(Mid(txtAutoSN.Text, 2, 5))
    lngLowerHalf = CLng(Right(txtAutoSN.Text, 5))
    For iLoop = 0 To Int(txtAutoAmount.Text) - 1
        'generate the lower half of the serial number by adding the step count
        lngNewLower = lngLowerHalf + iLoop
        If lngNewLower > 99999 Then
            'lower half exceeded 5 digit (now >= 100,000)
            'reduce to 5 digits
            lngNewLower = lngNewLower - 100000
            'increment upper half by 1 (least sig digit = 100,000)
            lngNewUpper = lngUpperHalf + 1
        Else
            'no overrun , use upper half as is
            lngNewUpper = lngUpperHalf
        End If
        
        strSN = strHoldLeftMostDigit & Format(lngNewUpper, "00000") & Format(lngNewLower, "00000")
        lstAutoSN.AddItem strSN
    Next iLoop
    
End Sub

Private Sub btnManualAdd_Click()
    Dim X As Integer
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
    txtManualSN.Text = Trim(txtManualSN.Text) ' DLimburg Change 2010_1 3Feb2010 - remove any preceeding spaces
    If Len(Trim(txtManualSN.Text)) <> 11 Then
        MsgBox "A serial number must be 11 digits.", vbInformation, "Error"
        txtManualSN.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtManualSN.Text) Then
        MsgBox "A serial number consists of digits.", vbInformation, "Error"
        txtManualSN.SetFocus
        Exit Sub
    End If
    If lstManualSN.ListCount <> -1 Then
        For X = 0 To lstManualSN.ListCount - 1
            If lstManualSN.List(X) = txtManualSN.Text Then
                MsgBox "Serial number " & txtManualSN.Text & " is already in the list.", vbInformation, "Error"
                txtManualSN.SetFocus
                Exit Sub
            End If
        Next
    End If
    lstManualSN.AddItem txtManualSN.Text
    txtManualSN.SetFocus
End Sub

Private Sub btnManualClear_Click()
    lstManualSN.Clear
    txtManualSN.SetFocus
End Sub

Private Sub btnSave_Click()
    '------------------------------------------------------------------------------------------------
    '   Rev'd:      10/21/03    By: Ray Raf
    '   Note:       Added handling of nicknames
    '------------------------------------------------------------------------------------------------
    Dim intListNum As Integer
    Dim strSNs() As String
    Dim bResult As Boolean
    Dim X As Integer
    Dim iLengthCatNum As Integer
    Dim strLookedUpCatNum As String
    
 ' Change by DLimburg on 5/30/06
 ' Change length test to accept a work order with a length of either 7 or 8
    If (Len(Trim(txtWO.Text)) < 7) Or (Len(Trim(txtWO.Text)) > 8) Then
        MsgBox "The work order must be exactly 7 or 8 digits.", vbInformation, "Error"
        txtWO.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtWO.Text) Then
        MsgBox "The work order only contains digits.", vbInformation, "Error"
        txtWO.SetFocus
        Exit Sub
    End If
    
    iLengthCatNum = Len(Trim(txtCatNum.Text))
    If iLengthCatNum < 10 And iLengthCatNum <> 6 Then
        MsgBox "The catalog number must be at least 10 characters.", vbInformation, "Error"
        txtCatNum.SetFocus
        Exit Sub
    End If
    
    If chkSN(0).Value = True Then
        If lstManualSN.ListCount <= 0 Then
            MsgBox "You need to enter in at least one serial number for this work order.", vbInformation, "Error"
            txtManualSN.SetFocus
            Exit Sub
        End If
    Else
        If lstAutoSN.ListCount <= 0 Then
            MsgBox "You need to enter in at least one serial number for this work order.", vbInformation, "Error"
            txtAutoSN.SetFocus
            Exit Sub
        End If
    End If
    
    txtWO.Text = UCase(txtWO.Text)
    If adoWorkOrderInDatabase(txtWO.Text) = True Then
        MsgBox "The work order " & txtWO.Text & " is already in the database.", vbInformation, "Error"
        txtWO.SetFocus
        Exit Sub
    End If
    
    'all cat numbers in database are capital letters, so cap whatever user entered
    txtCatNum.Text = UCase(txtCatNum.Text)
    If Len(txtCatNum.Text) < 10 Then
        'check to see if a valid nickname was entered
        strLookedUpCatNum = LookupNickname(txtCatNum.Text)
        If Len(strLookedUpCatNum) > 0 Then
            'cat number found for nickname
            txtCatNum.Text = UCase(strLookedUpCatNum)
        Else
            MsgBox "Invalid Nickname entered for catalog number.", vbInformation, "Error"
            txtCatNum.SetFocus
            Exit Sub
        End If
    End If
 ' if building a porc arrester, test to see that question for calcined disk has been answered
    If Left(txtCatNum, 1) = "A" And _
        opCalcinedYes = False And _
        opCalcinedNo = False Then
        MsgBox "Specify if calcined disks to be used, must select 'Yes' or 'No'.", vbInformation, "Error"
        Exit Sub
    End If
        
       
    '--This makes sure that there is data in the MCOV tables for the entered catalog number
   If VerifyCatalogNumber(Trim(txtCatNum.Text)) = False Then
        MsgBox "This catalog number is not valid.", vbInformation, "Error"
        txtCatNum.SetFocus
        Exit Sub
    End If
    
    
    If chkSN(0).Value = True Then
        'namually entering serial number list
        intListNum = lstManualSN.ListCount - 1
    Else
        'user selected auto generate
        intListNum = lstAutoSN.ListCount - 1
    End If
    
    ReDim strSNs(intListNum)
    'copy each of the serial numbers to an array for later use
    For X = 0 To intListNum
        If chkSN(0).Value = True Then
            'manual entry of each serial number
            strSNs(X) = lstManualSN.List(X)
        Else
            'automatically generate list of serial numbers
            strSNs(X) = lstAutoSN.List(X)
        End If
    Next X
    
    '--Checks to see if any units with the entered serial numbers already exist.
    'At this point, since this is a new work order it should not be posible, but
    'due to entry errors in legacy data, some of the existing serial numbers may
    'show units already built.
    If adoCheckForExistingSNsByList(strSNs) = True Then
        MsgBox "One or more of the serial numbers has already been built." & vbCr & "Saving aborted.", vbInformation, "Error"
        Exit Sub
    End If
       
    globCurrentWO.strWO = txtWO.Text
    '--the strType variable will be set when the user clicks next
    
    'create the WO in the database
    bResult = adoCreateWorkOrder(txtWO.Text, txtCatNum.Text, strSNs)
    If bResult = True Then
        MsgBox "The work order " & txtWO.Text & " has been successfully saved.", vbInformation, "Success"
Dim holder As String
holder = globCurrentUnit.DiskType

        'set the last WO number used
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastWO_Used", Setting:=txtWO.Text
        'Set it so user has no default for num disks to use
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNumDisksToUse", Setting:="0"
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum3kVDisksToUse", Setting:="0"
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum5kVDisksToUse", Setting:="0"
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastUnit", Setting:="0"
        'set global WO info
        With globCurrentWO
            .strWO = txtWO.Text
            .strCatNumber = txtCatNum.Text
            
            .strType = getType(txtCatNum.Text)
            .strBuildPath = getBuildPath(.strCatNumber, .strType)
            
            '--retrieve the number of serials
            .numSN = intListNum + 1
            
        End With
        
        
        'let calling routine know that WO was created OK
        globButtonResult = vbOK
        
        Unload Me
    End If
End Sub

Private Sub chkSN_Click(Index As Integer)
    If Index = 0 Then
        'doing manual entry of serial numbers
        btnManualAdd.Enabled = True
        lstManualSN.Enabled = True
        btnManualClear.Enabled = True
        txtManualSN.Enabled = True
        
        txtAutoSN.Enabled = False
        txtAutoAmount.Enabled = False
        btnGenerate.Enabled = False
        lstAutoSN.Enabled = False
        btnAutoClear.Enabled = False
        
        txtManualSN.SetFocus
    Else
        btnManualAdd.Enabled = False
        lstManualSN.Enabled = False
        btnManualClear.Enabled = False
        txtManualSN.Enabled = False
        
        txtAutoSN.Enabled = True
        txtAutoAmount.Enabled = True
        btnGenerate.Enabled = True
        lstAutoSN.Enabled = True
        btnAutoClear.Enabled = True
        
        txtAutoSN.SetFocus
    End If
End Sub

Private Sub Form_Load()
    txtWO.Text = frmWorkOrder.txtWorkOrder.Text
End Sub

Private Sub opCalcinedNo_Click()
    globCalcinedDisksUsedInPorc = False
End Sub

Private Sub opCalcinedYes_Click()
    globCalcinedDisksUsedInPorc = True
End Sub

Private Sub txtCatNum_Change()
    If UCase(Left(txtCatNum, 1)) = "A" Then
        picCalcineDiskBuild.Visible = True
        globCalcinedDisksUsedInPorc = False
        opCalcinedYes = False
        opCalcinedNo = False
    Else
        picCalcineDiskBuild.Visible = False
    End If
    
End Sub
