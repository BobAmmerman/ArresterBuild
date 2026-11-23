VERSION 5.00
Begin VB.Form frmNumDisks 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNum3kV_Disks_Gray 
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
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtNum5kV_Disks_Gray 
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
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2925
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtNum5kV_Disks_Calcined 
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
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2925
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtNum3kV_Disks_Calcined 
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
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   585
   End
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
      Left            =   2700
      TabIndex        =   18
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtNumDisks 
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
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2100
      Width           =   585
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "&Back"
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
      Left            =   4155
      TabIndex        =   6
      Top             =   3630
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
      Left            =   5280
      TabIndex        =   5
      Top             =   3630
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Gray"
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
      Index           =   6
      Left            =   3960
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Calcined"
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
      Index           =   7
      Left            =   5040
      TabIndex        =   21
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblLabel_5kVCnt 
      AutoSize        =   -1  'True
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
      Left            =   1710
      TabIndex        =   20
      Top             =   2985
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label lblLabel_3kVCnt 
      AutoSize        =   -1  'True
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
      Left            =   1710
      TabIndex        =   19
      Top             =   2565
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Total Number of Disks:"
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
      Top             =   2145
      Width           =   1890
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number:"
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
      Left            =   270
      TabIndex        =   16
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label lblSN 
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
      Left            =   1830
      TabIndex        =   15
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblUnit 
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
      Left            =   1830
      TabIndex        =   14
      Top             =   1605
      Width           =   2475
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Unit:"
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
      Index           =   2
      Left            =   270
      TabIndex        =   13
      Top             =   1605
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   285
      X2              =   4215
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Catalog Number:"
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
      Left            =   270
      TabIndex        =   12
      Top             =   240
      Width           =   1410
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Family:"
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
      Left            =   270
      TabIndex        =   11
      Top             =   825
      Width           =   585
   End
   Begin VB.Label lblCatalogNumber 
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
      Left            =   1830
      TabIndex        =   10
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblFamily 
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
      Left            =   1830
      TabIndex        =   9
      Top             =   825
      Width           =   1185
   End
   Begin VB.Label lblType 
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
      Left            =   1830
      TabIndex        =   8
      Top             =   525
      Width           =   1185
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
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
      Left            =   270
      TabIndex        =   7
      Top             =   525
      Width           =   435
   End
End
Attribute VB_Name = "frmNumDisks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBack_Click()
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    Load frmWOOptions
    frmWOOptions.Show
    frmWOOptions.ZOrder 0
    'setting the list index will cause the lstUnitLR_click event to fire
    'which will display and load the serial number listbox
    frmWOOptions.lstUnitLR.ListIndex = globCurrentUnit.UnitNumber - 1
    frmWOOptions.lstSerials.Text = globCurrentUnit.SN
    resizeForm
    Unload Me
End Sub

Private Sub btnCancel_Click()
    'cancels build of current arrester and resets to start of wizard
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
    CancelButton

End Sub

Private Sub btnNext_Click()
    Dim iNumDisks As Integer
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
'    Dim iNum3kVDisks As Integer
'    Dim iNum5kVDisks As Integer
    Dim iNum3kVDisks_Gray As Integer
    Dim iNum5kVDisks_Gray As Integer
    Dim iNum3kVDisks_Calcined As Integer
    Dim iNum5kVDisks_Calcined As Integer
    Dim bResult As Boolean
' DLimburg Change 2009_1 10Feb2009 Added code to allow for 3 and 5 kV Disk Build for the AT32G
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
    If Trim(txtNumDisks.Text) = "" Then
        MsgBox "You must specify the number of disks for this unit.", vbInformation, "Error"
        Exit Sub
    End If
    'make sure that they actually entered a number
    If Not (IsNumeric(Trim(txtNumDisks.Text))) Then
        MsgBox "You must enter numbers only for the number of disks!", vbInformation, "Error"
        Exit Sub
    Else
        iNumDisks = Int(Trim(txtNumDisks.Text))
'        If globCurrentUnit.DiskType = "AT32G" Then
        ' this test will change as the AT53 and AT54 5kV disk are introduced
'            If Not (IsNumeric(Trim(txtNum3kV_Disks.Text))) Then
'                iNum3kVDisks = 0
'            Else
'                iNum3kVDisks = Int(Trim(txtNum3kV_Disks.Text))
'            End If
'            If Not (IsNumeric(Trim(txtNum5kV_Disks.Text))) Then
'                iNum5kVDisks = 0
'            Else
'                iNum5kVDisks = Int(Trim(txtNum5kV_Disks.Text))
'            End If
'            If (iNum5kVDisks <> 0 Or iNum3kVDisks <> 0) And _
'                (iNum5kVDisks + iNum3kVDisks <> iNumDisks) Then
'                MsgBox "Sum of " & globCurrentUnit.DiskType & " disks and " & globCurrentUnit.DiskType2 & _
'                        " disks must equal total number of disks!", vbInformation, "Error"
'                Exit Sub
'            End If

' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
            If Not (IsNumeric(Trim(txtNum3kV_Disks_Gray.Text))) Then
                iNum3kVDisks_Gray = 0
            Else
                iNum3kVDisks_Gray = Int(Trim(txtNum3kV_Disks_Gray.Text))
            End If
            If Not (IsNumeric(Trim(txtNum5kV_Disks_Gray.Text))) Then
                iNum5kVDisks_Gray = 0
            Else
                iNum5kVDisks_Gray = Int(Trim(txtNum5kV_Disks_Gray.Text))
            End If
            If Not (IsNumeric(Trim(txtNum3kV_Disks_Calcined.Text))) Then
                iNum3kVDisks_Calcined = 0
            Else
                iNum3kVDisks_Calcined = Int(Trim(txtNum3kV_Disks_Calcined.Text))
            End If
            If Not (IsNumeric(Trim(txtNum5kV_Disks_Calcined.Text))) Then
                iNum5kVDisks_Calcined = 0
            Else
                iNum5kVDisks_Calcined = Int(Trim(txtNum5kV_Disks_Calcined.Text))
            End If
            If ((iNum5kVDisks_Gray + iNum3kVDisks_Gray + iNum5kVDisks_Calcined + iNum3kVDisks_Calcined) <> iNumDisks) Then
                MsgBox "Sum of " & globCurrentUnit.DiskType & " disks and " & globCurrentUnit.DiskType2 & _
                        " disks must equal total number of disks!", vbInformation, "Error"
                Exit Sub
            End If
'        Else
'            iNum3kVDisks = iNumDisks
'            iNum5kVDisks = 0
'        End If
    End If
    
'    If adoCheckAllowedDisks(iNumDisks, iNum3kVDisks, iNum5kVDisks) = True Then
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
    If adoCheckAllowedDisks(iNumDisks, iNum3kVDisks_Calcined, iNum5kVDisks_Calcined, iNum3kVDisks_Gray, iNum5kVDisks_Gray) = True Then
        globNumDisksToUse = iNumDisks
        globCurrentUnit.NumOfDisks = iNumDisks

        globCurrentUnit.NumOf5kVDisks = iNum5kVDisks_Gray + iNum5kVDisks_Calcined
        globCurrentUnit.NumOf3kVDisks = iNum3kVDisks_Gray + iNum3kVDisks_Calcined
'        SaveSet AppName:=App.Title, Section:="Memory", _
'                Key:="LastNumDisksToUse", Setting:=CStr(iNumDisks)
'        globCurrentWO.globNumDiskDefault = iNumDisks
'        SaveSet AppName:=App.Title, Section:="Memory", _
'                Key:="LastNum3kVDisksToUse", Setting:=CStr(iNum3kVDisks)
'        SaveSet AppName:=App.Title, Section:="Memory", _
'                Key:="LastNum5kVDisksToUse", Setting:=CStr(iNum5kVDisks)
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks        'MsgBox "The number of disks is valid, continuing to process (2.1) --- END OF PROGRAM"
        globCurrentUnit.NumOf3kVDisks_Gray = iNum3kVDisks_Gray
        globCurrentUnit.NumOf5kVDisks_Gray = iNum5kVDisks_Gray
        globCurrentUnit.NumOf3kVDisks_Calcined = iNum3kVDisks_Calcined
        globCurrentUnit.NumOf5kVDisks_Calcined = iNum5kVDisks_Calcined
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNumDisksToUse", Setting:=CStr(iNumDisks)
        globCurrentWO.globNumDiskDefault = iNumDisks
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum3kVGrayDisksToUse", Setting:=CStr(iNum3kVDisks_Gray)
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum5kVGrayDisksToUse", Setting:=CStr(iNum5kVDisks_Gray)
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum3kVCalcinedDisksToUse", Setting:=CStr(iNum3kVDisks_Calcined)
        SaveSet AppName:=App.Title, Section:="Memory", _
                Key:="LastNum5kVCalcinedDisksToUse", Setting:=CStr(iNum5kVDisks_Calcined)
        frmDiskQuantities.Show
        frmDiskQuantities.ZOrder 0
        resizeForm
        Unload Me
        
    End If
End Sub

Private Sub Form_Load()
    Dim iDefaultDisks As Integer
    Call setDefaultLocation(frmNumDisks)
    Call setCaption("Enter number of disks:")
    Call setProgress(2)
    Call moveForm
       
    lblCatalogNumber.Caption = globCurrentWO.strCatNumber
    lblType.Caption = getTypeName(globCurrentWO.strType)
    lblFamily.Caption = globCurrentUnit.Family
    lblSN.Caption = globCurrentUnit.SN
    lblUnit.Caption = Chr(globCurrentUnit.UnitNumber + 64)  '    converts numbers 1 -6 to letters A - F
    
    'Show the number of disks selected last time.  This number
    'will be set to zero if a different unit or work order is selected
    iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
                        Key:="LastNumDisksToUse", Default:="0"))
    If iDefaultDisks = 0 Then
        txtNumDisks.Text = ""
    Else
        txtNumDisks.Text = iDefaultDisks
    End If
    
' DLimburg Change 2009_1 10Feb2009 Added code to allow for 3 and 5 kV Disk Build for the AT32G
' ** Start of new code
'    If Left(globCurrentUnit.DiskType, 5) = "AT32G" Then
    
        lblLabel_3kVCnt.Caption = "Number of " & globCurrentUnit.DiskType & " Disks:"
        lblLabel_3kVCnt.Visible = True
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
'        txtNum3kV_Disks.Visible = True
        txtNum3kV_Disks_Gray.Visible = True
        txtNum3kV_Disks_Calcined.Visible = True
        lblLabel_5kVCnt.Caption = "Number of " & globCurrentUnit.DiskType2 & " Disks:"
        lblLabel_5kVCnt.Visible = True
'        txtNum5kV_Disks.Visible = True
        txtNum5kV_Disks_Gray.Visible = True
        txtNum5kV_Disks_Calcined.Visible = True
    
        'Show the number of 3kV disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
'        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
'                            Key:="LastNum3kVDisksToUse", Default:="0"))
'        If iDefaultDisks = 0 Then
'            txtNum3kV_Disks.Text = ""
'        Else
'            txtNum3kV_Disks.Text = iDefaultDisks
'        End If
        
        'Show the number of 5kV disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
'        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
'                            Key:="LastNum5kVDisksToUse", Default:="0"))
'        If iDefaultDisks = 0 Then
'            txtNum5kV_Disks.Text = ""
'        Else
'            txtNum5kV_Disks.Text = iDefaultDisks
'        End If

        'Show the number of 3kV GRAY disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
                            Key:="LastNum3kVGrayDisksToUse", Default:="0"))
        If iDefaultDisks = 0 Then
            txtNum3kV_Disks_Gray.Text = ""
        Else
            txtNum3kV_Disks_Gray.Text = iDefaultDisks
        End If
        
        'Show the number of 5kV GRAY disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
                            Key:="LastNum5kVGrayDisksToUse", Default:="0"))
        If iDefaultDisks = 0 Then
            txtNum5kV_Disks_Gray.Text = ""
        Else
            txtNum5kV_Disks_Gray.Text = iDefaultDisks
        End If
        
            'Show the number of 3kV CALCINED disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
                            Key:="LastNum3kVCalcinedDisksToUse", Default:="0"))
        If iDefaultDisks = 0 Then
            txtNum3kV_Disks_Calcined.Text = ""
        Else
            txtNum3kV_Disks_Calcined.Text = iDefaultDisks
        End If
        
        'Show the number of 5kV CALCINED disks selected last time.  This number
        'will be set to zero if a different unit or work order is selected
        iDefaultDisks = CInt(GetSet(AppName:=App.Title, Section:="Memory", _
                            Key:="LastNum5kVCalcinedDisksToUse", Default:="0"))
        If iDefaultDisks = 0 Then
            txtNum5kV_Disks_Calcined.Text = ""
        Else
            txtNum5kV_Disks_Calcined.Text = iDefaultDisks
        End If
'    Else
'        lblLabel_3kVCnt.Visible = False
'        txtNum3kV_Disks.Visible = False
'        lblLabel_5kVCnt.Visible = False
'        txtNum5kV_Disks.Visible = False
'    End If
' ** End of new code
    
End Sub

Private Sub txtNumDisks_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
        Case Else: KeyAscii = 0
    End Select
End Sub
