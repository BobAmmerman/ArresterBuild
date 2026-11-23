VERSION 5.00
Begin VB.Form frmPrintMultiple 
   Caption         =   "Print Build Reports for Multiple S/N's"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnits 
      Caption         =   "Select Units"
      Height          =   1860
      Left            =   480
      TabIndex        =   24
      Top             =   4104
      Width           =   3084
      Begin VB.CheckBox chkUnits 
         Caption         =   "F Units"
         Height          =   240
         Index           =   5
         Left            =   1812
         TabIndex        =   15
         Top             =   1452
         Width           =   972
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "E Units"
         Height          =   240
         Index           =   4
         Left            =   1812
         TabIndex        =   14
         Top             =   1140
         Width           =   972
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "D Units"
         Height          =   240
         Index           =   3
         Left            =   1812
         TabIndex        =   13
         Top             =   828
         Width           =   972
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "C Units"
         Height          =   240
         Index           =   2
         Left            =   528
         TabIndex        =   12
         Top             =   1440
         Width           =   972
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "B Units"
         Height          =   240
         Index           =   1
         Left            =   528
         TabIndex        =   11
         Top             =   1122
         Width           =   972
      End
      Begin VB.CheckBox chkUnits 
         Caption         =   "A Units"
         Height          =   240
         Index           =   0
         Left            =   528
         TabIndex        =   10
         Top             =   804
         Width           =   972
      End
      Begin VB.CheckBox ChkAll 
         Caption         =   "All Built Units"
         Height          =   264
         Left            =   276
         TabIndex        =   9
         Top             =   384
         Width           =   1440
      End
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
      Left            =   408
      TabIndex        =   0
      Top             =   204
      Value           =   -1  'True
      Width           =   3255
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
      Left            =   4464
      TabIndex        =   4
      Top             =   204
      Width           =   3255
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
      Height          =   1980
      ItemData        =   "frmPrintMultiple.frx":0000
      Left            =   444
      List            =   "frmPrintMultiple.frx":0002
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   972
      Width           =   3060
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
      ItemData        =   "frmPrintMultiple.frx":0004
      Left            =   4560
      List            =   "frmPrintMultiple.frx":0006
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1392
      Width           =   2730
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
      Left            =   456
      TabIndex        =   3
      Top             =   3516
      Width           =   975
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
      Left            =   6324
      TabIndex        =   8
      Top             =   3516
      Width           =   975
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
      Left            =   960
      MaxLength       =   12
      TabIndex        =   1
      Top             =   600
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
      Left            =   2904
      TabIndex        =   2
      Top             =   600
      Width           =   600
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
      Left            =   5460
      MaxLength       =   12
      TabIndex        =   5
      Top             =   600
      Width           =   1830
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
      Left            =   5460
      MaxLength       =   4
      TabIndex        =   6
      Top             =   996
      Width           =   720
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
      Left            =   6276
      TabIndex        =   7
      Top             =   996
      Width           =   1005
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Print"
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
      Left            =   5844
      TabIndex        =   16
      Top             =   4944
      Width           =   1545
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
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
      Height          =   405
      Left            =   5844
      TabIndex        =   17
      Top             =   5520
      Width           =   1545
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
      Height          =   216
      Index           =   2
      Left            =   3816
      TabIndex        =   23
      Top             =   264
      Width           =   396
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
      Height          =   228
      Index           =   3
      Left            =   468
      TabIndex        =   22
      Top             =   636
      Width           =   300
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
      Height          =   228
      Index           =   4
      Left            =   4500
      TabIndex        =   21
      Top             =   636
      Width           =   708
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
      Height          =   228
      Index           =   5
      Left            =   4524
      TabIndex        =   20
      Top             =   1044
      Width           =   888
   End
End
Attribute VB_Name = "frmPrintMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ReportWO As WorkOrder
Private m_ReportUnit As Unit

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
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
    txtAutoSN.Text = Trim(txtAutoSN.Text)   ' DLimburg Change 2010_1 3Feb2010 - remove any preceeding spaces
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
    lngUpperHalf = CLng(Left(txtAutoSN.Text, 5))
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
        
        strSN = Format(lngNewUpper, "00000") & Format(lngNewLower, "00000")
        lstAutoSN.AddItem strSN
    Next iLoop
    
End Sub

Private Sub btnManualAdd_Click()
    Dim X As Integer
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
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

Private Sub PrintControl()
    '------------------------------------------------------------------------------------------------
    '   Created:      05/01/04    By: Ray Raf
    '   Purpose:      Print multiple unit reports based on entered info.
    '------------------------------------------------------------------------------------------------
    Dim intListNum As Integer
    Dim strSNs() As String
    Dim bResult As Boolean
    Dim X As Integer
    Dim iLengthCatNum As Integer
    Dim strLookedUpCatNum As String
    Dim iSNLoop As Integer, iUnitLoop As Integer
    
    If chkSN(0).Value = True Then
        If lstManualSN.ListCount <= 0 Then
            MsgBox "You need to enter in at least one serial number to report on.", vbInformation, "Error"
            txtManualSN.SetFocus
            Exit Sub
        End If
    Else
        If lstAutoSN.ListCount <= 0 Then
            MsgBox "You need to enter in at least one serial number to report on.", vbInformation, "Error"
            txtAutoSN.SetFocus
            Exit Sub
        End If
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
    
    'we now know what serial numbers need to be printed.
    For iSNLoop = 0 To intListNum
        For iUnitLoop = 1 To 6
            If ChkAll.Value = 1 Or chkUnits(iUnitLoop - 1).Value = 1 Then
                'only print unit if all units or that particular unit is selected.
                bResult = CheckForBuiltUnit(strSNs(iSNLoop), CStr(iUnitLoop))
                If bResult = True Then
                    'unit found, go ahead and print
                    'If porcelain arrester, test to see if build uses calcined disks
                    If Test4CalcinedDiskInPorc(strSNs(iSNLoop), iUnitLoop) = False Then
                        MsgBox "Error serial number and unit not found."
                        Exit Sub
                    End If
                    PrintSingleUnit strSNs(iSNLoop), iUnitLoop
                End If
            End If
        Next iUnitLoop
    Next iSNLoop
       
    'let calling routine know that printing is completed
    globButtonResult = vbOK
     
    MsgBox "All Requested Unit Reports Printed!"
    
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim bUnitSelected As Boolean
    Dim iLoop As Integer
    
    'check to make sure that user selected units to print for
    'each listed serial number.
    bUnitSelected = False
    If Not (ChkAll.Value = 1) Then
        'user did not select all available units
        For iLoop = 0 To 5
            If chkUnits(iLoop).Value = 1 Then
                bUnitSelected = True
                Exit For
            End If
        Next iLoop
        If bUnitSelected = False Then
            MsgBox "You must select 'All Units' or at least one specific unit."
            Exit Sub
        End If
    End If
    
    'if we get this far, user selected all units or at least one individual unit
    PrintControl
    
End Sub

Private Sub ChkAll_Click()
    Dim bEnabled As Boolean
    Dim iUnitLoop As Integer
    
    If ChkAll.Value = 1 Then
        bEnabled = False
    Else
        bEnabled = True
    End If
    
    'disable all individual unit selections if "All Units" is selected
    For iUnitLoop = 0 To 5
        chkUnits(iUnitLoop).Enabled = bEnabled
    Next iUnitLoop
    
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub PrintSingleUnit(strSN As String, iUnitNum As Integer)
    '------------------------------------------------------------------------------------------------
    'Created:   12/18/03    By:Ray Rafalowski
    'Purpose:   To print the unit report for a single unit.
    '-------------------------------------------------------------------------------------------------
    Dim strMsg As String
    Dim bResult As Boolean
    
    'load unit data
    m_ReportUnit.SN = strSN
    m_ReportUnit.UnitNumber = iUnitNum
    bResult = LoadUnitInfo(m_ReportWO, m_ReportUnit)
    If bResult = False Then
        'could not retreive any/all info for the arrester.  Can't print report
        Exit Sub
    End If
    
    globCurrentWO = m_ReportWO
    globCurrentUnit = m_ReportUnit
            
    SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFilename", "c:\Data\" & m_ReportUnit.SN & ".pdf"
    Set frmUnitReport.dbconBuilt = dbconResultsData
    frmUnitReport.PrintUnitReport Printer
    Printer.EndDoc

End Sub


