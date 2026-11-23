VERSION 5.00
Begin VB.Form frmWOOptions 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   7920
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
      Left            =   4125
      TabIndex        =   16
      Top             =   3900
      Width           =   1050
   End
   Begin VB.TextBox txtNumBuilt 
      Height          =   285
      Left            =   1725
      TabIndex        =   15
      Top             =   4200
      Width           =   840
   End
   Begin VB.TextBox txtNumRequired 
      Height          =   285
      Left            =   225
      TabIndex        =   13
      Top             =   4200
      Width           =   840
   End
   Begin VB.ListBox lstSerials 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmWOOptions.frx":0000
      Left            =   3480
      List            =   "frmWOOptions.frx":0002
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.ListBox lstUnitLR 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmWOOptions.frx":0004
      Left            =   270
      List            =   "frmWOOptions.frx":0006
      TabIndex        =   4
      Top             =   1530
      Width           =   2430
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      Left            =   6825
      TabIndex        =   1
      Top             =   3930
      Width           =   1050
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
      Left            =   5700
      TabIndex        =   0
      Top             =   3930
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "Number Built:"
      Height          =   390
      Left            =   1800
      TabIndex        =   14
      Top             =   3750
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Number Required:"
      Height          =   390
      Left            =   300
      TabIndex        =   12
      Top             =   3750
      Width           =   765
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
      TabIndex        =   10
      Top             =   525
      Width           =   435
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
      TabIndex        =   9
      Top             =   525
      Width           =   1185
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
      TabIndex        =   8
      Top             =   825
      Width           =   1185
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
      TabIndex        =   7
      Top             =   240
      Width           =   2655
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
      TabIndex        =   6
      Top             =   825
      Width           =   585
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Select 1Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   270
      TabIndex        =   5
      Top             =   1305
      Width           =   2430
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Select 1 SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   1305
      Visible         =   0   'False
      Width           =   2430
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
      TabIndex        =   2
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmWOOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBack_Click()
    Dim bResult As Boolean
    
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
    Load frmWorkOrder
    frmWorkOrder.Show
    frmWorkOrder.ZOrder 0
    frmWorkOrder.txtWorkOrder.Text = globCurrentWO.strWO
    resizeForm
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Dim bResult As Boolean
    
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    CancelButton
End Sub

Private Sub btnNext_Click()
    '------------------------------------------------------------------------------------------------
    '   Created:    ??????      By: Jeff Pierce
    '   Rev'd   04/16/03    By: Ray Raf
    '   Note:   changed handling of unit from global string to Unit number and MCOV
    '                   in globCurrentUnit variable.
    '------------------------------------------------------------------------------------------------
    Dim iLastPos As Integer
    Dim iNextPos As Integer
    Dim iLastUnit As Integer
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    
        If lstSerials.ListIndex = -1 Then
            MsgBox "You must select a serial number.", vbInformation, "Error": Exit Sub
        End If
        If lstUnitLR.ListIndex = -1 Then
            MsgBox "You must select a unit.", vbInformation, "Error": Exit Sub
        End If
        With globCurrentUnit
            .SN = lstSerials.Text
            'this should give us the position of the number for the unit in the string
            iLastPos = InStr(1, lstUnitLR.Text, "Unit ") + 5
            'ex string A) Unit 1: 5.1kV MCOV
            .UnitNumber = Mid(lstUnitLR.Text, iLastPos, 1)
            iLastPos = iLastPos + 2
            iNextPos = InStr(iLastPos, lstUnitLR.Text, "kV")
            .UnitMCOV = Trim(Mid(lstUnitLR.Text, iLastPos, iNextPos - iLastPos))
            globFormMain.sbStatusBar.Panels("SN").Text = "S/N: " & .SN
            globFormMain.sbStatusBar.Panels("Unit").Text = "Unit: " & Chr(.UnitNumber + 64)
            'check to see if this is the same unit that was last built
            iLastUnit = GetSet(AppName:=App.Title, Section:="Memory", _
                    Key:="LastUnit", Default:="0")
            If iLastUnit <> .UnitNumber Then
                'need to wipe out number of disks
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastUnit", Setting:=CStr(.UnitNumber)
                SaveSet AppName:=App.Title, Section:="Memory", _
                        Key:="LastNumDisksToUse", Setting:="0"
            End If
        End With
        Load frmNumDisks
        frmNumDisks.Show
        frmNumDisks.ZOrder 0
        resizeForm
        Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim rstCount As New ADODB.Recordset
    Dim SQL As String
    
    Call setDefaultLocation(frmWOOptions)
    Call setCaption("Work Order: " & globCurrentWO.strWO)
    Call setProgress(1)
    Call moveForm
    
'    For x = 1 To globCurrentWO.numSN
'        lstSerials.AddItem globCurrentWO.strSN(x)
'    Next
    
    lblCatalogNumber.Caption = globCurrentWO.strCatNumber
    lblType.Caption = getTypeName(globCurrentWO.strType)
    lblFamily.Caption = globCurrentUnit.Family
    
    Call adoGetUnitLettersAndRatings
    
    SQL = "SELECT * FROM WO_Info WHERE WO_Number = '" & globCurrentWO.strWO & "' "
    rstCount.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    
    txtNumRequired.Text = rstCount!NumArrestersRequired
    
End Sub

Private Sub Form_Paint()
    resizeForm
End Sub

Private Sub lstSerials_Click()
    'don't want user to think they can go on before selecting
    'a serial number
    btnNext.Enabled = True
End Sub

Private Sub lstUnitLR_Click()
    Dim SQL As String
    Dim iUnit As Integer
    Dim rstSN As New ADODB.Recordset
    Dim rstBuilt As New ADODB.Recordset
    Dim rstWO As New ADODB.Recordset
    Dim strField As String
    
    iUnit = lstUnitLR.ListIndex + 1
    
    'the user just selected a unit for the list.  Get the serial numbers that have
    'been built for it.
    SQL = "SELECT  d.SerialNumber "
    SQL = SQL & "FROM  BuiltUnits b INNER JOIN "
    SQL = SQL & "WO_SerialNumbers d ON b.SerialNumber = d.SerialNumber "
    SQL = SQL & "WHERE  (b.Unit = " & iUnit & " ) "
    SQL = SQL & "AND (d.WO_Number = '" & globCurrentWO.strWO & "') "
    SQL = SQL & "Order by d.SerialNumber "
    rstSN.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    
    'now get all of the serial numbers for the WO
    SQL = "SELECT * FROM WO_SerialNumbers "
    SQL = SQL & "WHERE WO_Number = '" & globCurrentWO.strWO & "' "
    rstWO.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    
    
    lstSerials.Clear
    If rstSN.EOF = False Then
        'some serial numbers already built, don't show these
        Do While rstWO.EOF = False
            rstSN.MoveFirst
            'look for this serial number in built units data
            rstSN.Find "SerialNumber = '" & rstWO!SerialNumber & "' "
            If rstSN.EOF = True Then
                'WO serial number not found in built units, display it
                lstSerials.AddItem rstWO!SerialNumber
            End If
            rstWO.MoveNext
        Loop
    Else
        'no serial number built, show all
        Do While rstWO.EOF = False
            lstSerials.AddItem rstWO!SerialNumber
            rstWO.MoveNext
        Loop
    End If
    
    If rstSN.State > 0 Then
        rstSN.Close
    End If
    
    If rstWO.State > 0 Then
        rstWO.Close
    End If
    
    'show the list box
    lstSerials.Visible = True
    lblLabels(2).Visible = True
    
    SQL = "SELECT * FROM WO_Info WHERE WO_Number = '" & globCurrentWO.strWO & "' "
    rstWO.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    strField = "Unit" & iUnit & "NumBuilt"
    
    Label2.Visible = True
    txtNumBuilt.Visible = True
    txtNumBuilt.Text = rstWO(strField).Value
    
    If rstWO.State > 0 Then
        rstWO.Close
    End If
    
    
End Sub
