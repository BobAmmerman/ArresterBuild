VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDiskQuantities 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOr 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "frmDiskQuantities.frx":0000
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtDiskType2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7320
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2292
   End
   Begin VB.TextBox txtDiskType 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2292
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   924
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "frmDiskQuantities.frx":0005
      Top             =   600
      Width           =   2292
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
      Left            =   5475
      TabIndex        =   15
      Top             =   5160
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
      Left            =   6885
      TabIndex        =   3
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next"
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
      Left            =   7995
      TabIndex        =   2
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton btnVerify 
      Caption         =   "&Done"
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
      Left            =   2985
      TabIndex        =   5
      Top             =   435
      Width           =   1050
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   510
      Width           =   405
   End
   Begin VB.Frame fraDiskNum 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   315
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   8820
      Begin MSFlexGridLib.MSFlexGrid grdDiskNum 
         Height          =   1770
         Left            =   15
         TabIndex        =   8
         Top             =   285
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   6
         Cols            =   7
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblV10kACAT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6495
         TabIndex        =   14
         Top             =   2100
         Width           =   960
      End
      Begin VB.Label lblV10kA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5220
         TabIndex        =   13
         Top             =   2100
         Width           =   960
      End
      Begin VB.Label lblV1mA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3810
         TabIndex        =   12
         Top             =   2100
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Sums:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3165
         TabIndex        =   11
         Top             =   2100
         Width           =   450
      End
      Begin VB.Label lblAvg10kACat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1695
         TabIndex        =   10
         Top             =   2100
         Width           =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Average of 10kA Cat:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Provide batch and disk number information:"
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
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   3570
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdQuantities 
      Height          =   1900
      Left            =   330
      TabIndex        =   0
      Top             =   435
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3360
      _Version        =   393216
      Rows            =   7
      Cols            =   3
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Provide disk quantities and categories:"
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
      Left            =   345
      TabIndex        =   4
      Top             =   120
      Width           =   3180
   End
End
Attribute VB_Name = "frmDiskQuantities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Init As Boolean
Dim m_PosibleBarcode As Boolean     'indicates that user may have scanned a barcode into a cell
Dim strLastBatch As String
Public m_NumDisksEntered As Integer
Public m_AwaitingFirstChar As Boolean
Dim SumV1mA, SumV10kA, SumV10kCAT As Double
        
Private Sub btnBack_Click()
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    Load frmNumDisks
    frmNumDisks.Show
    frmNumDisks.ZOrder 0
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


Private Function myGetDiskData() As Boolean
    '----------------------------------------------------------------------------------------
    '   Created:    April 03    By: Jeff P
    '   Purpose:    To retrieve all disk data from the database, show
    '                    the data in the grid and check it against limits
    '                    for the arrester.
    '   Rev'd    01/23/04   By: Pete Wisniewski
    '  Returns:  True if all the data was loaded and verified, False otherwise
    '   Note:   converted from a private sub to a private function....this code
    '               used to belong to the the btnGetDiskData_Click sub
    '   Rev'd:      03/13/04    By: Ray Raf
    '   Note:       Changing from cheking V1ma against limits to checking
    '                  vref against limits.  Limits are looked up in table.
    '----------------------------------------------------------------------------------------
    
    Dim iLoop As Integer
    Dim Found As Boolean
    Dim dblTemp As Double
    Dim strQuanTable As String
    Dim strSQL As String
    Dim strMsg As String
    Dim bResult As Boolean
    Dim strCurrentBatch As String, strCurrentDisk As String
    Dim iCheckLoop As Integer
    Dim sglHighCum As Single, sglLowCum As Single, sglV10k
    Dim strFamily As String
    Dim rstVrefLimits As New ADODB.Recordset
    Dim sglMCOV As Single, sglVrefLow As Single, sglVrefHigh As Single
    Dim sqlQueryString As String
    ' initiate to false
    myGetDiskData = False
        
    'check to make sure something is entered for every row.
    Found = False
    With grdDiskNum
        For iLoop = 1 To .Rows - 1
            strCurrentBatch = Trim(.TextMatrix(iLoop, 1))
            strCurrentDisk = Trim(.TextMatrix(iLoop, 2))
            'make sure something is entered in this row
            If strCurrentBatch = "" Or strCurrentDisk = "" Then
                MsgBox "Row " & iLoop & " must have a batch number and a disk number.", vbInformation, "Error"
                grdDiskNum.SetFocus
                Exit Function
            Else
                'batch and disk found.  Make sure it isn't identical to another disk
                If iLoop < (.Rows - 1) Then
                    For iCheckLoop = iLoop + 1 To .Rows - 1
                        If strCurrentBatch = Trim(.TextMatrix(iCheckLoop, 1)) And _
                            strCurrentDisk = Trim(.TextMatrix(iCheckLoop, 2)) Then
                                'disks in row iloop and icheckloop are identical
                                strMsg = "The disks in rows " & iLoop & " and " & iCheckLoop & " are duplicates!"
                                MsgBox strMsg
                                Found = True
                        End If
                    Next iCheckLoop
                End If
            End If
        Next iLoop
    End With
    
    'we can't proceed if duplicate disks are entered.
    If Found = True Then Exit Function
    
    lblAvg10kACat.Caption = "?"
    lblV1mA.Caption = "?"
    lblV10kA.Caption = "?"
    lblV10kACAT.Caption = "?"
    
    
    bResult = adoGetDiskData
    If bResult = False Then
        'all disk data not found
        Exit Function
    End If
    
    bResult = checkDiskTypes
    If bResult = False Then
        Exit Function
    End If
    
    bResult = checkCatsAndGetSums
    If bResult = False Then
        Exit Function
    End If
    
    Select Case globCurrentWO.strType
        Case 1: 'porc
            strQuanTable = "DiskQuanPorc_Rev1"   ' changed table name from DiskQuanPorc to DiskQuanPorc_Rev1 - DLimburg16Feb2011
            strFamily = globCurrentUnit.Family
            sqlQueryString = strSQL & "AND DiskQuantity_Total =" & m_NumDisksEntered & " "    'added DLimburg 16Feb2011
        Case 2: 'poly
' DLimburg Change 2009_1 changed SQL table to DiskQuanPoly_Rev1 from DiskQuanPoly
'            strQuanTable = "DiskQuanPoly_Rev1"
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
            strQuanTable = "DiskQuanPoly_Rev2"
            strFamily = globCurrentUnit.Family
            sqlQueryString = strSQL & "AND DiskQuantity=" & m_NumDisksEntered & " "  'added DLimburg 16Feb2011
    End Select
    
    'get values to compare to for this unit
    strSQL = "SELECT * FROM " & strQuanTable & " WHERE DiskType='" & globCurrentUnit.DiskType & "' "
    strSQL = strSQL & "AND Design_MCOV='" & globCurrentUnit.UnitMCOV & "' "
    strSQL = strSQL & "AND Family = '" & strFamily & "' "
    
    strSQL = strSQL & sqlQueryString

    strSQL = strSQL & " AND DiskQuantity_3kV = " & globCurrentUnit.NumOf3kVDisks & " AND DiskQuantity_5kV = " & globCurrentUnit.NumOf5kVDisks

    
    Set adoRS = dbconSetupData.Execute(strSQL)
    If adoRS.EOF Then
        strMsg = "The Cum_V10CAT_Min and Cum_V10CAT_Max cannot be retrieved for " & vbCrLf
        strMsg = strMsg & "Disk Type = " & globCurrentUnit.DiskType & " and MCOV = " & globCurrentUnit.UnitMCOV & "."
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "myGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
        grdDiskNum.SetFocus
        Exit Function
    End If
    
    'check to make sure the total of the categories is within the limits
    sglV10k = CSng(lblV10kACAT.Caption)
    sglLowCum = CSng(Trim(adoRS!Cum_V10CAT_Min))
    sglHighCum = CSng(Trim(adoRS!Cum_V10CAT_Max))
    If sglV10k < sglLowCum Or _
        sglV10k > sglHighCum Then
        strMsg = "The Category sum of " & Format(sglV10k, "####0") & " must be between " _
                    & Format(sglLowCum, "####0") & " and " _
                    & Format(sglHighCum, "####0") & "."
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "myGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
        grdDiskNum.SetFocus
        Exit Function
    End If
    
    'make sure V1mA or VRef is within limits
    strSQL = "SELECT * FROM ArresterTestParameters "
    strSQL = strSQL & "WHERE ArresterFamily = '" & globCurrentUnit.Family & "' "
    strSQL = strSQL & "AND MinRating <= " & globCurrentUnit.Rating & " "
    strSQL = strSQL & "AND MaxRating >= " & globCurrentUnit.Rating & " "
    rstVrefLimits.Open strSQL, dbconSetupData, adOpenDynamic, adLockReadOnly
    
    If rstVrefLimits.EOF = True Then
        strMsg = "Could not find Vref limits in table ArresterTestParameters"
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "myGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
        Exit Function
    End If
    
    sglMCOV = CSng(Trim(adoRS!Design_MCOV))
    sglVrefLow = rstVrefLimits!MinVrefFactor * sglMCOV
    sglVrefHigh = rstVrefLimits!MAxVrefFactor * sglMCOV
    dblTemp = CDbl(lblV1mA.Caption)
    If dblTemp < sglVrefLow Or dblTemp > sglVrefHigh Then
        strMsg = "The Vref Sum(" & dblTemp & ") is not between the limits of " & sglVrefLow & " and "
        strMsg = strMsg & sglVrefHigh & ".  This unit cannot be built."
        MsgBox strMsg, vbInformation, "Error"
        ErrorLogger "myGetDiskData", strMsg      'this is a data error. log it so someone can follow up on it
        grdDiskNum.SetFocus
        Exit Function
    End If

    myGetDiskData = True
    
End Function

Public Function checkCatsAndGetSums() As Boolean
    '----------------------------------------------------------------------------------------
    '   Created:    April 03    By: Jeff P
    '   Purpose:    To generate the V1ma, V10kA, and V10kA catagory sums.
    '                    Addistionally the routine will check to make sure that
    '                    the quantity of disks entered with given categories
    '                    matches the number entered for each category
    '                    on the upper grid.
    '----------------------------------------------------------------------------------------
    
    Dim strCats() As String  'this array will contain all distinct categories found in the disk grid
    Dim intCount() As Integer   'This array contains the number of disks found for the category in the matching position in strCats
    Dim X, Y As Integer
    Dim Found As Boolean
    Dim strMsg As String
    Dim AvgV10kCATs As Double
    Dim CategoryHolder As String
    Dim DiskTypeHolder As String
    'handle 1st row by itself
    ReDim strCats(1)
    ReDim intCount(1)
    
'    strCats(1) = Trim(grdDiskNum.TextMatrix(1, 5))
'    If Int(Val(grdDiskNum.TextMatrix(1, 5))) >= 200 And Int(Val(grdDiskNum.TextMatrix(1, 5))) <= 530 Then
    ' DLimburg Change 2009_1 added this if statement to cover 5kV disk highs and lows
'  DLimburg Change 2010_2 March 4, 2010 moved this test and change to adoGetDiskData
    'Dlimburg Change 2009_2 changed above test to test on disk type rather that category range
    'Testing on cat's with a min does not cover all cat's for the 5kV Disk
    'and runs into conflict with othe disks
    'If Left(grdDiskNum.TextMatrix(1, 6), 4) = "AT52" Then
    '        strCats(1) = Trim(grdDiskNum.TextMatrix(1, 5)) + 1000
    'Else
        strCats(1) = Trim(grdDiskNum.TextMatrix(1, 5))
    'End If

    
    
    intCount(1) = 1
    
    SumV1mA = grdDiskNum.TextMatrix(1, 3)
    SumV10kA = grdDiskNum.TextMatrix(1, 4)
   ' SumV10kCAT = grdDiskNum.TextMatrix(1, 5)
    
'    If Int(Val(grdDiskNum.TextMatrix(1, 5))) >= 200 And Int(Val(grdDiskNum.TextMatrix(1, 5))) <= 530 Then
    ' DLimburg Change 2009_1 added this if statement to cover 5kV disk highs and lows
'  DLimburg Change 2010_2 March 4, 2010 moved this test and change to adoGetDiskData
'    If Left(grdDiskNum.TextMatrix(1, 6), 4) = "AT52" Then
        'Dlimburg Change 2009_2 changed above test to test on disk type rather that category range
        'Testing on cat's with a min does not cover all cat's for the 5kV Disk
        'and runs into conflict with othe disks
'       SumV10kCAT = CDbl(grdDiskNum.TextMatrix(1, 5)) + 1000
'    Else
        SumV10kCAT = CDbl(grdDiskNum.TextMatrix(1, 5))
'    End If
    
    'loop through the rest of the rows
    For X = 2 To grdDiskNum.Rows - 1
        Found = False
        For Y = 1 To UBound(strCats)
            'loop trough all of the cats found so far to see if the category for this
            'disk matchs any found so far
            
           ' If Int(Val(grdDiskNum.TextMatrix(X, 5))) >= 200 And Int(Val(grdDiskNum.TextMatrix(X, 5))) <= 530 Then
            ' DLimburg Change 2009_1 added this if statement to cover 5kV disk highs and lows
            'DLimburg Change 2010_2 March 4, 2010 moved this test and change to adoGetDiskData
            ' If Left(grdDiskNum.TextMatrix(X, 6), 3) = "AT52" Then
                'Dlimburg Change 2009_2 changed above test to test on disk type rather that category range
                'Testing on cat's with a min does not cover all cat's for the 5kV Disk
                'and runs into conflict with othe disks
            '    CategoryHolder = CDbl(grdDiskNum.TextMatrix(1, 5)) + 1000
            'Else
                CategoryHolder = CDbl(grdDiskNum.TextMatrix(X, 5))
            'End If
           
       '     If grdDiskNum.TextMatrix(X, 5) = strCats(Y) Then
            If CategoryHolder = strCats(Y) Then
                intCount(Y) = intCount(Y) + 1
                Found = True
                Exit For
            End If
        Next Y
        
        If Found = False Then
            'the category wasn't found.  expand arrays to make room for this new category
            ReDim Preserve strCats(UBound(strCats) + 1)
            ReDim Preserve intCount(UBound(intCount) + 1)
            
           ' strCats(UBound(strCats)) = grdDiskNum.TextMatrix(X, 5)
            'If Int(Val(grdDiskNum.TextMatrix(X, 5))) >= 200 And Int(Val(grdDiskNum.TextMatrix(X, 5))) <= 530 Then
            ' DLimburg Change 2009_1 added this if statement to cover 5kV disk highs and lows
            'DLimburg Change 2010_2 March 4, 2010 moved this test and change to adoGetDiskData
'            If Left(grdDiskNum.TextMatrix(X, 6), 3) = "AT52" Then
            'Dlimburg Change 2009_2 changed above test to test on disk type rather that category range
            'Testing on cat's with a min does not cover all cat's for the 5kV Disk
            'and runs into conflict with othe disks
'                strCats(UBound(strCats)) = (grdDiskNum.TextMatrix(1, 5)) + 1000
'            Else
                strCats(UBound(strCats)) = grdDiskNum.TextMatrix(X, 5)
'            End If
            intCount(UBound(strCats)) = 1
        End If
        
        'update sums
        SumV1mA = SumV1mA + CDbl(grdDiskNum.TextMatrix(X, 3))
        SumV10kA = SumV10kA + CDbl(grdDiskNum.TextMatrix(X, 4))
        
'        If Int(Val(grdDiskNum.TextMatrix(X, 5))) >= 200 And Int(Val(grdDiskNum.TextMatrix(X, 5))) <= 530 Then
        ' DLimburg Change 2009_1 added this if statement to cover 5kV disk highs and lows
        ' DLimburg Change 2010_2 March 4, 2010 moved this test and change to adoGetDiskData
        'If Left(grdDiskNum.TextMatrix(X, 6), 3) = "AT5" Then
        'Dlimburg Change 2009_2 changed above test to test on disk type rather that category range
        'This is a more correct way to test
        '    SumV10kCAT = SumV10kCAT + CDbl(grdDiskNum.TextMatrix(X, 5)) + 1000
       ' Else
            SumV10kCAT = SumV10kCAT + CDbl(grdDiskNum.TextMatrix(X, 5))
       ' End If
    Next X
    
    AvgV10kCATs = SumV10kCAT / m_NumDisksEntered
    
    lblAvg10kACat.Caption = Format(AvgV10kCATs, "###0")
    lblV1mA.Caption = Format(SumV1mA, "###0.00")
    lblV10kA.Caption = Format(SumV10kA, "###0.00")
    lblV10kACAT.Caption = Format(SumV10kCAT, "###0")
    
    
    'this section matches the quantities and categories of the entered disks
    'with the expected quantities entered in the upper grid.
    For X = 1 To grdQuantities.Rows - 1
        If grdQuantities.TextMatrix(X, 1) <> "" And grdQuantities.TextMatrix(X, 2) <> "" Then
            Found = False
            For Y = 1 To UBound(strCats)
                If grdQuantities.TextMatrix(X, 2) = strCats(Y) Then
                    If Int(grdQuantities.TextMatrix(X, 1)) = intCount(Y) Then
                        'an exact match for quantity and category were found
                        Found = True
                        Exit For
                    End If
                End If
            Next Y
            
            If Found = False Then
                strMsg = "The disk categories and quantities for the disks entered don't match requirements."
                MsgBox strMsg, vbInformation, "Error"
                checkCatsAndGetSums = False
                grdDiskNum.SetFocus
                Exit Function
            End If
        End If
    Next X
    
    checkCatsAndGetSums = True
End Function

Public Function checkDiskTypes() As Boolean
    '----------------------------------------------------------------------------------------
    '   Created:    10/30/03    By: Ray Raf
    '   Purpose:    To check the disk type for each disk to see if it is
    '                    compatable with the disk(s) allowed for this arrester
    '----------------------------------------------------------------------------------------
    
    Dim strTypes() As String  'this array will contain all distinct Disk types found in the disk grid
    Dim X, Y As Integer
    Dim Found As Boolean
    Dim strMsg As String
    Dim bBadDiskTypeFound As Boolean
    Dim strBaseDiskSize As String
    Dim strBaseDiskPartner As String
    Dim iMatchesFound As Integer
    Dim EpoxyDiskFound As Boolean   ' DLimburg  Change 2009_1 added because current test will no longer
    Dim GlassDiskFound As Boolean   '           function properly because there cant be two types of types
                                    '           the 3kV and the 5kV
    
    checkDiskTypes = False
    bBadDiskTypeFound = False
    GlassDiskFound = False
    EpoxyDiskFound = False
    
    'handle 1st row by itself
    ReDim strTypes(1)
    strTypes(1) = Trim(grdDiskNum.TextMatrix(1, 6))
    
    'loop through the rest of the rows
    For X = 2 To grdDiskNum.Rows - 1
        Found = False
        For Y = 1 To UBound(strTypes)
            'loop through all of the types found so far to see if the type for this
            'disk matchs any found so far
            If grdDiskNum.TextMatrix(X, 6) = strTypes(Y) Then
                Found = True
                Exit For
            End If
        Next Y
        
        If Found = False Then
            'the category wasn't found.  expand arrays to make room for this new category
            ReDim Preserve strTypes(UBound(strTypes) + 1)
            strTypes(UBound(strTypes)) = grdDiskNum.TextMatrix(X, 6)
        End If
        
    Next X
    
    strBaseDiskSize = Left(globCurrentUnit.DiskType, 4)
    strBaseDiskPartner = Left(globCurrentUnit.DiskType2, 4)
    iMatchesFound = 0
    
    'check each distinct disk type to see if it is allowed
    For Y = 1 To UBound(strTypes)
        If Left(strTypes(Y), 4) = strBaseDiskSize Or _
              Left(strTypes(Y), 4) = strBaseDiskPartner Then
            'looking to see how many times the base disk size (ie AT32,  AT33, AT34) is matched.
            'if matched more than once, operator is trying to mix glass and epoxy collars in the
            'same arrester.  Thihs is not allowed.
            iMatchesFound = iMatchesFound + 1
            ' DLimburg  Change 2009_1 test added to insure glass and epoxy disks are not mixed
            If Mid(strTypes(Y), 5, 1) = "G" Then
                GlassDiskFound = True
            Else
                If Mid(strTypes(Y), 5, 1) = "E" Then
                    EpoxyDiskFound = True
                End If
            End If
        End If
        
' DLimburg  Change 2009_1  Replace code line - If strTypes(Y) <> globCurrentUnit.DiskType then
        If strTypes(Y) = globCurrentUnit.DiskType Or _
            strTypes(Y) = globCurrentUnit.DiskType2 Then
        ' do nothing, disk matches primary disk or its partner
        Else
            Found = True
            'not an exact match, need to check if it can be used
            Found = CrossRefDiskType(globCurrentUnit.DiskType, strTypes(Y), globCurrentWO.strType)
            If Found = False Then
                MsgBox "Disks of type '" & strTypes(Y) & "' cannot be used in this arrester!"
                bBadDiskTypeFound = True
            End If
        End If
    Next Y
    
    ' If iMatchesFound > 1 Then
    ' DLimburg  Change 2009_1 test changed to insure glass and epoxy disks are not mixed
    '           as it is now possible to have more than one type of disk
    If GlassDiskFound = True And EpoxyDiskFound = True Then
        MsgBox "You can not mix Glass and Epoxy collared disks in the same arrester!"
        bBadDiskTypeFound = True
    End If
    
    checkDiskTypes = Not (bBadDiskTypeFound)
    
End Function

Private Sub btnNext_Click()
    'user is accepting the disks as they stand.
    'copy the disks to consumed disks and show the report.
    Dim bResult As Boolean
    
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    bResult = myGetDiskData
    If bResult = False Then
        ' get disk data failed
        Exit Sub
    End If
    
    If adoCopyDiskData = False Then
        grdDiskNum.SetFocus
        Exit Sub
    End If

    frmUnitReport.Show
    frmUnitReport.ZOrder 0
    resizeForm
    Unload Me
End Sub

Private Sub Form_Load()
    Call setDefaultLocation(frmDiskQuantities)
    Call setCaption("Provide disk quantities and categories")
    Call setProgress(3)
    Call moveForm
    
    Init = True
    grdQuantities.ColAlignment(0) = 3
    grdQuantities.ColWidth(0) = 650
    grdQuantities.ColAlignment(1) = 3
    grdQuantities.ColWidth(1) = 950
    grdQuantities.ColAlignment(2) = 3
    grdQuantities.ColWidth(2) = 950
    grdQuantities.TextMatrix(0, 0) = ""
    grdQuantities.TextMatrix(0, 1) = "Quantity"
    grdQuantities.TextMatrix(0, 2) = "Category"
    grdQuantities.TextMatrix(1, 0) = "1"
    grdQuantities.TextMatrix(2, 0) = "2"
    grdQuantities.TextMatrix(3, 0) = "3"
    grdQuantities.TextMatrix(4, 0) = "4"
    grdQuantities.TextMatrix(5, 0) = "5"
    grdQuantities.TextMatrix(6, 0) = "6"
    grdQuantities.Col = 0: grdQuantities.Row = 0
    grdQuantities.CellFontBold = True
    grdQuantities.Col = 1: grdQuantities.Row = 0
    grdQuantities.CellFontBold = True
    grdQuantities.Col = 2: grdQuantities.Row = 0
    grdQuantities.CellFontBold = True
    grdQuantities.Col = 1: grdQuantities.Row = 1
    Init = False
End Sub
Private Sub Form_Activate()
    grdQuantities.Col = 2: grdQuantities.Col = 1
    grdQuantities.Row = 2: grdQuantities.Row = 1
    grdQuantities_EnterCell
    txtDiskType.Text = globCurrentUnit.DiskType
' DLimburg Change 2009_1 11Feb09 added code to cover the use of a 5kV disks
    If Len(globCurrentUnit.DiskType2) > 0 Then
        txtDiskType2.Visible = True
        txtOr.Visible = True
        txtDiskType2.Text = globCurrentUnit.DiskType2
    Else
        txtDiskType2.Visible = False
        txtOr.Visible = False
    End If
End Sub

Private Sub btnVerify_Click()
    '----------------------------------------------------------------------------------------
    '   Created:    April 03    By: Jeff P
    '   Purpose:    to verify number of disks and categories entered
    '                    are valid before individual disk number entry.
    '   Rev'd:      05/02/03  By: Ray Raf
    '   Note:
    '   Rev'd:      10/19/05  By: D Limburg
    '   Note:       catagories of 100, 110 & 120 need to be mulpilied by 10 for this calculation
    '               locate by searching on "19Oct05"
    '----------------------------------------------------------------------------------------
    Dim X As Integer
    Dim lngSumCats As Long, lngCatsMin As Long, lngCatsMax As Long
    Dim strQuanTable As String
    Dim strSQL As String
    Dim rstCats As New ADODB.Recordset
    Dim strMsg As String
    Dim bResult As Boolean
    Dim strFamily As String
    Dim lngDiskCat As Long
    'every time a button is pressed, a check is made to see if the
    'login has expired
    bResult = CheckForValidLogin
    If bResult = False Then
        'login expired and user didn't login in
        Exit Sub
    End If
    
    ' clear highlites from quantities grid
    SetNormal grdQuantities
    
    If btnVerify.Caption = "&Done" Then

        m_NumDisksEntered = 0
        lngSumCats = 0
        'sum up the total of all disks to be used in each of the categories
' DLimburg Change 2009_1 changed quanity to 6 from 5
        For X = 1 To 6
            With grdQuantities
                If grdQuantities.TextMatrix(X, 1) <> "" And grdQuantities.TextMatrix(X, 2) <> "" Then
                    m_NumDisksEntered = m_NumDisksEntered + CInt(.TextMatrix(X, 1))
                    'mult. col 1 (num disks) time col 2 (Category -always a number)
                    ' *** Start of change 19Oct05 ***********
                    
                    lngDiskCat = CLng(.TextMatrix(X, 2))
                    

                    If Left(grdDiskNum.TextMatrix(X, 6), 4) = "AT52" Then
                    'Dlimburg Change 2009_2 removed case 200 to 530 test in select statement
                    'to if statement which tests on disk type rather that category range
                    'Testing on cat's with a min does not cover all cat's for the 5kV Disk
                    'and runs into conflict with othe disks
                        lngDiskCat = lngDiskCat + 1000
                    Else
                        Select Case lngDiskCat
                        ' catagories of 100, 110 & 120 need to be mulpilied by 10 for this calculation
                        ' DLimburg 10/19/05
                            Case 100
                                lngDiskCat = 1000
                           Case 110
                               lngDiskCat = 1000
                           Case 120
                               lngDiskCat = 1000
                         '   Case 200 To 530     ' DLimburg Change 2009_1 added this case stament to cover 5kV disk highs and lows
                          '      lngDiskCat = lngDiskCat + 1000
                         End Select
                    End If


                    lngSumCats = lngSumCats + (CLng(.TextMatrix(X, 1)) * lngDiskCat)
                    ' *** End of change 19Oct05 ***********
                End If
                
            End With
        Next
        
        If m_NumDisksEntered = 0 Then
            MsgBox "You must enter disk quantities and categories.", vbInformation, "Error"
            grdQuantities.SetFocus
            Exit Sub
        End If
        
        'check the number of disks in all categories against what the user said was going to be used.
        If m_NumDisksEntered <> globNumDisksToUse Then
            MsgBox "The total of all the disk quantities must be equal to " & globNumDisksToUse & ". (your previous entry)", vbInformation, "Error"
            grdQuantities.SetFocus
            Exit Sub
        End If
        
        'check to make sure that the sum of the categories is within the allowed range for the catalog number
        Select Case globCurrentWO.strType
            Case 1: 'porc
                strQuanTable = "DiskQuanPorc_Rev1"   ' changed table name from DiskQuanPorc to DiskQuanPorc_Rev1 - DLimburg16Feb2011
                strFamily = globCurrentUnit.Family
            Case 2: 'poly
' DLimburg Change 2009_1 changed SQL table to DiskQuanPoly_Rev1 from DiskQuanPoly
'                strQuanTable = "DiskQuanPoly_Rev1"
' DLimburg Change 2010_1 3Feb2010 Modified and added code to allow for the 3 and 5 calcined disks
                strQuanTable = "DiskQuanPoly_Rev2"
                strFamily = globCurrentUnit.Family
        End Select
        
        'get values to compare to for this unit
        strSQL = "SELECT * FROM " & strQuanTable & " WHERE DiskType='" & globCurrentUnit.DiskType & "' "
        strSQL = strSQL & "AND Design_MCOV='" & globCurrentUnit.UnitMCOV & "' "
        strSQL = strSQL & "AND Family = '" & strFamily & "' AND DiskQuantity = '" & globNumDisksToUse & "' "
        strSQL = strSQL & "AND Family = '" & strFamily & "'"

' DLimburg Change 2009_4 (November 30,2009)change query to test for number of 3kV and 5kV disks used as opposed to total number of disks
        If globCurrentWO.strType = 2 Then
            strSQL = strSQL & " AND DiskQuantity_3kV = " & globCurrentUnit.NumOf3kVDisks & " AND DiskQuantity_5kV = " & globCurrentUnit.NumOf5kVDisks
        End If
        
        Set rstCats = dbconSetupData.Execute(strSQL)
        If rstCats.EOF Then
            strMsg = "The Cum_V10CAT_Min and Cum_V10CAT_Max cannot be retrieved for " & vbCrLf
            strMsg = strMsg & "Disk Type = " & globCurrentUnit.DiskType & " and MCOV = " & globCurrentUnit.UnitMCOV & "."
            MsgBox strMsg, vbInformation, "Error"
            ErrorLogger "btnVeryfy", strMsg      'this is a data error. log it so someone can follow up on it
            grdQuantities.SetFocus
            Exit Sub
        End If
        
        lngCatsMin = CLng(rstCats!Cum_V10CAT_Min)
        lngCatsMax = CLng(rstCats!Cum_V10CAT_Max)
        'check to make sure the total of the categories is within the limits
        If lngSumCats < lngCatsMin Or lngSumCats > lngCatsMax Then
            strMsg = "The Category sum of " & lngSumCats & " must be between " & lngCatsMin & " and " & lngCatsMax & "."
            MsgBox strMsg, vbInformation, "Error"
            grdQuantities.SetFocus
            Exit Sub
        End If
        
        'Debug.Print "Min=" & lngCatsMin & ", Sum=" & lngSumCats & ", Max=" & lngCatsMax
        
        If rstCats.State > 0 Then
            rstCats.Close
        End If
        
        Init = True
        lblAvg10kACat.Caption = "?"
        lblV1mA.Caption = "?"
        lblV10kA.Caption = "?"
        lblV10kACAT.Caption = "?"
        
        'don't need to clear grid, just re-size it if neccessary
        'grdDiskNum.Clear
        
        grdDiskNum.Rows = m_NumDisksEntered + 1
        
        'set header rows
        grdDiskNum.TextMatrix(0, 1) = "Batch"
        grdDiskNum.TextMatrix(0, 2) = "Disk Number"
        grdDiskNum.TextMatrix(0, 3) = "Vref"
        grdDiskNum.TextMatrix(0, 4) = "Act V10kA"
        grdDiskNum.TextMatrix(0, 5) = "Category"
        grdDiskNum.TextMatrix(0, 6) = "Disk Type"
        
        'set column widths and header cell properties
        grdDiskNum.ColWidth(0) = 400
        grdDiskNum.ColAlignment(0) = 3
        For X = 1 To 6
            grdDiskNum.ColAlignment(X) = 3
            Select Case X
                Case Is > 2
                    grdDiskNum.ColWidth(X) = 1100
                Case Else
                    grdDiskNum.ColWidth(X) = 1340
            End Select
            grdDiskNum.Row = 0
            grdDiskNum.Col = X
            grdDiskNum.CellFontBold = True
        Next
        
        'label each row
        For X = 1 To m_NumDisksEntered
            grdDiskNum.TextMatrix(X, 0) = X
        Next
        
        grdDiskNum.Col = 1: grdDiskNum.Row = 1
        Init = False
        
        btnVerify.Caption = "&RE-ENTER"
        'lock-out change of categories quantities while entering disks
        grdQuantities.Enabled = False
        fraDiskNum.Visible = True
        
        btnNext.Enabled = True
        
        'position cursor for 1st disk data entry
        grdDiskNum.Col = 2: grdDiskNum.Col = 1
        grdDiskNum.Row = 1
        grdDiskNum_EnterCell
        grdDiskNum.SetFocus
    Else
        'btnVerify caption not equal to "&Done"
        fraDiskNum.Visible = False
        btnNext.Enabled = False
        btnVerify.Caption = "&Done"
        grdQuantities.Enabled = True
    End If
End Sub

Private Sub grdDiskNum_EnterCell()
    With grdDiskNum
        If Not Init Then
            
            'not setting up the grid, user is entering data
            If .Col = 1 And .Text = "" Then
                'puts the last batch number in a blank batch number column.
                .Text = strLastBatch
            End If
            If .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 6 Then
                'moved to one of the empty (for now) columns of the current row.
                '.Move user to 1st column of next row
                .Col = 1
                If .Row > 4 Then
                    'scroll grid
                    .TopRow = .Row - 3
                End If
                If .Row < .Rows - 1 Then
                    'move to next row
                    .Row = .Row + 1
                End If
                'now move to the disk number
                .Col = 2
            End If
            SetHighlight grdDiskNum
            m_AwaitingFirstChar = True
        End If
    End With
End Sub

Private Sub grdDiskNum_KeyPress(KeyAscii As Integer)
    flexInput grdDiskNum, KeyAscii
End Sub

Private Sub grdDiskNum_LeaveCell()
    '----------------------------------------------------------------------------------------
    '   Rev'd:      09/21/05  By: Ray Raf
    '   Note:       modified routine to parse cell data if it looked like a barcode
    '----------------------------------------------------------------------------------------
    Dim bd As BarcodeData
    Dim DiskBatch As String
    Dim DiskNum As Integer
    
    If Not Init Then
        SetNormal grdDiskNum
        
        If m_PosibleBarcode = True Then
            'parse cell contents to see if it really is a barcode
            bd = ParseBarcode(grdDiskNum.Text)
            'make sure it was a valid barcode
            If bd.BarcodeType = "D" Then
                'is a valid disk barcode type
                ' dlimburg 12/03/07 added test to verify disk has not been used in a FACTS tube.
                DiskBatch = Trim(bd.DiskData.strBatchNumber)
                DiskNum = Trim(bd.DiskData.iDiskNumber)
                With grdDiskNum
                    .TextMatrix(.Row, 1) = Trim(bd.DiskData.strBatchNumber)
                    strLastBatch = bd.DiskData.strBatchNumber
                    .TextMatrix(.Row, 2) = Trim(bd.DiskData.iDiskNumber)
                End With
            End If
            m_PosibleBarcode = False
        End If
        
        If grdDiskNum.Col = 1 Then
            strLastBatch = grdDiskNum.Text
        End If
    End If
    'reset posible barcode flag before moving to next cell
    m_PosibleBarcode = False
    
End Sub

Private Sub grdQuantities_EnterCell()
    If Not Init Then
        SetHighlight grdQuantities
        m_AwaitingFirstChar = True
    End If
End Sub

Private Sub grdQuantities_LeaveCell()
    '   Rev'd:      06//21/06   By: Ray Raf
    '   Note:       modified to handle 4 digit categrories
    If Not Init Then
        SetNormal grdQuantities
        If grdQuantities.Col = 1 Then
            If Not IsNumeric(grdQuantities.Text) Then
                grdQuantities.Text = ""
            End If
        Else
            grdQuantities.Text = Mid(grdQuantities.Text, 1, 4)  'modified 06/21/06
        End If
    End If
End Sub

Private Sub grdQuantities_KeyPress(KeyAscii As Integer)
    flexInput grdQuantities, KeyAscii
End Sub

Sub flexInput(tmpFlex As MSFlexGrid, KeyAscii As Integer)
    'first update access timer
    globLastActivity = Timer
        
    With tmpFlex
        Select Case KeyAscii
            Case vbKeyReturn, vbKeyTab
                'what to do if return or tab key is pressed
                If .Col + 1 <= .Cols - 1 Then
                    'not last col, move to next column
                    If m_PosibleBarcode = True Then
                        If .Col = 1 Or .Col = 2 Then
                            .Col = 3
                        Else
                            .Col = .Col + 1
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                    'Debug.Print "Case 1"
                ElseIf .Row + 1 <= .Rows - 1 Then
                    If .Row > 4 Then
                        'this will scroll the view window
                        .TopRow = .Row - 4
                    End If
                    .Row = .Row + 1
                    .Col = 1
                    'Debug.Print "Case 2"
                Else
                    'last row, last col, go back to beginning
                    .Row = 1
                    .Col = 1
                    'Debug.Print "Case 3"
                End If
            Case vbKeyBack
                'delete the last character of the string in the cell
                If Trim(.Text) <> "" Then
                  .Text = Mid(.Text, 1, Len(.Text) - 1)
                    ' reset the awaiting first char flag
                    m_AwaitingFirstChar = False
                End If
            Case 48 To 57
                'numeric values are always allowed
                    If m_AwaitingFirstChar Then
                        .Text = Chr(KeyAscii)
                        ' reset the awaiting first char flag
                        m_AwaitingFirstChar = False
                    Else
                        .Text = .Text & Chr(KeyAscii)
                    End If
                    'Debug.Print "Case 4"
            Case Is < 32
            Case Else
                'probably a letter
                If .Col = 0 Or .Row = 0 Then
                    'header row or column.  don't change
                    Exit Sub
                Else
                    'check to see if char is allowed
                    If m_PosibleBarcode = True Then
                        'we know that the data is possibly a barcode.  Accept all chars
                        .Text = .Text & Chr(KeyAscii)
                    Else
                        'if in column 1 or 2 on lower grid, check to see if barcode was possibly scanned.
                        Select Case .Col
                            Case 1, 2
                                'barcode only valid in lower grid
                                If LCase(.Name) = "grddisknum" Then
                                     If m_AwaitingFirstChar = True Then
                                        'check to see if char is "D", the 1st char of a disk barcode.
                                        If KeyAscii = 68 Then
                                            'highly probable that disk barcode was scanned.
                                            .Text = Chr(KeyAscii)
                                            m_AwaitingFirstChar = False
                                            m_PosibleBarcode = True
                                        Else
                                            'wasn't barcode but any char is valid for 1st char of col 1
                                            If .Col = 1 Then
                                                .Text = Chr(KeyAscii)
                                                ' reset the awaiting first char flag
                                                m_AwaitingFirstChar = False
                                            End If
                                        End If  'end of check for "D"
                                    End If  'if not 1st char and not m_posiblebarcode = true then letter not valid.
                                End If  '   end of check for lower grid
                        End Select
                    End If
                End If
        End Select
            
        
    End With
        
End Sub


Sub SetNormal(grdObject As MSFlexGrid)
    '----------------------------------------------------------------------------------------------
    '   Created:   02/28/03    By: Ray Raf
    '   Purpose:   shows the cell in it's normal, unselected state
    '----------------------------------------------------------------------------------------------
    grdObject.CellBackColor = vbWhite
    grdObject.CellForeColor = vbBlack
    
End Sub

Sub SetHighlight(grdObject As MSFlexGrid)
    '----------------------------------------------------------------------------------------------
    '   Created:   02/28/03    By: Ray Raf
    '   Purpose:   Shows the cell in a normal highlighted state
    '----------------------------------------------------------------------------------------------
    grdObject.CellBackColor = vbBlue
    grdObject.CellForeColor = vbWhite
End Sub

