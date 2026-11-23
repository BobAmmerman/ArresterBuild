VERSION 5.00
Begin VB.Form frmManualReport 
   Caption         =   "Select Unit for Report"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   4050
      TabIndex        =   6
      Top             =   1500
      Width           =   1290
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   465
      Left            =   2212
      TabIndex        =   5
      Top             =   1500
      Width           =   1290
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   465
      Left            =   375
      TabIndex        =   4
      Top             =   1500
      Width           =   1290
   End
   Begin VB.TextBox txtUnit 
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Top             =   825
      Width           =   1740
   End
   Begin VB.TextBox txtSN 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   300
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Unit (A-F):"
      Height          =   195
      Left            =   705
      TabIndex        =   1
      Top             =   825
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number:"
      Height          =   195
      Left            =   375
      TabIndex        =   0
      Top             =   375
      Width           =   1035
   End
End
Attribute VB_Name = "frmManualReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ReportWO As WorkOrder
Private m_ReportUnit As Unit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    '------------------------------------------------------------------------------------------------
    'Created:   12/18/03    By:Ray Rafalowski
    'Purpose:   To print the unit report based on the unit information
    '                that the user entered on this form.
    '-------------------------------------------------------------------------------------------------
    Dim strSN As String
    Dim strUnit As String
    Dim iUnitNum As Integer
    Dim strMsg As String
    Dim bResult As Boolean
    
    strSN = Trim(txtSN.Text)
    
    'make sure it is a valid serial number
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
    If Len(strSN) <> 11 Then
        MsgBox "The serial number must be exactly 11 digits"
        Exit Sub
    End If
    
    strUnit = Trim(txtUnit.Text)
    
    If Len(strUnit) <> 1 Then
        MsgBox "The Unit must be one character (A-F) or (1-6)"
        Exit Sub
    End If
    
    'check to make sure that the unit is valid
    Select Case UCase(strUnit)
        Case "A" To "F"
            iUnitNum = Asc(UCase(strUnit)) - 64
        Case "1" To "6"
            iUnitNum = CInt(strUnit)
        Case Else
            MsgBox "The Unit must be one character (A-F) or (1-6)"
            Exit Sub
    End Select
    
    'If porcelain arrester, test to see if build uses calcined disks
    If Test4CalcinedDiskInPorc(strSN, iUnitNum) = False Then
        MsgBox "Serial number for given unit not found."
        Exit Sub
    End If
    
    'load unit data
    m_ReportUnit.SN = strSN
    m_ReportUnit.UnitNumber = iUnitNum
    bResult = LoadUnitInfo(m_ReportWO, m_ReportUnit)
    If bResult = False Then
        'could not retreive any/all info for the arrester.  Can't print report
        strMsg = "Not all data could be retreived for the entered unit." & vbCrLf
        strMsg = strMsg & "Report can not be printed!"
        MsgBox strMsg
        Exit Sub
    End If
    
    globCurrentWO = m_ReportWO
    globCurrentUnit = m_ReportUnit
    Set frmUnitReport.dbconBuilt = dbconResultsData
    SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFilename", "c:\Data\" & m_ReportUnit.SN & ".pdf"
    frmUnitReport.PrintUnitReport Printer
    Printer.EndDoc

    MsgBox "Report was Printed!"
End Sub

Private Sub cmdView_Click()
    '------------------------------------------------------------------------------------------------
    'Created:   05/23/03    By:Ray Rafalowski
    'Purpose:   To display the unit report form based on the unit information
    '                that the user entered on this form.
    '-------------------------------------------------------------------------------------------------
    Dim strSN As String
    Dim strUnit As String
    Dim iUnitNum As Integer
    Dim strMsg As String
    Dim bResult As Boolean
    
    strSN = Trim(txtSN.Text)
    
    'make sure it is a valid serial number
    '10/4/05 serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    'DLimburg
    If Len(strSN) <> 11 Then
        MsgBox "The serial number must be exactly 11 digits"
        Exit Sub
    End If
    
    strUnit = Trim(txtUnit.Text)
    
    If Len(strUnit) <> 1 Then
        MsgBox "The Unit must be one character (A-F) or (1-6)"
        Exit Sub
    End If
    
    'check to make sure that the unit is valid
    Select Case UCase(strUnit)
        Case "A" To "F"
            iUnitNum = Asc(UCase(strUnit)) - 64
        Case "1" To "6"
            iUnitNum = CInt(strUnit)
        Case Else
            MsgBox "The Unit must be one character (A-F) or (1-6)"
            Exit Sub
    End Select
    
    'load unit data
    m_ReportUnit.SN = strSN
    m_ReportUnit.UnitNumber = iUnitNum
    bResult = LoadUnitInfo(m_ReportWO, m_ReportUnit)
    If bResult = False Then
        'could not retreive any/all info for the arrester.  Can't print report
        strMsg = "Not all data could be retreived for the entered unit." & vbCrLf
        strMsg = strMsg & "Report can not be printed!"
        MsgBox strMsg
        Exit Sub
    End If
    
    globCurrentWO = m_ReportWO
    globCurrentUnit = m_ReportUnit
    
    'now show report screen
    '--a report is not supposed to be able to be printed unless frmWorkOrder is the current form
    Me.Hide
    frmUnitReport.Show
    frmUnitReport.ZOrder 0
    resizeForm
    Unload frmWorkOrder
    Unload Me
    
    
End Sub

