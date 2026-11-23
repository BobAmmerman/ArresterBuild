VERSION 5.00
Begin VB.Form frmUnbuildSelection 
   Caption         =   "Select Arrester to Unbuild"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSN 
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   150
      Width           =   1740
   End
   Begin VB.TextBox txtUnit 
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      Top             =   675
      Width           =   1740
   End
   Begin VB.CommandButton cmdUnbuild 
      Caption         =   "&Unbuild"
      Height          =   465
      Left            =   225
      TabIndex        =   2
      Top             =   1350
      Width           =   1290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   465
      Left            =   3900
      TabIndex        =   3
      Top             =   1350
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number:"
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   225
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Unit (A-F):"
      Height          =   195
      Left            =   555
      TabIndex        =   4
      Top             =   675
      Width           =   705
   End
End
Attribute VB_Name = "frmUnbuildSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUnbuild_Click()
    '------------------------------------------------------------------------------------------------
    'Created:   05/23/03    By:Ray Rafalowski
    'Purpose:   To unbuild an arrester based on the unit information
    '                that the user entered on this form.
    '-------------------------------------------------------------------------------------------------
    Dim strSN As String
    Dim strUnit As String
    Dim iUnitNum As Integer
    Dim strMsg As String
    Dim bResult As Boolean
    Dim bTestSNLgth
    strSN = Trim(txtSN.Text)
    
    'make sure it is a valid serial number
    '10/4/05   DLimburg
    'serial number length test changed to 11 from 10 to
    'accommodate changes with conversion to Unity.
    '1/30/06   DLimburg
    'serial number length test changed to accept 10 or 11 digits
    'done at MRamages request to allow arrester with old 10 digit
    'serial numbers to be unbuilt
    
    If Len(strSN) = 10 Or Len(strSN) = 11 Then
        bTestSNLgth = True
    Else
        bTestSNLgth = False
    End If
    
    If bTestSNLgth = False Then
        MsgBox "The serial number must be 10 or 11 digits to unbuild"
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

    bResult = UnBuildUnit(strSN, Chr(iUnitNum + 64))
    If bResult = True Then
        MsgBox "Unit was unbuilt and disks returned to inventory!"
    End If
    
    Unload Me
    
End Sub
