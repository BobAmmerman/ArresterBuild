VERSION 5.00
Begin VB.Form frmEntry 
   Caption         =   "Create Disks"
   ClientHeight    =   4068
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4008
   LinkTopic       =   "Form1"
   ScaleHeight     =   4068
   ScaleWidth      =   4008
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVref 
      Height          =   312
      Left            =   1524
      TabIndex        =   3
      Top             =   1896
      Width           =   1464
   End
   Begin VB.TextBox txtType 
      Height          =   312
      Left            =   1536
      TabIndex        =   2
      Top             =   1446
      Width           =   1464
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Disks"
      Height          =   444
      Left            =   1368
      TabIndex        =   6
      Top             =   3240
      Width           =   1056
   End
   Begin VB.TextBox txtDiskCount 
      Height          =   312
      Left            =   1524
      TabIndex        =   5
      Top             =   2808
      Width           =   1464
   End
   Begin VB.TextBox txtCategory 
      Height          =   312
      Left            =   1524
      TabIndex        =   4
      Top             =   2352
      Width           =   1464
   End
   Begin VB.TextBox txtStartNum 
      Height          =   312
      Left            =   1536
      TabIndex        =   1
      Top             =   993
      Width           =   1464
   End
   Begin VB.TextBox txtBatch 
      Height          =   312
      Left            =   1536
      TabIndex        =   0
      Top             =   540
      Width           =   1464
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vref"
      Height          =   192
      Left            =   1008
      TabIndex        =   12
      Top             =   1908
      Width           =   288
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disk Type"
      Height          =   192
      Left            =   564
      TabIndex        =   11
      Top             =   1452
      Width           =   744
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Disk Count"
      Height          =   192
      Index           =   3
      Left            =   528
      TabIndex        =   10
      Top             =   2868
      Width           =   768
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Category"
      Height          =   192
      Index           =   2
      Left            =   636
      TabIndex        =   9
      Top             =   2424
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "1st disk number"
      Height          =   192
      Left            =   192
      TabIndex        =   8
      Top             =   1020
      Width           =   1116
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Batch"
      Height          =   192
      Index           =   0
      Left            =   900
      TabIndex        =   7
      Top             =   588
      Width           =   408
   End
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreate_Click()
    Dim strSQL As String
    Dim iLoop As Integer
    Dim V10kA As Single, v1ma As Single, vref As Single
    
    V10kA = txtCategory.Text / 100
    v1ma = V10kA / 1.72
    vref = CSng(txtVref.Text)
    
    For iLoop = CInt(txtStartNum.Text) To CInt(txtStartNum.Text) + CInt(txtDiskCount.Text) - 1
        strSQL = "INSERT INTO AvailableDisks VALUES('" & _
            txtBatch & "'," & iLoop & _
            ",'" & UCase(txtType.Text) & "','" & UCase(txtType.Text) & _
             "','" & Now() & "',444,444,444,444,444,444," & _
            v1ma & ",19.96," & vref & ",0.1153,100," & V10kA & _
            ",1000,0," & txtCategory.Text & "," & iLoop & ") ;"
        dbconSetupData.Execute (strSQL)
    Next iLoop
    
    MsgBox "Disk Creation Complete"
End Sub

Private Sub Form_Load()
ConnectDatabases
End Sub

Private Sub Label1_Click(Index As Integer)
    ConnectDatabases
End Sub
