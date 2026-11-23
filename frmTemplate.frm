VERSION 5.00
Begin VB.Form frmTemplate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9528
   ClientLeft      =   3024
   ClientTop       =   3024
   ClientWidth     =   14220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9528
   ScaleWidth      =   14220
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   11
      Tag             =   "0"
      Top             =   2115
      Width           =   2100
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Disk Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   10
      Tag             =   "0"
      Top             =   1830
      Width           =   2100
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number of Disks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Tag             =   "0"
      Top             =   1545
      Width           =   2100
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Unit and SN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Tag             =   "0"
      Top             =   1260
      Width           =   2100
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2652
      TabIndex        =   7
      Top             =   396
      Width           =   10200
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cooper Build Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H8000000D&
      Height          =   336
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12840
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Work Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Tag             =   "0"
      Top             =   975
      Width           =   2100
   End
   Begin VB.Label lblDecor 
      BackColor       =   &H00FFFFFF&
      Height          =   468
      Index           =   2
      Left            =   2616
      TabIndex        =   4
      Top             =   336
      Width           =   10188
   End
   Begin VB.Label lblDecor 
      BackColor       =   &H00404040&
      Height          =   7356
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   792
      Width           =   2616
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   465
      Width           =   1275
   End
   Begin VB.Label lblDecor 
      BackColor       =   &H00C0C0C0&
      Height          =   450
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   345
      Width           =   2610
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xpostion As Long, ypostion As Long

Private Sub Form_Activate()
    Forms(getFormIndex).Show
    Forms(getFormIndex).ZOrder 0
End Sub

Private Sub Form_Load()
    If curFormName = "" Then
        Load frmWorkOrder
    End If
End Sub

Public Sub Form_Resize()
    lblDecor(1).Height = Me.ScaleHeight
    lblDecor(2).Width = Me.ScaleWidth
    lblLabels(1).Width = Me.ScaleWidth
    Call resizeForm
End Sub

Private Sub lblLabels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 1, 2
            xpostion = X
            ypostion = Y
    End Select
End Sub

Private Sub lblLabels_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 1, 2
            If Button = vbLeftButton Then
                Me.Left = Me.Left - (xpostion - X)
                Me.Top = Me.Top - (ypostion - Y)
                Call moveForm
            End If
    End Select
End Sub
