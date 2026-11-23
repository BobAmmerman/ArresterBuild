VERSION 5.00
Begin VB.Form frmPrintOrient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Orientation"
   ClientHeight    =   1935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAbort 
      Left            =   4050
      Top             =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Page Orientation"
      Height          =   1140
      Left            =   375
      TabIndex        =   2
      Top             =   225
      Width           =   2490
      Begin VB.OptionButton optOrient 
         Caption         =   "Landscape"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   675
         Width           =   1140
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "Portrait"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   375
         Width           =   915
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3255
      TabIndex        =   1
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3255
      TabIndex        =   0
      Top             =   345
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintOrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public iResponse As Integer
Public iOldOrient As Integer


Private Sub CancelButton_Click()
    
    'let calling program know to cancel print
    iResponse = vbCancel
    Me.Hide
    
End Sub

Private Sub Form_Load()
    
    iOldOrient = Printer.Orientation
    optOrient(iOldOrient - 1).Value = True
    
    'set default response to cancel
    iResponse = vbCancel
    
    
End Sub

Private Sub OKButton_Click()
    
    If optOrient(0).Value = True Then
        Printer.Orientation = 1     'portrait
    Else
        Printer.Orientation = 2     'landscape
    End If
    
    'let calling routine know to proceed
    iResponse = vbOK
    
    Me.Hide
    
End Sub

Private Sub tmrAbort_Timer()

    'Cancel Print Because window was open too long.
    Display_Message "Print Window Closed because of excess time "
    iResponse = vbCancel
    Me.Hide

End Sub
