VERSION 5.00
Begin VB.Form frmViewErrorLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Error Log"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   6960
      Top             =   0
   End
   Begin VB.CommandButton cmdNewest 
      Caption         =   "Newest"
      Height          =   735
      Left            =   4080
      Picture         =   "frmViewErrorLog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOldest 
      Caption         =   "Oldest"
      Height          =   735
      Left            =   120
      Picture         =   "frmViewErrorLog.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurge 
      Caption         =   "&Purge old Log Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtLogData 
      Height          =   5295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   8415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextDay 
      Caption         =   "Next Day"
      Height          =   735
      Left            =   2760
      Picture         =   "frmViewErrorLog.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevDay 
      Caption         =   "Prev. Day"
      Height          =   735
      Left            =   1440
      Picture         =   "frmViewErrorLog.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLogFilePath As String
Dim mod_ActiveFile  As Integer
Dim iNumFiles As Integer
Dim strFileNames(100) As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewest_Click()
    ResetTimer
    
    mod_ActiveFile = 1
    LoadTextFile strLogFilePath & strFileNames(mod_ActiveFile), txtLogData
    
    cmdOldest.Enabled = True
    cmdPrevDay.Enabled = True
    cmdNextDay.Enabled = False
    cmdNewest.Enabled = False

End Sub

Private Sub cmdNextDay_Click()
    ResetTimer
    
    mod_ActiveFile = mod_ActiveFile - 1
    LoadTextFile strLogFilePath & strFileNames(mod_ActiveFile), txtLogData
    
    If mod_ActiveFile = 1 Then
        cmdNextDay.Enabled = False
        cmdNewest.Enabled = False
    Else
        cmdNextDay.Enabled = True
        cmdNewest.Enabled = True
    End If
    
    cmdOldest.Enabled = True
    cmdPrevDay.Enabled = True
End Sub

Private Sub cmdOldest_Click()
    ResetTimer
    
    mod_ActiveFile = iNumFiles
    LoadTextFile strLogFilePath & strFileNames(mod_ActiveFile), txtLogData
    
    cmdOldest.Enabled = False
    cmdPrevDay.Enabled = False
    cmdNextDay.Enabled = True
    cmdNewest.Enabled = True
    
End Sub

Private Sub cmdPrevDay_Click()
        ResetTimer
    
    mod_ActiveFile = mod_ActiveFile + 1
    LoadTextFile strLogFilePath & strFileNames(mod_ActiveFile), txtLogData
    
    If mod_ActiveFile = iNumFiles Then
        cmdOldest.Enabled = False
        cmdPrevDay.Enabled = False
    Else
        cmdOldest.Enabled = True
        cmdPrevDay.Enabled = True
    End If
    
    cmdNextDay.Enabled = True
    cmdNewest.Enabled = True

End Sub

Private Sub cmdPurge_Click()
    '------------------------------------------------------------------------------------------------
    '   Created:   12/20/01    By:Ray Raf
    '   Purpose:   If there are more than 10 log files, this routine will
    '                   keep the newest 10 and delete the rest.
    '------------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    Dim iAllowedFiles As Integer
    
    On Error GoTo GetOutOFHere
    
    iAllowedFiles = 10
    ResetTimer
    
    If iNumFiles > iAllowedFiles Then
        'there are enough files to delete
        'first check to see which file is open
        If mod_ActiveFile > iAllowedFiles Then
            'show last file that will be left after delete
            mod_ActiveFile = iAllowedFiles
            LoadTextFile strLogFilePath & strFileNames(mod_ActiveFile), txtLogData
            
            cmdOldest.Enabled = False
            cmdPrevDay.Enabled = False
            cmdNextDay.Enabled = True
            cmdNewest.Enabled = True
        End If
        
        'delete files and remove from list
        For iLoop = iNumFiles To (iAllowedFiles + 1) Step -1
            Kill strLogFilePath & strFileNames(iLoop)
            strFileNames(iLoop) = ""
        Next iLoop
        
        'set counter to max allowed files
        iNumFiles = iAllowedFiles
    End If
    
    Exit Sub
GetOutOFHere:
    Err.Clear
    
End Sub

Private Sub Form_Load()
    Dim strNextFile As String
    
    'the log file path is the directory that all log files should be in.
    strLogFilePath = App.Path & "\Resources\"
    
    strNextFile = Dir(strLogFilePath & "Log*")
    Do While strNextFile <> ""
        'insert name into list
        InsertFile strNextFile
        'get next file name.  Returns a blank when no more files
        strNextFile = Dir
    Loop
    
    LoadTextFile strLogFilePath & strFileNames(1), txtLogData
    mod_ActiveFile = 1
    
    cmdNextDay.Enabled = False
    cmdNewest.Enabled = False
        
    Timer1.Enabled = True
End Sub

Private Function InsertFile(strFileToInsert As String)
    '------------------------------------------------------------------------------------------------
    '   Created:   12/20/01    By:Ray Raf
    '   Purpose:   adds file name to sorted list of files.  Newest name at the
    '                   start.  Throws away oldest file if more files than
    '                   room in the array.
    '   Inputs:     strFileToInsert - the file name to insert into the list
    '------------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    Dim iNumNames As Integer
    Dim iMaxNames As Integer
    
    iMaxNames = UBound(strFileNames)
    'figure out how many file names are in the array
    If iNumFiles < UBound(strFileNames) Then
        iNumNames = iNumFiles
    Else
        iNumNames = iMaxNames
    End If
    
    If iNumNames = 0 Then
        strFileNames(1) = strFileToInsert
        iNumFiles = iNumFiles + 1
        Exit Function
    End If
    
    For iLoop = 1 To iNumNames
        If (strFileToInsert > strFileNames(iLoop)) Then
            'The new name goes in this position
            If iLoop = iMaxNames Then
                'last name in the list, don't have to shuffle
                strFileNames(iLoop) = strFileToInsert
                Exit For
            Else
                'names below this one, shuffle them down
                InsertIntoList strFileToInsert, iLoop, strFileNames
                iNumFiles = iNumFiles + 1
                Exit For
            End If
        Else
            If (iLoop = iNumNames) And (iLoop < iMaxNames) Then
                'we have reached the last name in the list but have room to
                'put more names in.
                strFileNames(iLoop + 1) = strFileToInsert
                iNumFiles = iNumFiles + 1
            End If
        End If
    Next iLoop
    
End Function

Private Sub InsertIntoList(strItem As String, iPos As Integer, ByRef strList() As String)
    '------------------------------------------------------------------------------------------------
    '   Created:   12/20/01    By:Ray Raf
    '   Purpose:   adds file name to array at specified postion.  Moves
    '                   rest of data 1 position greater in array.
    '   Inputs:     strItem - value to insert into list
    '                  iPos - position in list to insert strItem
    '                  strList() - List to insert item into
    '------------------------------------------------------------------------------------------------
    Dim iEOL As String
    Dim iLoop As Integer
    
    iEOL = UBound(strList)
    If iPos = iEOL Then
        'inserting at end of list, don't have to move anything.
        strList(iPos) = strItem
    Else
        For iLoop = iEOL To iPos + 1 Step -1
            strList(iLoop) = strList(iLoop - 1)
        Next iLoop
        strList(iPos) = strItem
    End If
    
End Sub

Public Sub LoadTextFile(strFileName As String, txtBox As TextBox)
    '------------------------------------------------------------------------------------------------
    '   Created:   12/20/01    By:Ray Raf
    '   Purpose:   Opens a text file and shows contents in a textbox.
    '   Inputs:     strFile - name of text file to load
    '                  txtBox - textbox to display data in
    '------------------------------------------------------------------------------------------------
    Dim iFileNum As Integer
    Dim strLine As String
    
    iFileNum = FreeFile
    On Error Resume Next
    Open strFileName For Input Access Read As iFileNum
    If Err.Number > 0 Then Exit Sub
    With txtBox
        .Text = ""
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, strLine
            .Text = .Text & vbCrLf & strLine
        Loop
    End With
    
    Close #iFileNum
End Sub


Private Sub Timer1_Timer()
    Unload Me
End Sub


Private Sub txtLogData_Click()
    ResetTimer
End Sub

Private Sub ResetTimer()
    Timer1.Enabled = False
    DoEvents
    Timer1.Enabled = True

End Sub
