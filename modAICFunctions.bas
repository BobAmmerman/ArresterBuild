Attribute VB_Name = "modAICFunctions"
Option Explicit


'DDE link mode constants
Public Const conLinkNone = 0
Public Const conLinkAuto = 1
Public Const conLinkManual = 2
Public Const conLinkNotify = 3

Public frmMsg As Form       'a form containing a listbox by the name
                                         'of lstStatus that messages will be displayed in
                                         
Public iDBGLoop As Integer         'to be used in immediate window only!!!!!

Sub ErrorLogger(strModuleName As String, strMessage As String)
    '***********************************************************************************
    '   Created:    11/01/99    By:Ray Raf
    '   Purpose:    This procedure is use to log any error or informational
    '                    events to a text file.                     .
    '   strModuleName:  for debugging purposes the module in which the
    '                             failure has occured is included in the data
    '   Rev'd:      01/15/01    By: Ray Rafalowski
    '   Note:       changed the routine to save the log file in the application path
    '                  and start a new file each month.
    '***********************************************************************************
    Dim filenum As Integer 'file number used for log file
    Dim wrtbuf As String 'string to write to the log file
    Dim strFileName As String
    
    On Error GoTo ErrorHandler
    
    wrtbuf = Format(Date, "short date") & "  " & Format(Time, _
        "medium time") & "  " & strModuleName & "  " & _
        strMessage
    strFileName = App.Path & "\" & Format(Now(), "mmmyyyy") & ".log"
    
    filenum = FreeFile
    Open strFileName For Append As filenum
    Print #filenum, wrtbuf
    Close #filenum
    
    'Display_Message "LOGGED ERROR " & wrtbuf
    Exit Sub
    
ErrorHandler:
    Err.Clear
    
End Sub

Function FindForm(strFormToFind As String) As Integer
    '***********************************************************************************
    '   Created:    11/01/99    By:Ray Raf
    '   Purpose:    This subroutine will interogate the Forms collection to find
    '                    the index number in the collection for a form with the passed
    '                    name.                                                                          .
    '   strFormToFind:  The name of the form to find in the collection.
    '   Returns:            Either the index number for the requested form, or
    '                           zero which indicates the form was not found.
    '   Rev'd:        03/30/00    By: Ray Raf
    '   Note:        Changed return value to -1 (0 is a valid form in the collection)
    '***********************************************************************************
    Dim iLoop As Integer
    
    For iLoop = 0 To Forms.Count - 1
        If Forms(iLoop).Name = strFormToFind Then
            FindForm = iLoop
            Exit Function
        End If
    Next iLoop
        
    FindForm = -1
End Function


Public Sub Increment(iValue As Integer)
    '***********************************************************************************
    '   Created:    11/07/99    By:Ray Raf
    '   Purpose:    Adds one to an integer value
    '***********************************************************************************
    iValue = iValue + 1
End Sub

Public Sub Decrement(iValue As Integer)
    '***********************************************************************************
    '   Created:    11/07/99    By:Ray Raf
    '   Purpose:    subtracts one from an integer value
    '***********************************************************************************
    iValue = iValue - 1
End Sub

Public Sub Delay(iMilliSeconds As Integer)
    '***********************************************************************************
    '   Created:    06/05/00    By:Ray Raf
    '   Purpose:    subroutine will stay active for AT LEAST the number of
    '                    milliseconds passed to it.  It is to be used to provide
    '                    time between two actions in the calling routine.  It is
    '                    not a precise measurement, only an approximation.
    '***********************************************************************************
    Dim sglReqDelay As Single
    Dim StartTime As Single, Diff As Single
    
    sglReqDelay = iMilliSeconds / 1000
    StartTime = Timer
    Diff = 0
    
    Do While Diff < sglReqDelay
        DoEvents
        Diff = Timer - StartTime
    Loop
    
End Sub


Public Sub Display_Message(strMessage As String)
    '*******************************************************************************************
    'Created 05/26/00    By:Ray Raf
    'Purpose:   To display a message to the operator in the list box.
    '*******************************************************************************************
    Dim LBox As ListBox
    
    Set LBox = frmMsg.lstStatus
    Do While LBox.ListCount > 149
        LBox.RemoveItem 0
    Loop
    
    'with no index listed the newest item will alway go to the end of the list.
    LBox.AddItem strMessage
    LBox.TopIndex = LBox.ListCount - 1
         
    Set LBox = Nothing
End Sub

Public Sub Clear_Messages()
    '*******************************************************************************************
    'Created 06/13/00    By:Ray Raf
    'Purpose:   To clear out the message box when needed.
    '*******************************************************************************************
    
    
    frmMsg.lstStatus.Clear
    

End Sub


Public Function IsEven(ByVal iNumber As Integer) As Boolean
    '*******************************************************************************************
    'Created 09/11/00    By:Ray Raf
    'Purpose:   to check the passed value to see if it is even.
    'Inputs:    iNumber - the value to check
    'Returns:   True if the number is even, False if odd.
    '*******************************************************************************************
    Dim iValuetoCheck As Integer
    Dim iProcessedNum As Integer
    
    'Negative numbers round different so make all numbers positive
    iValuetoCheck = Abs(iNumber)
    
    iProcessedNum = Int(iValuetoCheck / 2) * 2
    
    If iValuetoCheck = iProcessedNum Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

