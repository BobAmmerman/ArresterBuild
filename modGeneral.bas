Attribute VB_Name = "ModGeneral"
Option Explicit

Declare Function WNetGetUserA Lib "mpr" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public fMainForm As frmMain

Global globCommsIndex   As Integer
Global globFullAccess As Boolean



Public Sub GetFormIndexes()
    '*******************************************************************************************
    'Created 3/29/00    By:Ray Raf
    'Purpose:   To find the index numbers for each of the loaded forms.  This routine
    '                will change with every project so it should be here instead of TurainFunctions
    '*******************************************************************************************
    Dim iFormLoop As Integer
   
   For iFormLoop = 0 To Forms.Count - 1
        Select Case Forms(iFormLoop).Name
            Case "frmMain"
                globMainIndex = iFormLoop
            Case "frmComms"
                globCommsIndex = iFormLoop
            Case "frmTemplate"
                curFormIndex = iFormLoop
        End Select
    Next iFormLoop
End Sub

Public Sub Shutdown()
End Sub

Public Sub Display_Message(strMessage As String)
    '*******************************************************************************************
    'Created 05/26/00    By:Ray Raf
    'Purpose:   To display a message to the operator in the list box.
    '*******************************************************************************************
    Dim frmMsg As Form
    Dim LBox As ListBox
    
    'GetFormIndexes
    Set LBox = Forms(globTestResultsIndex).lstStatus
    Do While LBox.ListCount > 149
        LBox.RemoveItem 0
    Loop
    
    'with no index listed the newest item will alway go to the end of the list.
    LBox.AddItem strMessage
    LBox.TopIndex = LBox.ListCount - 1
         
End Sub

Public Sub Clear_Results()
    Dim frmMsg As Form
    Dim LBox As ListBox
    
    Set LBox = Forms(globTestResultsIndex).lstResults
    LBox.Clear
    
    Set LBox = Nothing
End Sub



