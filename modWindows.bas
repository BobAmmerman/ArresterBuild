Attribute VB_Name = "modWindow"



Global globCurrentUnit As Unit
Global globCurrentWO As WorkOrder
Global globHousingCode As String
Global globFormMain As frmMain
Global globNumDisksToUse As Integer
Global globButtonResult As Integer
Global globAutoPrint As Integer
Global globGandESubs As Integer


'--name of the current form being displayed
Global curFormName As String

Sub Main()
    Dim bResult As Boolean
    
    globCurrentAccessLevel = -1
    frmSplash.Show vbModal
    
    'initDatabase
    bResult = ConnectDatabases
    
    globAutoPrint = GetSet(AppName:=App.Title, Section:="Settings", _
                                                        Key:="AutoPrint", Default:="0")
    'added Gand E 06/21/06
    globGandESubs = GetSet(AppName:=App.Title, Section:="Settings", _
                                                        Key:="GandESubs", Default:="0")
                                                        
    
    globCurrentWO.strType = -1
    
    Set fMainForm = New frmMain
    Load fMainForm
    
    fMainForm.Show
    
    
End Sub

Public Sub SetAccess()
    '------------------------------------------------------------------------------
    '   Created:    01/30/04    By: Ray Raf
    '   Purpose:    This routine will enable and/or disable program
    '                     functions based on access level.
    '------------------------------------------------------------------------------
    
    If globCurrentAccessLevel = 1 Then
        'current user has admin privleges
        globFormMain.mnuFileUnbuild.Enabled = True
        globFormMain.mnuViewOptions.Enabled = True
        globFormMain.mnuFilePrintMultiple.Enabled = True
        globFormMain.mnuFileDeleteWO.Enabled = True
    Else
        globFormMain.mnuFileUnbuild.Enabled = False
        globFormMain.mnuViewOptions.Enabled = False
        globFormMain.mnuFileDeleteWO.Enabled = False
        globFormMain.mnuFilePrintMultiple.Enabled = True
        
    End If
End Sub

'--------------------------------------------------------
'  the following is for the wizard-based process
'--------------------------------------------------------

'--set the location of the form
Public Sub setDefaultLocation(frm As Form)
    curFormName = frm.Name
    frm.Left = frmTemplate.Left + 2670
    frm.Top = frmTemplate.Top + 860
End Sub

'--move the form (whenever the parent form is moving)
Public Sub moveForm()
    Forms(getFormIndex).Left = frmTemplate.Left + 2670
    Forms(getFormIndex).Top = frmTemplate.Top + 860
End Sub

Public Sub CancelButton()
    '------------------------------------------------------------------------------
    '   Created:    06/04/03    By: Ray Raf
    '   Purpose:    When the cancel button is clicked on any form, this routine will reset
    '                    the system so that the WO number entry form is displayed and all
    '                    data is set as if we were just starting.  This will also be used when an
    '                    arrester is completed to reset the prog. for the next arrester.
    '------------------------------------------------------------------------------
    Dim BlankWO As WorkOrder
    Dim BlankUnit As Unit
    
    If curFormName <> "frmWorkOrder" Then
        'unload current screen
        Unload Forms(getFormIndex)
        Load frmWorkOrder
    End If
    
    'clean out existing data
    globCurrentUnit = BlankUnit
    globCurrentWO = BlankWO
    
    frmWorkOrder.txtWorkOrder.Text = globCurrentWO.strWO
    resizeForm

End Sub

Public Sub resizeForm()
    '------------------------------------------------------------------------------
    '   Created:    ????    By: Jeff P
    '   Purpose:    resize the current form (whenever the parent form is resizing)
    '   Returns:    the index number for the form or -1 if form is not found
    '   Rev'd:      04/15/03    By: Ray Raf
    '   Note:       added check of form index returned by getFormIndex
    '------------------------------------------------------------------------------
    Dim frmCurrent As Form
    Dim iFormNum As Integer
    
    iFormNum = getFormIndex
    If iFormNum < 0 Then
        'the current form could not be found for some reason.
        MsgBox "resizeForm Current form: " & curFormName & " Could not be found!"
        ErrorLogger "resizeForm", "Current form: " & curFormName & " Could not be found!"
        Exit Sub
    End If
    
    On Error Resume Next
    Forms(iFormNum).Width = frmTemplate.Width - Forms(iFormNum).Left - 40
    Forms(iFormNum).Height = frmTemplate.Height - Forms(iFormNum).Top - 40
    Forms(iFormNum).btnNext.Left = Forms(iFormNum).Width - Forms(iFormNum).btnNext.Width - 40
    Forms(iFormNum).btnNext.Top = Forms(iFormNum).Height - Forms(iFormNum).btnNext.Height - 40
    Forms(iFormNum).btnBack.Left = Forms(iFormNum).btnNext.Left - 1100
    Forms(iFormNum).btnBack.Top = Forms(iFormNum).btnNext.Top
    Forms(iFormNum).btnCancel.Left = Forms(iFormNum).btnBack.Left - 1600
    Forms(iFormNum).btnCancel.Top = Forms(iFormNum).btnBack.Top
    
End Sub
'--set the main caption of the form
Public Sub setCaption(varCaption)
    frmTemplate.lblStatus.Caption = varCaption
End Sub

'--highlight the process bar on the wizard (the index of the label is passed)
Public Sub setProgress(iStep As Integer)
    Dim iLoop As Integer
    
    For iLoop = 0 To frmTemplate.lblProgress.UBound
        With frmTemplate.lblProgress(iLoop)
            If iLoop = iStep Then
                .ForeColor = &HFFFFFF
                .Tag = 1        'mark as active step
            Else
                .ForeColor = &H808080
                .Tag = 0        'mark as not the active step
            End If
        End With
    Next
End Sub

Public Function getFormIndex() As Integer
    '------------------------------------------------------------------------------
    '   Created:    ????    By: Jeff P
    '   Purpose:    determine which form is being displayed
    '   Returns:    the index number for the form or -1 if form is not found
    '   Rev'd:  04/15/03    By: Ray Raf
    '   Note:   added explicit return type for function
    '------------------------------------------------------------------------------
    Dim iFormLoop As Integer
    For iFormLoop = 0 To Forms.Count - 1
        If curFormName = Forms(iFormLoop).Name Then
            getFormIndex = iFormLoop
            Exit Function
        End If
    Next
    getFormIndex = -1

End Function

