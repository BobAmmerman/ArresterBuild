VERSION 5.00
Begin VB.Form frmComms 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUpdateRunStatus 
      Interval        =   3000
      Left            =   9240
      Top             =   2880
   End
   Begin VB.Timer tmrCommErrorRetry 
      Interval        =   8000
      Left            =   9360
      Top             =   720
   End
   Begin VB.Label lblCommError 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "There is a communications Error"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   14655
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Restart Perm"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Voltage Stabilized?"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Voltage Interlock"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Clearing Dial"
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Code:"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "V1mA Result:"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Station Enables:"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "V1ma Current:"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "V1ma Voltage:"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stations With Disks"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready for testing"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PLC Alive Flag (B3:0)"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmComms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        Option Explicit

'Object reference to RSLinx OPC interface
Private opsReadServer As New RsiOPCAuto.OPCServer
'Object reference to an OPCGroup (with events)
Private WithEvents opgGroup As RsiOPCAuto.OPCGroup 'used to continuously read data
Attribute opgGroup.VB_VarHelpID = -1

Private Sub CommP1_OnComm()
    '***********************************************************************************
    '   Created:    01/10/02    By:Ray Raf
    '   Purpose:    This routine Will handle anyComm port events.  If the event
    '                    is an error, it will display an error message.
    '***********************************************************************************
    
    Dim iWhatHappened As Integer
    Dim strMsg As String
    Dim strClean As String
    
    iWhatHappened = CommP1.CommEvent
    
    Select Case iWhatHappened
        Case comEventBreak
            strMsg = "1001 A Break signal was received."
        Case comEventFrame
            strMsg = "1004 Framing Error. The hardware detected a framing error."
        Case comEventOverrun
            strMsg = "1006 Port Overrun. A character was not read from the hardware before the next character arrived and was lost."
        Case comEventRxOver
            strMsg = "1008 Receive Buffer Overflow. There is no room in the receive buffer."
        Case comEventRxParity
            strMsg = "1009 Parity Error. The hardware detected a parity error."
        Case comEventTxFull
            strMsg = "1010 Transmit Buffer Full. The transmit buffer was full while trying to queue a character."
        Case comEventDCB
            strMsg = "1011 Unexpected error retrieving Device Control Block (DCB) for the port."
        Case comEvSend
            strMsg = "1 There are fewer than Sthreshold number of characters in the transmit buffer."
        Case comEvReceive
            strMsg = "2 Received " & globFormComm.CommP1.RThreshold & " characters."
            'set flag to indicate printer is communicating
            globPrinterOK = True
            'clear characters from port
            strClean = globFormComm.CommP1.Input
        Case comEvCTS
            strMsg = "3 Change in Clear To Send line."
        Case comEvDSR
            strMsg = "4 Change in Data Set Ready line. This event is only fired when DSR changes from 1 to 0."
        Case comEvCD
            strMsg = "5 Change in Carrier Detect line."
        Case comEvRing
            strMsg = "6 Ring detected. "
        Case comEvEOF
            strMsg = "7 End Of File (ASCII character 26) character received."
    End Select
    
    If iWhatHappened >= 1000 Then
        'it was an error, not just a regular event
        Display_Message ">>>>>>Comm Error " & strMsg
    End If
    
End Sub

Private Sub Form_Load()
    '***********************************************************************************
    '   Created:    03/30/00    By:Ray Raf
    '   Purpose:    This routine will check what buttons should be activiated
    '                    based on logged in user and start communications.
    '   Rev'd:        08/20/01   By: Ray Raf
    '   Note:         replaced fixed Comms initialization with routine that
    '                    inits comms based on which type of printer is selected.
    '***********************************************************************************
    
    'enable or disable any PB's that might be for full access use only
    If globFullAccess = True Then
        'Frame1.Visible = True
    Else
        'Frame1.Visible = False
    End If
    TurnOnSLCLinks
    
    
End Sub

Private Sub TurnOnSLCLinks()
    '***********************************************************************************
    '   Created:    11/11/02    By:Ray Raf
    '   Purpose:    This routine will setup all of the OPC Groups and items that
    '                    are needed by the program to communicate to the plc(s)/ SLC(s).
    '                    If communcations can't be
    '                    established, error messages will be displayed and a comm
    '                    retry timer started.
    '***********************************************************************************
    Dim lngResult As Long
    Dim iLoop As Integer
    Dim Frm As frmMain
    
    On Error GoTo Error_Handler
    'connect to RSLinx OPC Server on this machine
    opsReadServer.Connect "RSLinx OPC Server"
    
    ' Need to Create an OPC Group
    Set opgGroup = opsReadServer.OPCGroups.Add(App.EXEName)
    'set default state subscription to capture data changes
    opgGroup.IsSubscribed = True
    opgGroup.UpdateRate = 50
    'sample Sdditem statements.  The first shows a connection to
    'B3:0 for a 1 word read.  The 1 on the end is the client handle.
    'TesterOPC is the topic name in RSLinx.
    '   opgGroup.OPCItems.AddItem "[TesterOPC]B3:0", 1
    '                                                                               .
    'this second sample is a connection to 45 words, begining at
    'N7:0.  the 2 on the end is the client handle.
    '   opgGroup.OPCItems.AddItem "[TesterOPC]N7:0,L45", 2
    '                                                                                       .
    'To allow checking of the communication status,  all routines in
    'this form assume that the first OPC item is a register that cam be
    'written to at any time without affecting machine operation.
    opgGroup.OPCItems.AddItem "[OPCTest]B3:0", 1
    opgGroup.IsActive = True
    
    
    CheckForLiveComms opgGroup
    
        
        
    Exit Sub
Error_Handler:
    If Err.Number = 282 Then
        'RSLinx is not started
        Err.Clear
        lngResult = Shell("C:\Program Files\Rockwell Software\RSLinx\Rslinx.exe", vbMinimizedNoFocus)
        tmrCommErrorRetry.Enabled = True
        ErrorLogger "TurnOnSLCLinks", "Had to Start RSLinx"
    Else
        ErrorLogger "TurnOnSLCLinks", Err.Number & " " & Err.Description
    End If

End Sub

Private Sub opgGroup_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
    '***********************************************************************************
    '   Created:    10/11/02    By:Ray Raf
    '   Purpose:    This routine is triggered when any of the data items in the
    '                     OPC group change.  It will place the data in the proper
    '                     text boxes so that the main program logic does not have
    '                     to change because of the switch to OPC.  This routine will
    '                     only return the OPC item that has changed.  The unchanged
    '                    items in the group will not be returned.
    '***********************************************************************************

    Dim i As Long
    
    On Error Resume Next
    
    For i = 1 To NumItems
        'client handles indicates which item in the group has changed.
        Select Case ClientHandles(i)
            Case 1
                txtB3_0.Text = ItemValues(i)
                txtRunStatus.Text = CInt(ItemValues(i)) And 8
            Case 2
                'most suff.  In the N7 registers
                FormatN7Data ItemValues(i)
            Case 3
                txtDiskPresent.Text = ItemValues(i)
            Case 4
                'floating point data
                FormatF8Data ItemValues(i)
        End Select
'        If VarType(ItemValues(i)) And vbArray Then
'            '/* Convert arrayed items into a string before displaying
'            txtData(ClientHandles(i)) = ConvertArrayToString(ItemValues(i))
'        Else
'            '/* Update display
'            txtData(ClientHandles(i)) = ItemValues(i)
'        End If
'        'Update Quality field
'        txtQuality(ClientHandles(i)) = GetQualityString(Qualities(i))
    Next i
End Sub

Private Sub FormatN7Data(N7 As Variant)
    '***********************************************************************************
    '   Created:    10/14/02    By:Ray Raf
    '   Purpose:    To parse out the n7 data to each of the text boxes.  Some
    '                    boxes will be the value of multiple words combined, others
    '                    single words of data.
    '***********************************************************************************
    
        txtCommand.Text = N7(0)
        txtCycleStep.Text = N7(4)
        txtAutoRestartPerm.Text = N7(22) & vbCrLf & N7(23)
        txtStationEnables.Text = N7(10) & vbCrLf & N7(11) & vbCrLf & N7(12) & vbCrLf & N7(13)
        txtStationEnables.Text = txtStationEnables.Text & vbCrLf & N7(14) & vbCrLf & N7(15) & vbCrLf & N7(16)
        txtAlarmCode.Text = N7(2)
        txtClearingDial.Text = N7(3)
        txtHVInterlock.Text = N7(7)
        txtFlashover.Text = N7(31) & vbCrLf & N7(32)
        txtLastCycleTime.Text = N7(34) & vbCrLf & N7(35) & vbCrLf & N7(36) & vbCrLf & N7(37)
        txtLastCycleTime.Text = txtLastCycleTime.Text & vbCrLf & N7(38) & vbCrLf & N7(39) & vbCrLf & N7(40)
        txtStabilized.Text = N7(21)

End Sub

Private Sub FormatF8Data(F8 As Variant)
    '***********************************************************************************
    '   Created:    10/14/02    By:Ray Raf
    '   Purpose:    To parse out the F8 data to each of the text boxes.  Some
    '                    boxes will be the value of multiple words combined, others
    '                    single words of data.
    '***********************************************************************************
    
    txtV1maVoltage.Text = F8(2)
    txtV1mACurrent.Text = F8(3)
    txtV1maResult.Text = F8(2)
    txtWattsTemp.Text = F8(5)
    txtWatts.Text = F8(6)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Remove OPC group definition in server
    opsReadServer.OPCGroups.RemoveAll
    'Drop object reference
    Set opgGroup = Nothing
    'Disconnect from server
    opsReadServer.Disconnect
    'Drop object reference
    Set opsReadServer = Nothing

End Sub

Private Sub tmrCommErrorRetry_Timer()
    '***********************************************************************************
    '   Created:    05/25/00    By:Ray Raf
    '   Purpose:    This routine runs when the comm retry timer times out.
    '                    The routine is used to try to re-establish comms to the
    '                    SLC if comms fail during use or on initial startup.
    '***********************************************************************************
   Dim iMainForm As Integer, frmMN As frmMain
   
    tmrCommErrorRetry.Enabled = False
   GetFormIndexes
   Set frmMN = Forms(globMainIndex)
   frmMN.sbStatusBar.Panels(7).Text = "Retrying Communications @ " & Time()
   
    
    DoEvents
    CheckForLiveComms opgGroup
    
End Sub




Private Sub tmrUpdateRunStatus_Timer()
    tmrUpdateRunStatus.Enabled = False
    
    CheckForLiveComms opgGroup
        
    tmrUpdateRunStatus = True
    
End Sub


Public Function OPCRegWrite(iOPCItem As Integer, sglValue As Single) As Integer
    '***********************************************************************************
    '   Created:    11/06/02    By:Ray Raf
    '   Purpose:    This routine is used to Write a value to the SLC in synch. mode.
    '                     That way if the write fails, the error can be handles immediately.
    '   Inputs:      iOPCItem - The position in the OPC Group Items array of the value to
    '                                    be written to.  This routine assumes that the OPC item
    '                                    points to only a single register.
    '   Returns:     0 if successfull, -1 if Bad comms, -2 if bad item number
    '***********************************************************************************
    Dim INumItems As Long
    Dim arHandles(1 To 1) As Long
    Dim arErrors() As Long
    Dim arValues(1 To 1) As Variant
    Dim arQualities() As Variant
    Dim strQuality As String
    Dim i As Long
    Dim strErrorString As String
    Dim iMaxItems As Integer
    
    iMaxItems = opgGroup.OPCItems.Count
    If iOPCItem < 0 Or iOPCItem > iMaxItems Then
        'item number is wrong.  0 is not allowed or is greater than the largest item configured.
        OPCRegWrite = -2
        Exit Function
    End If
    
    
    'Specify that we are writing 1 item
    INumItems = 1
    'load the server handle for the opc item
    arHandles(1) = opgGroup.OPCItems(iOPCItem).ServerHandle
    arValues(1) = 0
    
    'try to write the data to the device
    opgGroup.SyncWrite INumItems, arHandles, arValues, arErrors
     If arErrors(1) = 0 Then
        'comms OK
        CheckForLiveComms = True
        lblCommError.Visible = False
        Forms(globMainIndex).sbStatusBar.Panels(1).Text = ""
        Forms(globMainIndex).sbStatusBar.Panels(7).Visible = False ' bad comm icon
        Forms(globMainIndex).sbStatusBar.Panels(6).Visible = True  'good comm icon
        globCommStatusOK = True
        txtAlarmCode_Change
    Else
        'RSlinx is running, but can't comm. to SLC
        'comms are in error.  Get error data
        strErrorString = opsReadServer.GetErrorString(arErrors(1))
        Display_Message "Comm Error: " & strErrorString
        CheckForLiveComms = False
        lblCommError.Visible = True
        Forms(globMainIndex).sbStatusBar.Panels(7).Visible = True    'bad comm icon
        Forms(globMainIndex).sbStatusBar.Panels(6).Visible = False   'good comm icon
        globCommStatusOK = False
        tmrCommErrorRetry.Enabled = True
    End If
    
    
    
End Function


Public Function CheckForLiveComms(opgGroup As RsiOPCAuto.OPCGroup) As Boolean
    '***********************************************************************************
    '   Created:    11/04/02    By:Ray Raf
    '   Purpose:    This routine is used to Verify that SLC coms are good by
    '                    trying a synchronos write to the SLC.  This routine assumes
    '                    that the first OPC item is a throwaway register to be used
    '                    for checking comms.  It WILL try to write a value to this register.
    '   Inputs:      opgGroup - the OPC group that we want to verify comms for.
    '   Returns:     True if good comms, False otherwise.
    '***********************************************************************************
    Dim INumItems As Long
    Dim arHandles(1 To 1) As Long
    Dim arErrors() As Long
    Dim arValues(1 To 1) As Variant
    Dim arQualities() As Variant
    Dim strQuality As String
    Dim i As Long
    Dim strErrorString As String
    
    'Specify that we are reading 1 item
    INumItems = 1
    'load the server handle for the 1st opc item (which is only used to check comms)
    arHandles(1) = opgGroup.OPCItems(1).ServerHandle
    arValues(1) = 0
    
    'try to write the data to the device
    opgGroup.SyncWrite INumItems, arHandles, arValues, arErrors
     If arErrors(1) = 0 Then
        'comms OK
        CheckForLiveComms = True
        lblCommError.Visible = False
        Forms(globMainIndex).sbStatusBar.Panels(1).Text = ""
        Forms(globMainIndex).sbStatusBar.Panels(7).Visible = False ' bad comm icon
        Forms(globMainIndex).sbStatusBar.Panels(6).Visible = True  'good comm icon
        globCommStatusOK = True
        txtAlarmCode_Change
    Else
        'RSlinx is running, but can't comm. to SLC
        'comms are in error.  Get error data
        strErrorString = opsReadServer.GetErrorString(arErrors(1))
        Display_Message "Comm Error: " & strErrorString
        CheckForLiveComms = False
        lblCommError.Visible = True
        Forms(globMainIndex).sbStatusBar.Panels(7).Visible = True    'bad comm icon
        Forms(globMainIndex).sbStatusBar.Panels(6).Visible = False   'good comm icon
        globCommStatusOK = False
        tmrCommErrorRetry.Enabled = True
    End If
    
    
    
End Function

Public Function PLCRegWrite(strAddr As String, sglValue As Single) As Integer

    '***********************************************************************************
    '   Created:    11/06/02    By:Ray Raf
    '   Purpose:    This routine is used to Write a value to the SLC in synch. mode.
    '                     That way if the write fails, the error can be handles immediately.
    '   Inputs:      iOPCItem - The position in the OPC Group Items array of the value to
    '                                    be written to.  This routine assumes that the OPC item
    '                                    points to only a single register.
    '   Returns:     0 if successfull, -1 if Bad comms, -2 if bad item number
    '***********************************************************************************
    Dim INumItems As Long
    Dim arHandles(1 To 1) As Long
    Dim arErrors() As Long
    Dim arValues(1 To 1) As Variant
    Dim arQualities() As Variant
    Dim strQuality As String
    Dim i As Long
    Dim strErrorString As String
    Dim iMaxItems As Integer
    Dim iHandle As Long
    Const strTopic = "[TesterOPC]"
    Dim opiThisItem As OPCItem
    
    With opgGroup
        iMaxItems = .OPCItems.Count
        iHandle = opgGroup.OPCItems(iMaxItems).ClientHandle
        'increment to the next available handle
        iHandle = iHandle + 1
        'add the passed address to the OPC group as a new item
        Set opiThisItem = opgGroup.OPCItems.AddItem(strTopic & strAddr, iHandle)
        
        
        'Specify that we are writing 1 item
        INumItems = 1
        'load the server handle for the opc item
        arHandles(1) = opiThisItem.ServerHandle
        arValues(1) = sglValue
        
        'try to write the data to the device
        opgGroup.SyncWrite INumItems, arHandles, arValues, arErrors
         If arErrors(1) = 0 Then
            'comms OK
            PLCRegWrite = 0
        Else
            'RSlinx is running, but can't comm. to SLC
            'comms are in error.  Get error data
            strErrorString = opsReadServer.GetErrorString(arErrors(1))
            Display_Message "Comm Error: " & strErrorString
            PLCRegWrite = -1
        End If
            
        Set opiThisItem = Nothing
        
        'now delete the added item from the group.
        opgGroup.OPCItems.Remove INumItems, arHandles, arErrors()
    End With
            
    

End Function

