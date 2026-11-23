Attribute VB_Name = "modPrinting"
Option Explicit

'-------------------------------MODULE COMMENT------------------------------------
'Created:   09/22/00    By:Ray Raf
'Purpose:   All printing functions should be in this module.  This includes
'                both Barcode printing and Report printing.
'-------------------------------------------------------------------------------------------------
    
 Private m_PrevPrinter As String
 
Public Sub PrintAllBatchGroupLabels(BGToPrint As BatchGroup)
    '------------------------------------------------------------------------------------------------
    'Created:   09/01/00    By:Ray Raf
    'Purpose: To send all of the required data to the barcode printer to
    '              print all of the labels for this batch.
    '-------------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    Dim bbd As BatchBarcodeData
    Dim cbd As ContainerBarcodeData
    Dim wcEmpty As WeighCompound
    Dim str1stBatchNum As String, strLastBatchNum As String
    Dim iContLoop As Integer, iWeighLoop As Integer
    Dim strGroup As String, iMatNumber As Integer
    Dim strWGroup As String

    With BGToPrint
        'do the batch barcodes first
        PrintAllBatchLabels BGToPrint

        'do the container barcodes next--------------------------------------------------------
        iContLoop = 1
        strGroup = Trim(.Containers(iContLoop).strGroup)
        'loop thru each Group in Container array
        Do
            PrintAllLabelsForCont BGToPrint, " ", iContLoop
            'move to next container
            iContLoop = iContLoop + 1
            'get container group
            strGroup = Trim(.Containers(iContLoop).strGroup)
        Loop While Len(strGroup) > 0 And iContLoop < 21

        'check for a pre-batched material container
        If .sglPBWeight > 0.009 Then
            'there is a pre-batched material container label to print
            'LoadDataForContPrint BGToPrint, 99, .lngPB1stContainer, .iTotalBatchesInGroup, cbd
            PrintAllLabelsForCont BGToPrint, " ", 99
        End If

    End With
End Sub

Public Sub PrintAllBatchLabels(BG As BatchGroup)
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   Print a batch label for each batch in Batch Group
    'Inputs:    BG - Batch Group to use
    '   Rev'd:  12/27/02        By: Ray Raf
    '   Note:   Changed routine to print 3 sets of Batch Labels
    '-------------------------------------------------------------------------------------------------
    Dim iLinkLoop As Integer
    Dim bbd As BatchBarcodeData
    
    For iLinkLoop = 1 To BG.iNumberOfLinks
        LoadDataforBatchLabelPrint BG, iLinkLoop, BG.str1stBatchNum(iLinkLoop), _
                            BG.iNumBatches(iLinkLoop), bbd
        PrintBatchBarcode bbd
        PrintBatchBarcode bbd
        PrintBatchBarcode bbd
    Next iLinkLoop
        
End Sub

Public Sub PrintOneBatchLabel(BG As BatchGroup, strBatchNumber As String)
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   Print a batch label for the passed batch number in Batch Group
    'Inputs:    BG - Batch Group to use
    '              strBatchNumber - the batch number to print a label for
    '-------------------------------------------------------------------------------------------------
    Dim bbd As BatchBarcodeData
    Dim iLoop As Integer
    Dim iLink As Integer
    Dim lngBatchNum As Long
    Dim lngStartBatch As Long, lngEndBatch As Long
    
    'need to find the link number for the batch to print
    If BG.iNumberOfLinks > 1 Then
        lngBatchNum = CLng(strBatchNumber)
        For iLoop = 1 To BG.iNumberOfLinks
            lngStartBatch = CLng(BG.str1stBatchNum(iLoop))
            lngEndBatch = lngStartBatch + BG.iNumBatches(iLoop) - 1
            If lngBatchNum >= lngStartBatch And lngBatchNum <= lngEndBatch Then
                'this is the link group that the lable belongs in
                iLink = iLoop
                Exit For
            End If
        Next iLoop
    Else
        iLink = 1
    End If
    
    LoadDataforBatchLabelPrint BG, iLink, strBatchNumber, 1, bbd
    PrintBatchBarcode bbd
    
End Sub


Public Sub LoadDataforBatchLabelPrint(BG As BatchGroup, iLink As Integer, _
                strStartBatchNum As String, iNumBatches As Integer, bbd As BatchBarcodeData)
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   load data for the batch label to print
    'Inputs:    BG - Batch Group to use
    '              iLink - the link number for the 1st batch label to print
    '              lngStartBatchNum - starting Batch Number to Print
    '              iNumBatches - Number of batch labels to print
    '              bbd - filled in BatchBarcodeData variable to use for printing
    '-------------------------------------------------------------------------------------------------
    Dim rstBatch As New ADODB.Recordset
    Dim SQL As String
    
    With BG
        bbd.strCompound = .strCompound
        bbd.strCycle = .strCycle
        bbd.strMixer = .strMixer
        bbd.strTargetWeight = Format(.sglBatchWeight, "0.00") & " LBS"
    End With
    
    If iNumBatches > BG.iNumBatches(iLink) Then
        bbd.iNumToPrint = BG.iNumBatches(iLink)
    Else
        bbd.iNumToPrint = iNumBatches
    End If
    bbd.strBatchNumber = strStartBatchNum
    bbd.strWorkOrder = BG.strWorkOrder(iLink)
    
    'get issue time and date
    SQL = "SELECT * FROM BatchRecord WHERE BatchNumber = '" & strStartBatchNum & "' "
    rstBatch.Open SQL, dbconResults, adOpenDynamic, adLockReadOnly
    bbd.strIssueTime = Format(rstBatch![IssueTime], "hh:mm AM/PM")
    bbd.strIssueDate = Format(rstBatch![IssueTime], "mm/dd/yy")
    rstBatch.Close

End Sub

Public Sub PrintAllLabelsForCont(BG As BatchGroup, strGroup As String, Optional iContLoc As Integer = 0)
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   Print passed group's label for each batch in Batch Group
    'Inputs:    BG - Batch Group to use
    '              strGroup - the Group to print the labels for.
    '              iContLoc - if known, the position that the Group is found in.
    '                              if specified, strgroup can be blank
    '-------------------------------------------------------------------------------------------------
    Dim iLocation As Integer
    Dim cbd As ContainerBarcodeData
    Dim lngStartContNum As Long
    Dim iNumCont As Integer
    
    'get location in container array (or as a Pre-batches material)
    If iContLoc > 0 Then
        iLocation = iContLoc
    Else
        iLocation = FindLocForGroup(BG, strGroup)
    End If
    
    'get 1st container number
    If iLocation = 99 Then
        lngStartContNum = BG.lngPB1stContainer
    Else
        lngStartContNum = BG.Containers(iLocation).lng1stContNum
    End If
    
    iNumCont = BG.iTotalBatchesInGroup
    
    LoadDataForContPrint BG, iLocation, lngStartContNum, iNumCont, cbd
    PrintContainerBarcode cbd
    
End Sub
    
Public Sub PrintOneContLabel(BG As BatchGroup, strGroup As String, _
                        lngContNum As Long, Optional iContLoc As Integer = 0)
    '------------------------------------------------------------------------------------------------
    'Created:   12/29/00    By:Ray Raf
    'Purpose:   Print passed Container's label
    'Inputs:    BG - Batch Group to use
    '              strGroup - the Group to print the labels for.
    '              lngContNum - The container to print the label for
    '              iContLoc - if known, the position that the Group is found in.
    '                              if specified, strGroup can be blank
    '-------------------------------------------------------------------------------------------------
    Dim iLocation As Integer
    Dim cbd As ContainerBarcodeData
    Dim iNumCont As Integer
    
    'get location in container array (or as a Pre-batches material)
    If iContLoc > 0 Then
        iLocation = iContLoc
    Else
        iLocation = FindLocForGroup(BG, strGroup)
    End If

    LoadDataForContPrint BG, iLocation, lngContNum, 1, cbd
    PrintContainerBarcode cbd
    
End Sub

Public Function FindLocForGroup(BG As BatchGroup, strGroup As String) As Integer
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   To find the location of a group in the Container array for
    '                   the given batch group
    'Inputs:    BG - Batch Group to use
    '              strGroup - the Group to find
    'Returns:   array location (1 to 20) if group found in Containers array,
    '               99 if found in Pre-Batched area.
    '               0 if not found
    '-------------------------------------------------------------------------------------------------
    Dim iLoop As Integer
    
    With BG
        For iLoop = 1 To 20
            If Trim$(strGroup) = Trim$(.Containers(iLoop).strGroup) Then
                'we found our group
                FindLocForGroup = iLoop
                Exit Function
            End If
        Next iLoop
        
        'didn't find it in the containers array, check the PB material
        If Trim$(.strPBGroup) = Trim$(strGroup) Then
            FindLocForGroup = 99
            Exit Function
        End If
        
        'didn't find group at all
        FindLocForGroup = 0
    End With
End Function

Public Function LoadDataForContPrint(BG As BatchGroup, iLoc As Integer, _
                lngStartContNum As Long, iNumCont As Integer, ByRef cbd As ContainerBarcodeData)
    '------------------------------------------------------------------------------------------------
    'Created:   12/15/00    By:Ray Raf
    'Purpose:   To load the data needed to print a container label into
    '                 a ContainerBarcodeData variable
    'Inputs:    BG - Batch Group to use
    '              iLoc - the location in the container array for the data or
    '                       99 to indicate the pre-batched material
    '              lngStartContNumber - the starting container number to use
    '                       (not neccessarily = lng1stContainer)
    '               iNumCont - number of consecutive container labels to print
    'Output     cbd - a filled in ContainerBarcodeData variable
    '-------------------------------------------------------------------------------------------------
    Dim wcEmpty As WeighCompound
    Dim iLoop As Integer
    Dim str1stBatchNum As String, strLastBatchNum As String
    Dim iMatNumber As Integer, iWeighLoop As Integer
    Dim strGroup As String, strWGroup As String
    
    With BG
        cbd.strCompound = .strCompound
        cbd.strCycle = .strCycle
        cbd.strMixer = .strMixer
        'generate batch number(s) string
        str1stBatchNum = .str1stBatchNum(1)
        cbd.iNumToPrint = iNumCont
        strLastBatchNum = Format(CLng(.str1stBatchNum(1)) + .iTotalBatchesInGroup - 1, "00000000")
        cbd.strBatchString = str1stBatchNum
        If .iTotalBatchesInGroup > 1 Then
            cbd.strBatchString = cbd.strBatchString & " to " & strLastBatchNum
        End If
    
        If iLoc = 99 Then
            'use pre-batched stuff
            cbd.lng1stContainerNum = lngStartContNum
            cbd.WC(1).sglTargetWeight = .sglPBWeight
            cbd.WC(1).sglTolerance = .sglPBTol
            cbd.WC(1).strGroup = .strPBGroup
            cbd.WC(1).strMaterialCode = .strPBMaterial
            cbd.WC(1).strUnits = .strPBUnits
            cbd.WC(1).strContainerType = .strPBContainer
            For iLoop = 2 To 10
                'clear out all leftover data
                CopyWeighCompound cbd.WC(iLoop), wcEmpty
            Next iLoop
        Else
            'use container array
            strGroup = Trim(.Containers(iLoc).strGroup)
            cbd.lng1stContainerNum = lngStartContNum
            iMatNumber = 1
            iWeighLoop = 1
            'get group for 1st weighment
            strWGroup = Trim(.Weighments(iWeighLoop).strGroup)
            'loop thru each weightment
            Do
                If strWGroup = strGroup Then
                    'if groups match we need to put this weighment on this label
                    CopyWeighCompound cbd.WC(iMatNumber), .Weighments(iWeighLoop)
                    'index to next position in the weighment array in cbd
                    iMatNumber = iMatNumber + 1
                End If
                iWeighLoop = iWeighLoop + 1
                'get group for next weighment
                strWGroup = Trim(.Weighments(iWeighLoop).strGroup)
            'loop as long as we find good groups and haven't reached the end of the arrays
            Loop While Len(strWGroup) > 0 And iWeighLoop < 21 And iMatNumber < 11
            
            'now that we found all of the weighments that belong on this label, clean out
            'the rest of the cbd.WC array so that no left over data gets printed
            If iMatNumber < 11 Then
                For iLoop = iMatNumber To 10
                    CopyWeighCompound cbd.WC(iLoop), wcEmpty
                Next iLoop
            End If
        End If
    End With
End Function

Public Sub PrintBatchBarcode(bbdLabel As BatchBarcodeData)
    '------------------------------------------------------------------------------------------------
    'Created:   09/22/00    By:Ray Raf
    'Purpose:   print a single or a range of batch barcode labels
    'Inputs:    bbdLabel - contains the data needed to print a Batch
    '                    barcode label.
    'Rev'd:     01/22/01    By:Ray Raf
    'Note:      Changed compound, Cycle, Mixer data print to match that
    '              on the container labels
    '-------------------------------------------------------------------------------------------------
    Dim bStillPrint As Boolean
    
    'change default printer to label printer
    bStillPrint = ChangeToLabelPrinter
    If bStillPrint = False Then
        'there was some trouble with the printer.  aborting label print
        Exit Sub
    End If
       
    With bbdLabel
        Printer.Print "     "       'not sure why but need this line for NT platforms
        Printer.Print "^XA"
        Printer.Print "^LH10,10"
        Printer.Print "^FO509,250^A0R,50^FDBatch Number: ^FS" 'font 0 , rotated 90 deg, 50 dots high
        Printer.Print "^FO509,570^A0R,50^SN" & .strBatchNumber & ",1,Y^FS"
        Printer.Print "^FO350,325^BY2,3.0^B3R,N,152,Y,N^SNB" & .strBatchNumber & ",1,Y^FS"
        Printer.Print "^FO300,50^GB0,915,5^FS"     'horz line
        Printer.Print "^FO200,140^A0R,80^FD" & .strCompound & "^FS"
        Printer.Print "^FO150,182^A0R,40^FD" & .strCycle & "^FS"
        Printer.Print "^FO150,280^A0R,40^FD" & .strMixer & "^FS"
'        Printer.Print "^FO250,141^A0R,40^FDCompound: " & .strCompound & "^FS"
'        Printer.Print "^FO200,226^A0R,40^FDCycle: " & .strCycle & "^FS"
'        Printer.Print "^FO150,200^A0R,40^FDMixer #: " & .strMixer & "^FS"
        Printer.Print "^FO100,85^A0R,40^FDTarget Weight:  " & .strTargetWeight & "^FS"
        Printer.Print "^FO30,550^GB270,0,5^FS"      'vert line
        Printer.Print "^FO250,600^A0R,40^FDWork Order: " & .strWorkOrder & "^FS"
        Printer.Print "^FO200,600^A0R,40^FDIssue Date: " & .strIssueDate & "^FS"
        Printer.Print "^FO150,600^A0R,40^FDIssue Time: " & .strIssueTime & "^FS"
        Printer.Print "^FO30,980^A0I,20^FDForm 0901378-02  rev-:01/01/01^FS"
        Printer.Print "^PV1"
        Printer.Print "^PQ" & .iNumToPrint
        Printer.Print "^XZ"
    End With
    
    Printer.EndDoc

    'resets system default printer to what it was before label printing (if necessary)
    ChangePrinterBack
    
End Sub

Public Sub PrintContainerBarcode(cbdLabel As ContainerBarcodeData)
    '------------------------------------------------------------------------------------------------
    'Created:   09/25/00    By:Ray Raf
    'Purpose:   print a single or a range of Container barcode labels
    'Inputs:    cbdLabel - contains the data needed to print a Container
    '                    barcode label.
    '-------------------------------------------------------------------------------------------------
    Dim iNumMatsFound As Integer
    Dim strCont As String
    Dim iLoop As Integer
    Dim iRow As Integer
    Dim strWeightString As String
    Dim strWeightStart As Integer
    Dim bStillPrint As Boolean
    
    'change default printer to label printer
    bStillPrint = ChangeToLabelPrinter
    If bStillPrint = False Then
        'there was some trouble with the printer.  aborting label print
        Exit Sub
    End If
    
    With cbdLabel
        strCont = Format(.lng1stContainerNum, "0000000000")
        Printer.Print "     "       'not sure why but need this line for NT platforms
        Printer.Print "^XA"
        Printer.Print "^LH10,10"
        Printer.Print "^FO329,50^A0R,210^FD" & .WC(1).strGroup & "^FS"
        Printer.Print "^FO509,350^A0R,50^SNCONTAINER #:" & strCont & ",1,Y^FS"
        Printer.Print "^FO380,420^BY2,3.0^B3R,N,120,Y,N^SNC" & strCont & ",1,Y^FS"
        Printer.Print "^FO335,50^GB0,915,5^FS"
        Printer.Print "^FO290,100^A0R,40^FDFor Batch(es): " & .strBatchString & "^FS"
        Printer.Print "^FO250,100^A0R,30^FDMaterial #     Target Weight^FS"
        Printer.Print "^FO249,100^GB0,350,2^FS"
        For iLoop = 1 To 10
            'we can print up to 10 materials for weighment on one label
            If Len(.WC(iLoop).strGroup) > 2 Then
                'there is an entry in this position
                Select Case iLoop
                    Case 1 To 7
                        iRow = 250 - (iLoop * 35)   'positioning on label
                        Printer.Print "^FO" & iRow & ",126^A0R,30^FD" & .WC(iLoop).strMaterialCode & "^FS"
                        If .WC(iLoop).sglTargetWeight > 9.99 Then
                            strWeightStart = "313"
                        Else
                            strWeightStart = "326"
                        End If
                        strWeightString = Format(.WC(iLoop).sglTargetWeight, "0.00 ")
                        strWeightString = strWeightString & .WC(iLoop).strUnits
                        Printer.Print "^FO" & iRow & "," & strWeightStart & "^A0R,30^FD" & strWeightString & "^FS"
                    Case 8 To 10
                        'second column of weights
                        If iLoop = 8 Then
                            'print header
                            Printer.Print "^FO250,557^A0R,30^FDMaterial #     Target Weight^FS"
                            Printer.Print "^FO249,557^GB0,350,2^FS"
                        End If
                        iRow = 250 - ((iLoop - 7) * 35)   'positioning on label
                        Printer.Print "^FO" & iRow & ",573^A0R,30^FD" & .WC(iLoop).strMaterialCode & "^FS"
                        If .WC(iLoop).sglTargetWeight > 9.99 Then
                            strWeightStart = "760"
                        Else
                            strWeightStart = "773"
                        End If
                        strWeightString = Format(.WC(iLoop).sglTargetWeight, "0.00 ")
                        strWeightString = strWeightString & .WC(iLoop).strUnits
                        Printer.Print "^FO" & iRow & "," & strWeightStart & "^A0R,30^FD" & strWeightString & "^FS"
                End Select
            Else
                'no more entries - exit loop
                Exit For
            End If
        Next iLoop
            
        Printer.Print "^FO00,507^GB280,0,4^FS"
        Printer.Print "^FO145,507^GB0,462,4^FS"
        'Printer.Print "^FO100,597^A0R,40^FDCompound: " & .strCompound & "^FS"
        'Printer.Print "^FO50,682^A0R,40^FDCycle: " & .strCycle & "^FS"
        'Printer.Print "^FO00,557^A0R,40^FDWhere Mixed: " & .strMixer & "^FS"
        Printer.Print "^FO50,640^A0R,80^FD" & .strCompound & "^FS"
        Printer.Print "^FO00,682^A0R,40^FD" & .strCycle & "^FS"
        Printer.Print "^FO00,780^A0R,40^FD" & .strMixer & "^FS"
        Printer.Print "^FO30,980^A0I,20^FDForm 0901378-01  rev-:01/01/01^FS"
        Printer.Print "^PQ" & .iNumToPrint
        Printer.Print "^XZ"
    End With
    
    Printer.EndDoc
    
    'resets system default printer to what it was before label printing (if necessary)
    ChangePrinterBack
    
End Sub

Public Sub PrintIt(LMarg As Integer, TMarg As Integer, InptStr As Variant, PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    'Created:   11/20/00    By:Ray Raf
    'Purpose:   print a string to a passed location on the page
    'Inputs:    LMarg - Left position of text
    '              TMarg - Top Position of Text
    '              InptStr - String to print
    '              PrintDev - either a printer object or a PictureBox object that can be drawn on.
    '-------------------------------------------------------------------------------------------------
    
    PrintDev.CurrentX = LMarg
    PrintDev.CurrentY = TMarg
    PrintDev.Print InptStr
    
End Sub


Public Function ChangeToLabelPrinter() As Boolean
    '------------------------------------------------------------------------------------------------
    'Created:   12/05/00    By:Ray Raf
    'Purpose:   to change the default printer to the default label printer for
    '                printing labels.
    'Returns:   True if printing should proceed, false if printing should be canceled
    '-------------------------------------------------------------------------------------------------
    Dim X As Printer
    Dim bFound As Boolean
    
    '1st save previous printer
    m_PrevPrinter = Printer.DeviceName
    
    'only need to change to label printer if it is not already the default
    If Printer.DeviceName <> globDefaultLabelPrinter Then
        ChangeToLabelPrinter = False  'initialize to: printer not found
        For Each X In Printers
            If X.DeviceName = globDefaultLabelPrinter Then
                Set Printer = X
                ChangeToLabelPrinter = True
                Exit Function 'exit function since we found our printer
            End If
        Next X
        
        'if we got here, then the default label printer could not be found
        MsgBox "The Default label printer could not be found." & vbCrLf & _
                    "In the Printer selection window, select another" & vbCrLf & _
                    "printer or press 'Cancel' to abort printing.", vbOKOnly + vbCritical
        fMainForm.dlgCommonDialog.CancelError = True
        On Error GoTo NoPrint
        fMainForm.dlgCommonDialog.ShowPrinter
        ChangeToLabelPrinter = True
    Else
        ChangeToLabelPrinter = True
    End If
    
    Exit Function

NoPrint:
    'default not found, another printer not selected
    MsgBox "Label Printing will be aborted!", vbOKCancel + vbCritical
    ChangeToLabelPrinter = False
       
End Function

Public Sub ChangePrinterBack()
    '------------------------------------------------------------------------------------------------
    'Created:   12/05/00    By:Ray Raf
    'Purpose:   to reset the system default printer after printing labels.
    '-------------------------------------------------------------------------------------------------
    Dim prt As Printer
    
    For Each prt In Printers
        If prt.DeviceName = m_PrevPrinter Then
            Set Printer = prt
            Exit Sub 'exit sub since we found our printer
        End If
    Next prt

    'if we got here, the previous printer was deleted after it was saved by the program.
    'just leave the default printer as is and exit
    
End Sub
