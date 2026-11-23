Attribute VB_Name = "modDBarcodeFunctions"
Option Explicit

'contains the data needed to label a disc
Type DiskData
    sglV10kA As Single
    sglVref As Single
    sglSCP As Single
    strBatchNumber As String
    iDiskNumber As Integer
    strCategory As String
    strDiskType As String
End Type


Type RackInfo
    'contains data related to a stack location
    Stop As Integer     '(0 or greater)
    Shelf As String      '(A,B,C,D,W,X,Y,Z)
    Position As Integer     '(1 to 10)
End Type

Type BoxInfo
    'info encoded in Box barcode
    BoxNumber As String
    BatchNumber As String
    FirstDisk As Integer
    LastDisk As Integer
End Type

Type BarcodeData
    'can contain info for any of the barcodes in the SCP project
    'only the relevant info for the type indicated in BarcodeType would
    'be filled in by the ParseBarcode routine.  However other logic may
    'use multiple parts of the type simultaniously
    BarcodeType As String       'empty, B, C, D, R, T
    DiskData As DiskData        'all disk data (D type)
    BoxInfo As BoxInfo            'info about a disk box (B type)
    CartPos As String               'location in disk cart (C type)
    TubeNumber As String         'the tube ID (T type)
    RackInfo As RackInfo          'location of a tube (R type)
End Type

Public Function ParseBarcode(strBarcode As String) As BarcodeData
    '------------------------------------------------------------------------------------
    '   Created:    05/06/05    By:Ray Raf
    '   Purpose:    parse all of the barcodes associated with the FACTS project
    '   Output(s):  bd - a variable containing the data parsed fromt the barcode.
    '                    bd.BarcodeType - indicates what type of barcode was parsed
    '                           (B, C, D, R, or T) or is blank if un-successfully parsed.
    '                           Only the variables associated with the indicated barcode
    '                           type are valid on return.
    '   Returns:    True if barcode was successfully parsed, false otherwise
    '   Rev'd:          06/22/06    By: Ray Raf
    '   Note:       modified V10ka parse for Disk barcode.  Disk with V10ka values
    '                  in excess of 10kv are now being tested and v10Ka value had to
    '                  be recoded.
    '------------------------------------------------------------------------------------
    Dim bd As BarcodeData
    Dim strType As String
    Dim strBlock As String
    Dim iVal As Integer
    Dim bIsOK As Boolean
    Dim strLast2 As String
    
    If Len(strBarcode) < 3 Then
        'no valid bacode is less than 3 characters
        bd.BarcodeType = ""
        ParseBarcode = bd
        Exit Function
    End If
    
    'get type id from beginning of string
    strType = Left(strBarcode, 1)
    Select Case LCase(strType)
        Case "b"    'box barcode--------------------------------------------------------------------------------------------
            If Len(strBarcode) = 19 Then
                With bd.BoxInfo
                    bIsOK = True
                    .BoxNumber = Mid(strBarcode, 2, 4)
                    .BatchNumber = Mid(strBarcode, 6, 6)
                    'parse 1st disk number
                    strBlock = Mid(strBarcode, 12, 4)
                    If IsNumeric(strBlock) Then
                        .FirstDisk = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    'parse last disk number
                    strBlock = Mid(strBarcode, 16, 4)
                    If IsNumeric(strBlock) Then
                        .LastDisk = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    
                    If bIsOK = True Then
                        'valid barcode
                        bd.BarcodeType = strType
                    Else
                        'not a valid barcode
                        bd.BarcodeType = ""
                    End If
                End With
            Else
                'not a valid box barcode
                bd.BarcodeType = ""
            End If
        
        Case "c"    'cart position barcode------------------------------------------------------------------------
            If Len(strBarcode) = 3 Then
                bd.CartPos = Right(strBarcode, 2)
                bd.BarcodeType = strType
            Else
                'not a valid car position barcode
                bd.BarcodeType = ""
            End If
        
        Case "d"    'disk data barcode-----------------------------------------------------------------------------
            'begining of barcoded data. format = DB444445555CCCC1111222333TYPE
            'D = Disk Data barcode indicator
            '1111 = SCP volts or blanks
            '222 = v10ka
            '333 = Vref
            'TYPE = disc type (IE AT32,  AT33, AT34)
            'B44444 = Batch number
            '5555 = disc number
            'CCCC = category or reject code
            If Len(strBarcode) >= 29 Then
                With bd.DiskData
                    bIsOK = True
                    'batchnumber
                    .strBatchNumber = Mid(strBarcode, 2, 6)
                    'disk number
                    strBlock = Mid(strBarcode, 8, 4)
                    If IsNumeric(strBlock) Then
                        .iDiskNumber = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    'Category
                    .strCategory = Mid(strBarcode, 12, 4)
                    'scp value
                    strBlock = Trim(Mid(strBarcode, 16, 4))
                    If Len(strBlock) <> 4 Then
                        'no SPC value, set to -1
                        .sglSCP = 0
                    Else
                        If IsNumeric(strBlock) Then
                            .sglSCP = CInt(strBlock)
                        Else
                            bIsOK = False
                        End If
                    End If
                    'v10ka
                    'if V10kA is >= 10kV then the value will be encoded "X01" with
                    '"X" representing a value of 10.  the remaining digits will be added
                    'to the end of the value.  In the example the V10kA value would be 1001
                    strBlock = Mid(strBarcode, 20, 3)
                    If IsNumeric(strBlock) Then
                        .sglV10kA = CInt(strBlock)
                    Else
                        'could posibly be >= 10Kv. and have an X in the number
                        If Left(strBlock, 1) = "X" Then
                            'v10ka pobably >=10 kV.
                            strLast2 = Right(strBlock, 2)
                            If IsNumeric(strLast2) Then
                                .sglV10kA = 1000 + CInt(strLast2)
                            Else
                                bIsOK = False
                            End If
                        Else
                            bIsOK = False
                        End If
                    End If
                    'Vref
                    strBlock = Mid(strBarcode, 23, 3)
                    If IsNumeric(strBlock) Then
                        .sglVref = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    'Disk type
                    .strDiskType = Mid(strBarcode, 26, 4)
                    
                    If bIsOK = True Then
                        'valid barcode
                        bd.BarcodeType = strType
                    Else
                        'not a valid barcode
                        bd.BarcodeType = ""
                    End If
                End With
            Else
                'not a valid disk data barcode
                bd.BarcodeType = ""
            End If
            
        Case "r"    'rack position barcode--------------------------------------------------------------------------
            If Len(strBarcode) = 6 Then
                With bd.RackInfo
                    bIsOK = True
                    'stop 1st
                    strBlock = Mid(strBarcode, 2, 2)
                    If IsNumeric(strBlock) Then
                        .Stop = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    'shelf
                    .Shelf = Mid(strBarcode, 4, 1)
                    'position
                    strBlock = Mid(strBarcode, 5, 2)
                    If IsNumeric(strBlock) Then
                        .Position = CInt(strBlock)
                    Else
                        bIsOK = False
                    End If
                    
                    If bIsOK = True Then
                        'valid barcode
                        bd.BarcodeType = strType
                    Else
                        'not a valid barcode
                        bd.BarcodeType = ""
                    End If
                End With
            Else
                'not a valid car position barcode
                bd.BarcodeType = ""
            End If
        
        Case "t"    'tube ID barocode--------------------------------------------------------------------------
            If Len(strBarcode) = 5 Then
                bd.TubeNumber = Right(strBarcode, 4)
                bd.BarcodeType = strType
            Else
                'not a valid car position barcode
                bd.BarcodeType = ""
            End If
        
        Case Else   'barcode did not start with any of the allowable type identifiers
            bd.BarcodeType = ""
    End Select
    
    ParseBarcode = bd
        
End Function

Public Function TestParseRoutine(strTestBarcode As String)
    'used to test parse routine
    Dim fred As BarcodeData
    
    fred = ParseBarcode(strTestBarcode)
    DisplayBarcodeInfo fred, strTestBarcode
    
End Function
    
Public Sub DisplayBarcodeInfo(fred As BarcodeData, strBarcode As String)
    Dim strMsg As String

    Select Case LCase(fred.BarcodeType)
        Case "b"
            With fred.BoxInfo
                strMsg = strBarcode & vbCrLf & "BXXXXYYYYYYAAAABBBB" & vbCrLf & vbCrLf
                strMsg = strMsg & "box number = " & Format(.BoxNumber, "@@@@") & vbCrLf
                strMsg = strMsg & "Batch Number = " & Format(.BatchNumber, "@@@@@@") & vbCrLf
                strMsg = strMsg & "First Disk = " & Format(.FirstDisk, "0000") & vbCrLf
                strMsg = strMsg & "Last Disk = " & Format(.LastDisk, "0000")
            End With
        Case "c"        'cart data
                strMsg = strBarcode & vbCrLf & "CXX" & vbCrLf & vbCrLf
                strMsg = strMsg & "Position = " & fred.CartPos & vbCrLf
        Case "d"
            With fred.DiskData
                strMsg = strBarcode & vbCrLf & "DB444445555CCCC1111222333TYPE" & vbCrLf & vbCrLf
                strMsg = strMsg & "SCP Voltage = " & .sglSCP & vbCrLf
                strMsg = strMsg & "V10kA Voltage = " & .sglV10kA & vbCrLf
                strMsg = strMsg & "Vref Voltage = " & .sglVref & vbCrLf
                strMsg = strMsg & "Disk Type = " & .strDiskType & vbCrLf
                strMsg = strMsg & "Batch Number = " & .strBatchNumber & vbCrLf
                strMsg = strMsg & "Disk Number = " & .iDiskNumber & vbCrLf
                strMsg = strMsg & "Category = " & .strCategory & vbCrLf
            End With
        Case "r"        'rack position data
            With fred.RackInfo
                strMsg = strBarcode & vbCrLf & "RXXYZZ" & vbCrLf & vbCrLf
                strMsg = strMsg & "Stop = " & .Stop & vbCrLf
                strMsg = strMsg & "Shelf = " & .Shelf & vbCrLf
                strMsg = strMsg & "Position = " & .Position & vbCrLf
            End With
        Case "t"        'tube data
            With fred.BoxInfo
                strMsg = strBarcode & vbCrLf & "TXXXX" & vbCrLf & vbCrLf
                strMsg = strMsg & "Tube ID = " & fred.TubeNumber & vbCrLf
            End With
        Case Else
            strMsg = strBarcode & vbCrLf & vbCrLf
            strMsg = strMsg & "Invalid data!"
    End Select
    
    MsgBox strMsg, vbOKOnly, "Result of barcode parse"
    
    
End Sub


'Public Sub commBarcode_OnComm()
'    '------------------------------------------------------------------------------------------------
'    '   Created:   01/05/04    By:Ray Raf
'    '   Purpose:   To retreive barcode data from a serial port using an MSComm
'    '                   Control.  Since the barcodes
'    '                   have diferent lengths and structures, the barcode readers are
'    '                   programmed to add a { at the beginning of the string and a }
'    '                   at the end of the string.
'    '   Note:        the comm control must have it's RThreshold property set to 1
'    '                   in order for the OnComm event to occur.
'    '-------------------------------------------------------------------------------------------------
'    Dim strOneChar As String
'    Dim strBarcode As String
'    Dim iCharsLeft As Integer
'    Dim strJunk As String
'    Dim iResult As Integer
'    Dim strData As String
'    Dim strAddr As String
'    Dim iLoop As Integer
'    Dim strCodeToCheck As String
'    Dim iInputLength As Integer
'    Dim strStartChar As String
'    Dim strEndChar As String
'    Dim strDataFromPort As String
'    Dim iStartPos As Integer, iEndPos As Integer
'    Dim bcdReadBarcodeData As BarcodeData
'
'    If commBarcode.CommEvent = 2 Then
'        'data has come into the port
'        'm_CommActive = True
'        Delay 500       'delay to allow all barcode characters to show up
'CheckForAnotherBarcode:         'in case other barcodes came in while first one being handled
'        DoEvents
'
'        ' *************************************************************************************************
'        '  the following code is used to read the barcode input.
'        '*******************************************************************************************************
'        commBarcode.InputLen = 0 ' read the whole input buffer
'        strDataFromPort = commBarcode.Input
'        Debug.Print strDataFromPort
'        iEndPos = Len(strDataFromPort)
'        Do While iEndPos > 0 And Len(strDataFromPort) > 0
'            'find opening brace
'            iEndPos = 0
'            iStartPos = InStr(1, strDataFromPort, "{")
'            If iStartPos > 0 Then
'                'opening brace found, find end brace
'                iEndPos = InStr(iStartPos, strDataFromPort, "}")
'                If iEndPos > 0 Then
'                    'end brace found, get data between braces
'                    strCodeToCheck = Mid(strDataFromPort, iStartPos + 1, iEndPos - iStartPos - 1)
'                    iInputLength = Len(strCodeToCheck)
'                    strBarcode = strCodeToCheck
'                    'delete the data already used from the string
'                    strDataFromPort = Right(strDataFromPort, Len(strDataFromPort) - iEndPos)
'                    Debug.Print strBarcode, strDataFromPort
'
'                    'we have the barcode at this point.
'                    bcdReadBarcodeData = ParseBarcode(strBarcode)
'
'                    'decide what to do with read barcode
'                    'not all cases will be valid in all event code.  If a valid but inproper
'                    'for situation barcode is scanned, either notify user or ignore.
'                    Select Case LCase(bcdReadBarcodeData.BarcodeType)
'                        Case "b"    'Box info barcode
'
'                        Case "c"     'cart position barcode
'
'                        Case "d"    'Disk data barcode
'
'                        Case "r"        'rack position barcode
'
'                        Case "t"        'tube position barcode
'
'                        Case Else       'anything that wasn't one of the above
'
'                    End Select
'
'                Else
'                    'no closing brace found - bad data
''                    Display_Message lstStatus, "Incomplete barcode reading!" & vbCrLf & _
'                                "Input should begin with '{' and end with '}'." & vbCrLf & _
'                                "barcode = " & strDataFromPort
'                    Debug.Print "no close brace"
'                End If
'            Else
'                'no opening brace found - bad data
''                Display_Message lstStatus, "Incomplete barcode reading!" & vbCrLf & _
'                            "Input should begin with '{' and end with '}'." & vbCrLf & _
'                            "barcode = " & strDataFromPort
'                Debug.Print "no open brace"
'            End If
'        Loop
'    End If
'
'
'End Sub
'
'
'
