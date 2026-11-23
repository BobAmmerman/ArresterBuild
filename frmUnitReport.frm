VERSION 5.00
Begin VB.Form frmUnitReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7500
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   450
      TabIndex        =   5
      Top             =   5925
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton btnUpdateReport 
      Caption         =   "TEMP - UPDATE REPORT"
      Height          =   465
      Left            =   2160
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Finish"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11850
      TabIndex        =   0
      Top             =   6000
      Width           =   1050
   End
   Begin VB.PictureBox pictContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   5000
      Left            =   200
      ScaleHeight     =   4935
      ScaleWidth      =   12345
      TabIndex        =   1
      Top             =   500
      Width           =   12408
      Begin VB.PictureBox pictReport 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   20000
         Left            =   0
         ScaleHeight     =   19995
         ScaleWidth      =   12000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   12000
      End
      Begin VB.VScrollBar vscReport 
         Height          =   4905
         LargeChange     =   15
         Left            =   12000
         Max             =   200
         SmallChange     =   3
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmUnitReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_RptLength As Long
Public dbconBuilt As New ADODB.Connection






Private Sub btnNext_Click()

    'this will take the user back to the begining
    frmWorkOrder.Show
    frmWorkOrder.ZOrder 0
    resizeForm
    Unload Me

End Sub

Private Sub btnUpdateReport_Click()
    'temp (probably) for debug
    'globCurrentUnit.SN = "2030800222"
    'globCurrentUnit.UnitNumber = 3
    'globCurrentUnit.Type = 1
    
    DisplayUnitReport
End Sub

Private Sub Command1_Click()

    pictReport.AutoRedraw = True
    pictReport.Cls
    pictReport.Height = 30000
    pictReport.AutoRedraw = True
    
    PrintPolyPicture pictReport
    
    pictReport.Height = m_RptLength
    If pictReport.Height > pictContainer.Height Then
        vscReport.Max = pictReport.Height - pictContainer.Height
        vscReport.Visible = True
        vscReport.SmallChange = vscReport.Max / 16
        vscReport.LargeChange = pictContainer.Height * 0.9
    Else
        vscReport.Visible = False
    End If

End Sub

Private Sub Form_Load()
    Call setDefaultLocation(frmUnitReport)
    Call setCaption("Unit Report for: " & globCurrentUnit.SN & ", Unit " & globCurrentUnit.UnitNumber)
    Call setProgress(4)
    Call moveForm
    
    Set dbconBuilt = dbconResultsData
    DisplayUnitReport
    
    If globAutoPrint = 1 Then
        PrintUnitReport Printer
        Printer.EndDoc
        ' As of 01/23/04, 2 unit reports are printed if the auto print flag is equal to 1
        ' .....requested by Alice Zebracki and Mike Ramarge
        'As of 02/24/04 back to 1 report
        'PrintUnitReport Printer
        'Printer.EndDoc
    End If
    
End Sub

Private Sub Form_Paint()
    resizeForm
End Sub

Private Sub DisplayUnitReport()

    '------------------------------------------------------------------------------------------------
    'Created:   04/16/03    By:Ray Raf
    'Purpose:   causes a Unit Build Report to be printed on the screen.
    '-------------------------------------------------------------------------------------------------
    
    pictReport.AutoRedraw = True
    pictReport.Cls
    pictReport.Height = 30000
    pictReport.AutoRedraw = True
    
    PrintUnitReport pictReport
    
    pictReport.Height = m_RptLength
    If pictReport.Height > pictContainer.Height Then
        vscReport.Max = pictReport.Height - pictContainer.Height
        vscReport.Visible = True
        vscReport.SmallChange = vscReport.Max / 16
        vscReport.LargeChange = pictContainer.Height * 0.9
    Else
        vscReport.Visible = False
    End If


End Sub

Private Sub Form_Resize()
    pictContainer.Width = Me.Width - 600
    pictContainer.Height = Me.Height - 2000
    
    vscReport.Left = pictContainer.ScaleWidth - vscReport.Width
    vscReport.Height = pictContainer.ScaleHeight
    
    pictReport.Width = pictContainer.ScaleWidth - vscReport.Width
    
    
End Sub



Private Sub vscReport_Change()
    pictReport.Top = -vscReport.Value

End Sub

Public Sub PrintUnitReport(PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    '   Created:   04/15/03    By:Ray Raf
    '   Purpose:   prints a Unit Report to the passed object.
    '-------------------------------------------------------------------------------------------------
    Dim rstDisks As New ADODB.Recordset
    Dim rstDiskTypes As New ADODB.Recordset
    'Dim rstBatchSteps As New ADODB.Recordset
    Dim SQL As String
    Dim strLine As String
    Dim strWidthCheck As String
    Dim iWidth As Integer
    Dim iNextLine As Integer
    Dim iLoop As Integer
    Dim iLeftEdge As Integer, iRightEdge As Integer, iStartPos As Integer
    Dim iLineSpacing As Integer
    Dim i1stHeaderLinePos As Integer, i2ndHeaderLinePos As Integer
    Dim i1stColPos As Integer, i2ndColPos As Integer, iNextColPos As Integer
    Dim i1stColData As Integer, i2ndColData As Integer
    Dim iTotalDisks As Integer
    Dim iLastPictureLine As Integer
    Dim iDetailsLastLine As Integer
    Dim strDate As String
    
    iLineSpacing = 250
    
    iNextLine = 500
    With PrintDev
        .Font.Name = "Arial"
        .FontSize = 10
        .FontBold = False
    End With
    
    Select Case globCurrentWO.strType
        Case 1  'Porcelain
            i1stColPos = 500
            i2ndColPos = i1stColPos + 2000
        Case 2
            i1stColPos = 4000
            i2ndColPos = i1stColPos + 2000
            'Polymer arrester has an extra Mounting code to display
            PrintIt i1stColPos, iNextLine, "Arrester Mounting:", PrintDev
            If Len(globCurrentWO.strCatNumber) >= 13 Then
                strLine = Mid(globCurrentWO.strCatNumber, 13, 1)
            Else
                strLine = "N/A"
            End If
            PrintIt i2ndColPos, iNextLine, strLine, PrintDev
            iNextLine = iNextLine + iLineSpacing
    End Select
    
    'print Catalog number , Serial Number and Unit
    PrintIt i1stColPos, iNextLine, "Catalog Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentWO.strCatNumber, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "WO Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentWO.strWO, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Unit:", PrintDev
    PrintIt i2ndColPos, iNextLine, Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    
    'print disk selection info
    PrintIt i1stColPos, iNextLine, "Disk Selection - Unit " & Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Disk Type", PrintDev
    PrintIt i2ndColPos, iNextLine, "Qty", PrintDev
    iNextLine = iNextLine + iLineSpacing

    ' get disk counts from the DB
    SQL = "SELECT DiskType, COUNT(DiskType) as NumTypes "
    SQL = SQL & "FROM ConsumedDisks "
    SQL = SQL & "WHERE SerialNumber = '" & globCurrentUnit.SN & "' "
    SQL = SQL & "AND UnitNumber = " & globCurrentUnit.UnitNumber & " "
    SQL = SQL & "GROUP BY DiskType "

    On Error Resume Next
    rstDiskTypes.Open SQL, dbconBuilt, adOpenStatic, adLockReadOnly
    If Err.Number > 0 Then
        ' an error occured accessing the database
        ErrorLogger "UnitReport", "Error Accessing Consumed Disks" & vbCrLf & Err.Number & _
                                            " " & Err.Description
        MsgBox "Could not access Consumed Disks Data!" & vbCrLf & "Build Report Not Printed!" & _
                        vbCrLf & "See error log for details."
        Exit Sub
    End If
    If rstDiskTypes.EOF = True Then
        ' no disks found for the unit. Shouldn't happen but might.
        PrintIt i1stColPos, iNextLine, "No Disks Found for this Unit!", PrintDev
        iNextLine = iNextLine + (2 * iLineSpacing)
    Else
        ' print disk type and quantity for each group of types
        Do While rstDiskTypes.EOF = False
            PrintIt i1stColPos, iNextLine, rstDiskTypes!DiskType, PrintDev
            PrintIt i2ndColPos, iNextLine, rstDiskTypes!NumTypes, PrintDev
            iNextLine = iNextLine + iLineSpacing
            rstDiskTypes.MoveNext
        Loop
        iNextLine = iNextLine + (2 * iLineSpacing)
    End If
    rstDiskTypes.Close

    PrintIt i1stColPos, iNextLine, "Qty", PrintDev
    PrintIt i2ndColPos, iNextLine, "Category Label", PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    'get disk counts from the DB
    SQL = "SELECT Category, COUNT(Category) as NumDisks "
    SQL = SQL & "FROM ConsumedDisks "
    SQL = SQL & "WHERE SerialNumber = '" & globCurrentUnit.SN & "' "
    SQL = SQL & "AND UnitNumber = " & globCurrentUnit.UnitNumber & " "
    SQL = SQL & "GROUP BY Category "

    On Error Resume Next
    rstDisks.Open SQL, dbconBuilt, adOpenStatic, adLockReadOnly
    If Err.Number > 0 Then
        'an error occured accessing the database
        ErrorLogger "UnitReport", "Error Accessing Consumed Disks" & vbCrLf & Err.Number & _
                                            " " & Err.Description
        MsgBox "Could not access Consumed Disks Data!" & vbCrLf & "Build Report Not Printed!" & _
                        vbCrLf & "See error log for details."
        Exit Sub
    End If
    If rstDisks.EOF = True Then
        'no disks found for the unit. Shouldn't happen but might.
        PrintIt i1stColPos, iNextLine, "No Disks Found for this Unit!", PrintDev
        iNextLine = iNextLine + (2 * iLineSpacing)
    Else
        'print quantity and category for each group of disks
        iTotalDisks = 0
        Do While rstDisks.EOF = False
            PrintIt i1stColPos, iNextLine, rstDisks!NumDisks, PrintDev
            PrintIt i2ndColPos, iNextLine, rstDisks!Category, PrintDev
            iNextLine = iNextLine + iLineSpacing
            iTotalDisks = iTotalDisks + rstDisks!NumDisks
            rstDisks.MoveNext
        Loop
        PrintDev.FontBold = True
        PrintIt i1stColPos, iNextLine, "Total Disks: " & iTotalDisks, PrintDev
        PrintDev.FontBold = False
        iNextLine = iNextLine + (2 * iLineSpacing)
    End If
    rstDisks.Close
    
    If globCurrentWO.strType = 1 Then  ' porcelain
        iNextLine = PrintArresterNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintPorcUnitNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iDetailsLastLine = PrintPorcAssyDetail(PrintDev, 5000, 7000, 500)
        'to properly position build notes, take larger of iNextLine and iDetailsLastLine
        If iDetailsLastLine > iNextLine Then
            iNextLine = iDetailsLastLine
        End If
        iNextLine = PrintPorcBuildNotes(PrintDev, i1stColPos, i2ndColPos, iNextLine)
    Else  ' polymer
        iNextLine = PrintPolyAssyDetail(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = iNextLine + 2 * (iLineSpacing)

        iNextLine = PrintArresterNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintPolyStackedNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintCoverPlateEtching(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintPolyBuildNotes(PrintDev, i1stColPos, i2ndColPos, iNextLine)

' DLimburg Change 2010_1 3Feb2010 Removed picture of component placement
'                                   This information will come with work order from this point forward.
'        iLastPictureLine = PrintPolyPicture(PrintDev)

    End If
    
    iNextLine = iNextLine + 2 * (iLineSpacing)
    
    If iLastPictureLine > iNextLine Then
        m_RptLength = iLastPictureLine + iLineSpacing
    Else
        m_RptLength = iNextLine + iLineSpacing
    End If
    
End Sub

Private Function PrintPorcAssyDetail(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    '   Created:   04/16/03    By:Ray Raf
    '   Purpose:   Print the Assemble details portion of the Unit Build
    '                   Report for Porcelain Arresters.
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical positiont that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim rstHousing As New ADODB.Recordset
    Dim rstDetails As New ADODB.Recordset
    Dim rstNumDisks As New ADODB.Recordset
    Dim rstNotes As New ADODB.Recordset
    Dim SQL As String
    Dim strLine As String
    Dim strWidthCheck As String
    Dim iWidth As Integer
    Dim iLoop As Integer
    Dim iLeftEdge As Integer, iRightEdge As Integer, iStartPos As Integer
    Dim iLineSpacing As Integer
    Dim i1stHeaderLinePos As Integer, i2ndHeaderLinePos As Integer
    Dim i1stColData As Integer, i2ndColData As Integer
    Dim iTotalDisks As Integer
    Dim strUnitField As String
    Dim iNumDisks As Integer
    Dim strBuildTable As String
    Dim strHousing As String
    Dim iTotalStacks As Integer, iStacksTop As Integer, iStacksBottom As Integer
    Dim iBuildNoteLoop As Integer
    Dim strBuildNoteLines() As String
    
    On Error GoTo ErrHandler
    
    iLineSpacing = 250
    'print heading
    strLine = "Arrester Unit Assembly Details"
    PrintDev.FontBold = True
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    PrintDev.FontBold = False
    
    PrintIt i1stColPos, iNextLine, "Porcelain", PrintDev
    PrintIt i2ndColPos, iNextLine, "MCOV/COV", PrintDev
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Housing", PrintDev
    PrintIt i2ndColPos, iNextLine, "(kV)", PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintDev.DrawWidth = 2
    PrintDev.Line (i1stColPos, iNextLine)-(i2ndColPos + 1250, iNextLine)
    PrintDev.DrawWidth = 1
    
    'get housing number
    With globCurrentUnit
        If .Nameplate <> "D" Then
            'normal (or cubicle) arrester
            SQL = "SELECT * FROM HousingPorc WHERE RingCode = '" & .PorcelainRingCode & .HousingCode & "' "
        Else
            'dual rated.
            SQL = "SELECT * FROM HousingDrPorc Where CatalogNumber = '" & globCurrentWO.strCatNumber & "' "
        End If
    End With
    rstHousing.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    strUnitField = "Unit_" & globCurrentUnit.UnitNumber & "_Housing"
    strHousing = rstHousing(strUnitField).Value
    
    'print housing and Mcov
    PrintIt i1stColPos, iNextLine, strHousing, PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.UnitMCOV, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    
    'get total disks for this unit from BuiltUnits table
    SQL = "SELECT * FROM BuiltUnits WHERE SerialNumber ='" & globCurrentUnit.SN & "' "
    SQL = SQL & "AND Unit = " & globCurrentUnit.UnitNumber & " "
    rstNumDisks.Open SQL, dbconResultsData, adOpenDynamic, adLockBatchOptimistic
    iNumDisks = rstNumDisks!NumDisks
    PrintIt i1stColPos, iNextLine, "No. Disks Total", PrintDev
    PrintIt i2ndColPos, iNextLine, iNumDisks, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    'get rest of build information
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "dual rated"
            strBuildTable = "BuildDrPorc"
        Case "cubicle mount"
            strBuildTable = "BuildTableCube_Rev1"   ' Chg'd from BuildTableCube to BuildTableCube_Rev1 DLimburg 2/16/2011
        Case Else
            strBuildTable = "BuildTablePorc_Rev1"   ' Chg'd from BuildTablePorc to BuildTablePorc_Rev1 DLimburg 2/16/2011
    End Select
    SQL = "SELECT * FROM " & strBuildTable & " WHERE "
    SQL = SQL & "DiskType = '" & globCurrentUnit.DiskType & "' "
    SQL = SQL & "AND PorcelainHousing = '" & strHousing & "' "
    SQL = SQL & "AND totDisks = " & iNumDisks & " "
    rstDetails.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    
    'print stack count info
    PrintIt i1stColPos, iNextLine, "No. stacks total", PrintDev
    iTotalStacks = rstDetails!TotDiskStacks
    PrintIt i2ndColPos, iNextLine, iTotalStacks, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "No. stacks Bot.", PrintDev
    iStacksBottom = rstDetails!NumStacksBottom
    PrintIt i2ndColPos, iNextLine, iStacksBottom, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "No. stacks Top.", PrintDev
    iStacksTop = rstDetails!NumStacksTop
    PrintIt i2ndColPos, iNextLine, iStacksTop, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    
    'print all of the stacks info
    For iLoop = iTotalStacks To 1 Step -1
        'start with the top most (highest numbered) stack and work down
        PrintIt i1stColPos, iNextLine, "Disk Stack #" & iLoop, PrintDev
        PrintIt i2ndColPos, iNextLine, Format(rstDetails("Stack" & iLoop & "_NumDisks").Value, "#0"), PrintDev
        iNextLine = iNextLine + iLineSpacing
        PrintIt i1stColPos, iNextLine, "Spacer #" & iLoop & " (in.)", PrintDev
        PrintIt i2ndColPos, iNextLine, Format(rstDetails("Space" & iLoop).Value, "0.000"), PrintDev
        iNextLine = iNextLine + iLineSpacing
        If (iTotalStacks - iLoop) + 1 = iStacksTop Then
            'print a blank line to separate top stacks from bottoms stacks
            iNextLine = iNextLine + iLineSpacing
        End If
    Next iLoop
    
    'iNextLine = iNextLine + (10 * iLineSpacing) '10 lines is temp rev until build notes reworked
    PrintPorcAssyDetail = iNextLine + iLineSpacing
    
    Exit Function
ErrHandler:
    Dim msg$
    
    msg$ = "Error in PrintPorcAssyDetail :  " & vbCrLf
    msg$ = msg$ & "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & Err.Description
    MsgBox msg$, vbOKOnly, "Error"
    
    Call ErrorLogger("frmUnitReport: ", msg$)
    Err.Clear
    PrintPorcAssyDetail = iNextLine + iLineSpacing
    
    Exit Function
    
End Function

Private Function PrintPorcBuildNotes(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer

    '------------------------------------------------------------------------------------------------
    '   Created:   04/16/03    By:Ray Raf
    '   Purpose:   Print the Build Notes portion of the Unit Build
    '                   Report for Porcelain Arresters.
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical positiont that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim SQL As String, strLine As String
    Dim strBuildNoteLines() As String
    Dim rstHousing As New ADODB.Recordset, rstNotes As New ADODB.Recordset
    Dim strUnitField As String, strHousing As String
    Dim strMsg As String
    Dim iBuildNoteLoop As Integer
    
    iLineSpacing = 250
    'print heading
    strLine = "Build Notes"
    PrintDev.FontBold = True
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    PrintDev.FontBold = False
    
    'get housing number
    With globCurrentUnit
        If .Nameplate <> "D" Then
            'normal (or cubicle) arrester
            SQL = "SELECT * FROM HousingPorc WHERE RingCode = '" & .PorcelainRingCode & .HousingCode & "' "
        Else
            'dual rated.
            SQL = "SELECT * FROM HousingDrPorc Where CatalogNumber = '" & globCurrentWO.strCatNumber & "' "
        End If
    End With
    rstHousing.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    strUnitField = "Unit_" & globCurrentUnit.UnitNumber & "_Housing"
    strHousing = rstHousing(strUnitField).Value
    
    'make sure a housing can be found
    If Trim(strHousing) = "" Then
        strMsg = "The housing for this unit could not be found!"
        MsgBox strMsg
        ErrorLogger "PrintPorcBuildNotes", strMsg & " Ring code = " & globCurrentUnit.PorcelainRingCode
        PrintPorcBuildNotes = iNextLine
        Exit Function
    End If
        
    SQL = "SELECT * FROM BuildNotesPorc WHERE HousingCode = '" & strHousing & "' "
    rstNotes.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstNotes.EOF = False Then
        'print the build notes
        ParseBuildNoteToLines rstNotes!BuildNotes, strBuildNoteLines
        For iBuildNoteLoop = 1 To UBound(strBuildNoteLines)
            PrintIt i1stColPos, iNextLine, strBuildNoteLines(iBuildNoteLoop), PrintDev
            iNextLine = iNextLine + iLineSpacing
        Next iBuildNoteLoop
        'PrintIt i1stColPos, iNextLine, rstNotes!BuildNotes, PrintDev
    End If
    PrintPorcBuildNotes = iNextLine + iLineSpacing

    If rstHousing.State > 0 Then
        rstHousing.Close
    End If
    
    If rstNotes.State > 0 Then
        rstNotes.Close
    End If
    
End Function

Private Function PrintPolyBuildNotes(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer

    '------------------------------------------------------------------------------------------------
    '   Created:   01/28/04    By:Pete Wisniewski
    '   Purpose:   Print the Build Notes portion of the Unit Build
    '                   Report for Polymer Arresters.
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical positiont that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim SQL As String, strLine As String
    Dim strBuildNoteLines() As String
    Dim rstNotes As New ADODB.Recordset
    Dim strTestChar As String
    Dim strMsg As String
    Dim iBuildNoteLoop As Integer
    
    ' only print polymer build notes for catalog numbers with "D" in the 13th digit
    strTestChar = UCase(Mid(globCurrentWO.strCatNumber, 13, 1))
    If strTestChar <> "D" Then
        PrintPolyBuildNotes = iNextLine
        Exit Function
    End If
    
    iLineSpacing = 250
    'print heading
    strLine = "Build Notes"
    PrintDev.FontBold = True
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    PrintDev.FontBold = False
    
    
    SQL = "SELECT * FROM BuildNotesPoly WHERE CatalogNumber = '" & globCurrentWO.strCatNumber & "' "
    rstNotes.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstNotes.EOF = False Then
        'print the build notes
        ParseBuildNoteToLines rstNotes!BuildNotes, strBuildNoteLines
        For iBuildNoteLoop = 1 To UBound(strBuildNoteLines)
            PrintIt i1stColPos, iNextLine, strBuildNoteLines(iBuildNoteLoop), PrintDev
            iNextLine = iNextLine + iLineSpacing
        Next iBuildNoteLoop
        'PrintIt i1stColPos, iNextLine, rstNotes!BuildNotes, PrintDev
    Else  ' none found
        PrintIt i1stColPos, iNextLine, "None found.", PrintDev
        iNextLine = iNextLine + iLineSpacing
    End If
    PrintPolyBuildNotes = iNextLine + iLineSpacing
    
    If rstNotes.State > 0 Then
        rstNotes.Close
    End If
    
End Function

Private Function PrintArresterNameplateData(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    ' Arrester Nameplate
    '------------------------------------------------------------------------------------------------
    '   Created:   04/16/03    By:Ray Raf
    '   Purpose:   Print the Assemble details portion of the Unit Build
    '                   Report for Arresters.
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical position that can be printed at.
    '   Rev'd:      05/29/03        By: Ray Raf
    '   Note:       Added logic to use a new cubicle mount nameplate table for porc.
    '   Rev'd:      10/21/03        By: Ray Raf
    '   Note:       Added logic to use a new cubicle mount nameplate table for polymer
    '   Rev'd:      12/19/03    By: Ray Raf
    '   Note:       changed nameplate lookup to use actually family instead of
    '                  Family4Lookup
    '  Rev'd:       01/26/04    By: Pete Wisniewski
    '   Note:       added code to only allow printing Arrester Nameplates for Unit 1
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim rstNameplate As New ADODB.Recordset
    Dim rstTime As New ADODB.Recordset
    Dim SQL As String
    Dim strNameplateTableName As String
    Dim strArresterMCOV  As String
    
    ' only print Arrester Nameplates for Unit 1 (Unit A)
    If globCurrentUnit.UnitNumber <> 1 Then
        PrintArresterNameplateData = iNextLine
        Exit Function
    End If
    
    On Error Resume Next
    
    iLineSpacing = 250
    PrintDev.FontBold = True
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Arrester Nameplate", PrintDev
    PrintDev.FontBold = False
    iNextLine = iNextLine + iLineSpacing
    
    'get info for nameplate
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "cubicle mount"
            strNameplateTableName = "NameplateCubePorc"
            SQL = "SELECT * FROM " & strNameplateTableName & " "
            SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family & "' "
        Case "aze/azg"
            'porelain arresters
            strNameplateTableName = "NameplatePorc"
            SQL = "SELECT * FROM " & strNameplateTableName & " "
            SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family & "' "
            SQL = SQL & "AND MinRating <= " & globCurrentUnit.Rating
            SQL = SQL & " AND MaxRating >= " & globCurrentUnit.Rating & " "
        Case "12mm", "1in"
            'polymer arresters
            If Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "B" Or _
                Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "F" Or _
                Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "G" Or _
                Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "J" Or _
                Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "K" Or _
                Trim(Mid(globCurrentWO.strCatNumber, 13, 1)) = "L" Then
                
                strNameplateTableName = "NameplateCubePoly"
            Else
                strNameplateTableName = "NameplatePoly"
            End If
            SQL = "SELECT * FROM " & strNameplateTableName & " "
            SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family & "' "
            SQL = SQL & "AND MinRating <= " & globCurrentUnit.Rating
            SQL = SQL & " AND MaxRating >= " & globCurrentUnit.Rating & " "
        Case "dual rated"
            'dual rated is always a special case
            If globCurrentWO.strType = 1 Then
                strNameplateTableName = "NameplateDrPorc"
            Else
                strNameplateTableName = "NameplateDrPoly"
            End If
            SQL = "SELECT * FROM " & strNameplateTableName & " "
            SQL = SQL & "WHERE CatalogNumber = '" & globCurrentWO.strCatNumber & "' "
    End Select
    
    rstNameplate.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    

    PrintIt i1stColPos, iNextLine, "Catalog Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentWO.strCatNumber, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Year of Manuf.:", PrintDev
    
    ' rstTime is used to get the year of manufacture for the report
    SQL = "SELECT * FROM BuiltUnits "
    SQL = SQL & "WHERE SerialNumber = '" & globCurrentUnit.SN & "' "
    rstTime.Open SQL, dbconResultsData, adOpenStatic, adLockReadOnly
    
    If rstTime.EOF = False Then  ' record exists
        PrintIt i2ndColPos, iNextLine, Format(rstTime!TimeBuilt, "yy"), PrintDev
    Else
        PrintIt i2ndColPos, iNextLine, "Unknown", PrintDev
    End If
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Rating:", PrintDev
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "dual rated"
            PrintIt i2ndColPos, iNextLine, Trim(CStr(rstNameplate!ArresterRating)), PrintDev
        Case Else
            PrintIt i2ndColPos, iNextLine, Trim(CStr(globCurrentUnit.Rating)), PrintDev
    End Select
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "MCOV/COV:", PrintDev
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "dual rated"
            PrintIt i2ndColPos, iNextLine, rstNameplate!MCOV_Arrester, PrintDev
        Case Else
            strArresterMCOV = LookupDesignMCOV(globCurrentUnit.MCOV)
    End Select
    PrintIt i2ndColPos, iNextLine, strArresterMCOV, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Pressure Relief:", PrintDev
    PrintIt i2ndColPos, iNextLine, Trim(CStr(rstNameplate!PressureRelief)), PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Class:", PrintDev
    PrintIt i2ndColPos, iNextLine, rstNameplate!Class, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Certification:", PrintDev
    PrintIt i2ndColPos, iNextLine, rstNameplate!Certification, PrintDev
    iNextLine = iNextLine + iLineSpacing
        
    If rstNameplate.State > 0 Then rstNameplate.Close
    If rstTime.State > 0 Then rstTime.Close
    
    PrintArresterNameplateData = iNextLine
    
End Function

Private Function PrintPolyStackedNameplateData(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    ' Stacked Nameplate  -- Only pertains to Polymer Arresters
    '------------------------------------------------------------------------------------------------
    '   Created:   01/26/04    By:Pete Wisniewski
    '   Purpose:   Print a nameplate with MCOV for each unit in arrester...
    '                    this nameplate is used for arrester assembly by customer
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical position that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim rstNameplate As New ADODB.Recordset
    Dim SQL As String
    Dim strMCOV_TableName As String
    Dim strCompTableName As String
    Dim X As Integer
    Dim strMCOV As String
    Dim i3rdColPos As Integer
    Dim iShedCode As Integer
    
    ' only print Stacked Nameplates for Unit 1 (Unit A)
    If globCurrentUnit.UnitNumber <> 1 Then
        PrintPolyStackedNameplateData = iNextLine
        Exit Function
    End If
    
    ' only print Stacked Nameplates when total sheds >= 62
    'if shed code is not numeric, it represents more that 99 sheds
    If IsNumeric(globCurrentUnit.PolyShedRingCode) Then
        iShedCode = CInt(globCurrentUnit.PolyShedRingCode)
        If iShedCode < 62 Then
            PrintPolyStackedNameplateData = iNextLine
            Exit Function
        End If
    End If
    
    i3rdColPos = i2ndColPos + 500
    
    iLineSpacing = 250
    PrintDev.FontBold = True
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Stacked Nameplate", PrintDev
    PrintDev.FontBold = False
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Catalog Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentWO.strCatNumber, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    'get info for stacked nameplate
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "12mm", "1in"
            'polymer arresters
            GetPolyTableNames globCurrentUnit, strCompTableName, strMCOV_TableName
            SQL = "SELECT * FROM " & strMCOV_TableName & " "
            SQL = SQL & "WHERE Family='" & globCurrentUnit.Family & "' "
            SQL = SQL & "AND String_MCOV='" & globCurrentUnit.MCOV & "' "
            SQL = SQL & "AND ShedCode='" & globCurrentUnit.PolyShedRingCode & "' "
            SQL = SQL & "AND ArresterRating = " & globCurrentUnit.Rating & " "
                    
            Set rstNameplate = dbconSetupData.Execute(SQL)
            If Not rstNameplate.EOF Then
                'arrester was found.  Get Unit data
                For X = 6 To 1 Step -1
                    strMCOV = rstNameplate.Fields("MCOV_Unit_" & X)
                    If strMCOV <> "0" Then
                        PrintIt i1stColPos, iNextLine, "Unit " & Chr(X + 64) & ":", PrintDev
                        PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                        PrintIt i3rdColPos, iNextLine, "kV rms", PrintDev
                        iNextLine = iNextLine + iLineSpacing
                        strMCOV = "0"
                    End If
                Next
            End If
        
        Case "dual rated"
            SQL = "SELECT * FROM MCOV_Dr_Poly "
            SQL = SQL & "WHERE CatalogNumber = '" & globCurrentWO.strCatNumber & "' ;"
            Set rstNameplate = dbconSetupData.Execute(SQL)
            If Not rstNameplate.EOF Then
                strMCOV = rstNameplate.Fields("MCOV_Unit_2_Lower")
                If strMCOV <> "0" Then
                    PrintIt i1stColPos, iNextLine, "Unit B:", PrintDev
                    PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                    PrintIt i3rdColPos, iNextLine, "kV rms", PrintDev
                    iNextLine = iNextLine + iLineSpacing
                    strMCOV = "0"
                End If
                strMCOV = rstNameplate.Fields("MCOV_Unit_1")
                If strMCOV <> "0" Then
                    PrintIt i1stColPos, iNextLine, "Unit A:", PrintDev
                    PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                    PrintIt i3rdColPos, iNextLine, "kV rms", PrintDev
                    iNextLine = iNextLine + iLineSpacing
                End If
            End If
    End Select
    
    If rstNameplate.State > 0 Then rstNameplate.Close
    
    PrintPolyStackedNameplateData = iNextLine
    
End Function

Private Function PrintPorcUnitNameplateData(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    ' Porc Unit Nameplate  -- Only pertains to Porcelain Arrester Units
    '------------------------------------------------------------------------------------------------
    '   Created:   01/26/04    By:Pete Wisniewski
    '   Purpose:   Print a unit nameplate
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical position that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim rstNameplate As New ADODB.Recordset
    Dim SQL As String
    Dim strMCOV_TableName As String
    Dim strCompTableName As String
    Dim X As Integer
    Dim strMCOV As String
    Dim i3rdColPos As Integer
    Dim iNumOfUnits As Integer
    Dim strMCOVTable As String
    i3rdColPos = i2ndColPos + 500
    iNumOfUnits = 0
    
    iLineSpacing = 250
    PrintDev.FontBold = True
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Unit Nameplate", PrintDev
    PrintDev.FontBold = False
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Catalog Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentWO.strCatNumber, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN & Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    'get info for stacked nameplate
    Select Case LCase(globCurrentWO.strBuildPath)
        Case "dual rated"
            'dual rated arresters ALWAYS have exactly two units.
            SQL = "SELECT * FROM MCOV_Dr_Porc "
            SQL = SQL & "WHERE CatalogNumber='" & globCurrentWO.strCatNumber & "' "
            Set rstNameplate = dbconSetupData.Execute(SQL)
            If Not rstNameplate.EOF Then
                'the catalog number was found in the table
                If globCurrentUnit.UnitNumber = 1 Then
                    strMCOV = rstNameplate.Fields("MCOV_Unit_1")
                Else
                    strMCOV = rstNameplate.Fields("MCOV_Unit_2_Lower")
                End If
                PrintIt i1stColPos, iNextLine, "MCOV/COV:", PrintDev
                PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                iNextLine = iNextLine + iLineSpacing
                
                iNumOfUnits = 2
            End If
        Case "cubicle mount"
            'cubicle mount can have up to 3 units
            With globCurrentUnit
                SQL = "SELECT * FROM MCOV_Cube "
                SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                SQL = SQL & "AND String_MCOV ='" & .MCOV & "' "
                SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                Set rstNameplate = dbconSetupData.Execute(SQL)
            End With
            
            If Not rstNameplate.EOF Then
                For X = 1 To 3
                    strMCOV = rstNameplate.Fields("MCOV_Unit_" & X)
                    If strMCOV <> "0" Then
                        iNumOfUnits = iNumOfUnits + 1
                    End If
                Next X
                    
                strMCOV = rstNameplate.Fields("MCOV_Unit_" & globCurrentUnit.UnitNumber)
                PrintIt i1stColPos, iNextLine, "MCOV/COV:", PrintDev
                PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                iNextLine = iNextLine + iLineSpacing
            End If
            
        Case Else
        ' added test for calcined disk builds - 14Apr2011
            If globCalcinedDisksUsedInPorc = True Then
                strMCOVTable = "MCOV_Porc_Rev1"
            Else
                strMCOVTable = "MCOV_Porc"
            End If
            'all others can have up to 4 units

            With globCurrentUnit
                SQL = "SELECT * FROM " & strMCOVTable & " "
                SQL = SQL & "WHERE Family='" & .Family4Lookup & "' "
                SQL = SQL & "AND ArresterRating='" & .Rating & "' "
                SQL = SQL & "AND string_MCOV='" & .MCOV & "' "
                SQL = SQL & "AND PorcRingNumber='" & .PorcelainRingCode & "' ;"
                Set rstNameplate = dbconSetupData.Execute(SQL)
            End With
            If Not rstNameplate.EOF Then
                For X = 1 To 4
                    strMCOV = rstNameplate.Fields("MCOV_Unit_" & X)
                    If strMCOV <> "0" Then
                        iNumOfUnits = iNumOfUnits + 1
                    End If
                Next X
                
                strMCOV = rstNameplate.Fields("MCOV_Unit_" & globCurrentUnit.UnitNumber)
                PrintIt i1stColPos, iNextLine, "MCOV/COV:", PrintDev
                PrintIt i2ndColPos, iNextLine, strMCOV, PrintDev
                iNextLine = iNextLine + iLineSpacing
            End If
    
    End Select
    
    PrintIt i1stColPos, iNextLine, "Unit " & globCurrentUnit.UnitNumber & " of " & iNumOfUnits, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Unit Position:", PrintDev
    PrintIt i2ndColPos, iNextLine, PorcUnitPosition(globCurrentUnit.UnitNumber, iNumOfUnits), PrintDev
    iNextLine = iNextLine + iLineSpacing

    
    If rstNameplate.State > 0 Then rstNameplate.Close
    
    PrintPorcUnitNameplateData = iNextLine + iLineSpacing
    
End Function

Private Function PrintCoverPlateEtching(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    ' Cover Plate Etching  -- Only pertains to Polymer Arresters
    '------------------------------------------------------------------------------------------------
    '   Created:   01/26/04    By:Pete Wisniewski
    '   Purpose:   Print a nameplate with SN and Unit
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical position that can be printed at.
    '   Rev'd:      05/18/04    By: Ray Raf
    '   Note:       Cover plate should only be printed for arresters with 60
    '                   or more sheds.  This means that it would only apply to
    '                   1 in arresters with 3 or more units.
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim strCompTable As String, strMCOVTable As String
    Dim SQL As String
    Dim rstMCOV As New ADODB.Recordset
    Dim iLoop As Integer
    Dim strMCOV As String
    
    If globCurrentUnit.Stud <> "1in" Then
        'not a 1 inch stud unit, can't be 60 sheds, don't print.
        PrintCoverPlateEtching = iNextLine
        Exit Function
    End If
    
    GetPolyTableNames globCurrentUnit, strCompTable, strMCOVTable
    SQL = "SELECT * FROM " & strMCOVTable & " "
    SQL = SQL & "WHERE Family='" & globCurrentUnit.Family & "' "
    SQL = SQL & "AND String_MCOV='" & globCurrentUnit.MCOV & "' "
    SQL = SQL & "AND ShedCode='" & globCurrentUnit.PolyShedRingCode & "' "
    SQL = SQL & "AND ArresterRating = " & globCurrentUnit.Rating & " "
    Set rstMCOV = dbconSetupData.Execute(SQL)
    If Not rstMCOV.EOF Then
        'arrester was found.  Get Unit data
        strMCOV = rstMCOV.Fields("MCOV_Unit_3")
        If strMCOV = "0" Then
            '2 or 1 unit, can't be more than 60 sheds, don't print.
            PrintCoverPlateEtching = iNextLine
            Exit Function
        End If
    Else
        'no data found, exit
        ErrorLogger "PrintCoverPlateEtching", "Could not find MCOV info for " & SQL
        PrintCoverPlateEtching = iNextLine
        Exit Function
    End If
    
    iLineSpacing = 250
    PrintDev.FontBold = True
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Cover Plate Etching", PrintDev
    PrintDev.FontBold = False
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Unit:", PrintDev
    PrintIt i2ndColPos, iNextLine, Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintCoverPlateEtching = iNextLine
    
End Function

Private Function PrintPolyAssyDetail(PrintDev As Object, i1stColPos As Integer, _
                                    i2ndColPos As Integer, iNextLine As Integer) As Integer
    '------------------------------------------------------------------------------------------------
    '   Created:   04/16/03    By:Ray Raf
    '   Purpose:   Print the Assemble details portion of the Unit Build
    '                   Report for Polymer Arresters.
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '                 i1stColPos - the position for the left data column
    '                 i2ndColPos - the position for the right data column
    '                 iNextLine - the veritcal page position to print at.
    '   Returns:    The next vertical positiont that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim rstParts As New ADODB.Recordset
    Dim rstCompCount As New ADODB.Recordset
    Dim SQL As String
    Dim strLine As String
    Dim strWidthCheck As String
    Dim iWidth As Integer
    Dim iLoop As Integer
    Dim iLineSpacing As Integer
    Dim strBuildTable As String
    Dim strHousing As String
    Dim strComponentsTable As String, strMCOVTable As String
    Dim strPartsTable As String
    Dim AZ298BTest As Boolean ' DLimburg Change 2009_5 (December 1,2009) Polymer product defined by AZ298B can not be defined on build sheet.
    
    iLineSpacing = 250

    On Error GoTo ErrorHandler
    
    'print heading
    strLine = "Arrester Unit Assembly Details"
    PrintDev.FontBold = True
    PrintItCenter i2ndColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (iLineSpacing)
    PrintDev.FontBold = False
    
    strLine = "Shed"
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    strLine = "MCOV/COV"
    PrintIt i2ndColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + iLineSpacing

    strLine = "Sections"
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    strLine = "(kV)"
    PrintIt i2ndColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + iLineSpacing
    PrintDev.Line (i1stColPos, iNextLine)-Step(750, 0)
    PrintDev.Line (i2ndColPos, iNextLine)-Step(350, 0)

    'show the data
    strLine = globCurrentUnit.UnitSections
    PrintIt i1stColPos, iNextLine, strLine, PrintDev
    strLine = globCurrentUnit.UnitMCOV
    PrintIt i2ndColPos, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    PrintPolyAssyDetail = iNextLine + iLineSpacing
Exit Function

' DLimburg Change 2010_1 3Feb2010 Next section of code no loger used by Arrester Build program
'                                 This information will be provided on a sheet attached to the workorder
    '------Components area----
    strPartsTable = "PartNumbers_" & globCurrentUnit.Process
    SQL = "SELECT * FROM " & strPartsTable & " "
    SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family & "' "
    SQL = SQL & "AND Stud = '" & globCurrentUnit.Stud & "' "
    SQL = SQL & "AND MinRating <= " & globCurrentUnit.Rating
    SQL = SQL & " AND MaxRating >= " & globCurrentUnit.Rating & " "
    rstParts.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    
    'get piece counts from proper component table
    GetPolyTableNames globCurrentUnit, strComponentsTable, strMCOVTable
'    Select Case LCase(globCurrentUnit.Stud)
'        Case "12mm"
'            strComponentsTable = "Components_12mm"
'        Case "1 in"
'            strComponentsTable = "Components_1in"
'    End Select
    
'    SQL = "SELECT * FROM " & strComponentsTable & " "
'    SQL = SQL & "WHERE DiskType = '" & globCurrentUnit.DiskType & "' "
'    SQL = SQL & "AND DiskQuantity = " & globNumDisksToUse & " "
'    SQL = SQL & "AND Sections = " & globCurrentUnit.UnitSections & " "
    
' ' DLimburg Change 2009_1 20Feb09 modified above code to accommodate the use of 3kV and 5kV disks
    SQL = "SELECT * FROM " & strComponentsTable & " "
    SQL = SQL & "WHERE DiskType = '" & globCurrentUnit.DiskType & "' "
    SQL = SQL & "AND DiskQuantity_3kV = " & globCurrentUnit.NumOf3kVDisks & " "
    SQL = SQL & "AND DiskQuantity_5kV = " & globCurrentUnit.NumOf5kVDisks & " "
    SQL = SQL & "AND Sections = " & globCurrentUnit.UnitSections & " "
    
    
    rstCompCount.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
        
' 2009_5 start
    AZ298BTest = TestCatalogNumber(globCurrentWO.strCatNumber)
    If AZ298BTest = False Then
' 2009_5 end, continued at else
        strLine = "Component"
        PrintIt i1stColPos, iNextLine, strLine, PrintDev
        strLine = "Qty"
        PrintIt i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + iLineSpacing
        PrintDev.Line (i1stColPos, iNextLine)-(i1stColPos + 1000, iNextLine)
        PrintDev.Line (i2ndColPos, iNextLine)-(i2ndColPos + 300, iNextLine)
    
        'end Electrode
        strLine = rstParts!EndElectrode
        PrintIt i1stColPos, iNextLine, strLine, PrintDev
        globCurrentUnit.UnitEndElectrodeCount = rstCompCount!EndElectrodes
        strLine = rstCompCount!EndElectrodes
        PrintIt i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + iLineSpacing
        
        'Middle Electrode
        strLine = rstParts!MidElectrode
        PrintIt i1stColPos, iNextLine, strLine, PrintDev
        strLine = rstCompCount!midElectrodes
        globCurrentUnit.UnitMidElectrodeCount = rstCompCount!midElectrodes
        PrintIt i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + iLineSpacing
        
        'Heatsink
        strLine = rstParts!HeatSink
        PrintIt i1stColPos, iNextLine, strLine, PrintDev
        strLine = rstCompCount!HeatSinks
        globCurrentUnit.UnitHeatSinkCount = rstCompCount!HeatSinks
    

        PrintIt i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + iLineSpacing
        
        'now show the variable stuff
        If LCase(globCurrentUnit.Process) = "new" Or globCurrentUnit.Stud = "1in" Then
            'show the additional items for the new process
            
            'Pre Preg
            strLine = rstParts!PrePreg
            PrintIt i1stColPos, iNextLine, strLine, PrintDev
            strLine = Format(rstCompCount!PrePreg, "0.000")
            PrintIt i2ndColPos, iNextLine, strLine, PrintDev
            iNextLine = iNextLine + iLineSpacing
            
            'Scrim
            strLine = rstParts!Scrim
            PrintIt i1stColPos, iNextLine, strLine, PrintDev
            strLine = Format(rstCompCount!Scrim, "0.000")
            PrintIt i2ndColPos, iNextLine, strLine, PrintDev
            iNextLine = iNextLine + iLineSpacing
            
            'shrink tape
            strLine = rstParts!Shrink
            PrintIt i1stColPos, iNextLine, strLine, PrintDev
            strLine = Format(rstCompCount!ShrinkTape, "0.000")
            PrintIt i2ndColPos, iNextLine, strLine, PrintDev
            iNextLine = iNextLine + iLineSpacing
            
            'no ORings needed
            globCurrentUnit.UnitORingCount = 0
        Else
            'show orings for the old process
            'ORing
            strLine = rstParts!ORing
            PrintIt i1stColPos, iNextLine, strLine, PrintDev
            strLine = rstCompCount!ORings
            globCurrentUnit.UnitORingCount = rstCompCount!ORings
            PrintIt i2ndColPos, iNextLine, strLine, PrintDev
            iNextLine = iNextLine + iLineSpacing
        End If
    Else ' 2009_5 cont
        'end Electrode
        globCurrentUnit.UnitEndElectrodeCount = rstCompCount!EndElectrodes
        'Middle Electrode
        globCurrentUnit.UnitMidElectrodeCount = rstCompCount!midElectrodes
        'Heatsink
        globCurrentUnit.UnitHeatSinkCount = rstCompCount!HeatSinks
        
        strLine = "Component"
        PrintIt i1stColPos, iNextLine, strLine, PrintDev
        strLine = "Qty"
        PrintIt i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + iLineSpacing
        PrintDev.Line (i1stColPos, iNextLine)-(i1stColPos + 1000, iNextLine)
        PrintDev.Line (i2ndColPos, iNextLine)-(i2ndColPos + 300, iNextLine)
        
        strLine = "Ref Work Order for part numbers and quantities."
        PrintDev.FontBold = True
        PrintItCenter i2ndColPos, iNextLine, strLine, PrintDev
        iNextLine = iNextLine + (iLineSpacing)
        PrintDev.FontBold = False
    End If '2009_5 end
    
    PrintPolyAssyDetail = iNextLine + iLineSpacing
    
    Exit Function
ErrorHandler:
    
    'if we get here there was an unexpected error in the data somewhere
    MsgBox "unexpected error during report print.  See error log for details"
    ErrorLogger "PrintPolyAssyDetails", Err.Description
    Err.Clear
    PrintPolyAssyDetail = iNextLine + iLineSpacing
    
End Function

Private Function PrintPolyPicture(PrintDev As Object) As Integer
    '------------------------------------------------------------------------------------------------
    '   Created:   05/20/03    By:Ray Raf
    '   Purpose:   Print the polymer parts layout portion of the
    '                   Unit Build Report
    '   Inputs:     PrintDev - The object to print to: the Printer or a picture box
    '   Returns:    The next vertical positiont that can be printed at.
    '-------------------------------------------------------------------------------------------------
    Dim varParts() As Variant
    Dim iCenter As Integer
    Dim iLineSpacing As Integer
    Dim iNextLine As Integer
    Dim iLoop As Integer
    Dim strLine As String
    Dim sglMaxWidth As Single
    Dim sglThisWidth As Single
    Dim iTop As Integer
    Dim iLinePos As Integer
    Dim bUseOrings As Boolean
    Dim iOldFontSize As Integer
    
    iCenter = 2000
    iTop = 500
    iNextLine = iTop
    iLineSpacing = 250
    
    iOldFontSize = PrintDev.FontSize
    PrintDev.FontSize = 8
    
    strLine = globCurrentUnit.UnitSections & " Shed"
    PrintItCenter iCenter, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    strLine = globNumDisksToUse & " Disk(s)"
    PrintItCenter iCenter, iNextLine, strLine, PrintDev
    iNextLine = iNextLine + (iLineSpacing * 1.5)
    iTop = iNextLine
    
    With globCurrentUnit
    ' DLimburg 23Feb2009 Added NumberOf5kVSisks
        varParts = CreateModBuild(.UnitSections, .UnitHeatSinkCount, _
                            .UnitMidElectrodeCount, .NumOfDisks, _
                            .NumOf5kVDisks, .UnitORingCount)
    End With
    
    'loop thru each part entry returned and display it
    For iLoop = 0 To UBound(varParts)
        strLine = varParts(iLoop)
        PrintItCenter iCenter, iNextLine, strLine, PrintDev
        'get the width of the string just displayed
        sglThisWidth = PrintDev.TextWidth(strLine)
        If sglThisWidth > sglMaxWidth Then
            'store the width of the text if it is the greatest so far
            sglMaxWidth = sglThisWidth
        End If
        iNextLine = iNextLine + iLineSpacing
    Next iLoop
    
    'draw box
    iNextLine = iNextLine - iLineSpacing
    sglMaxWidth = sglMaxWidth + 100
    PrintDev.Line (iCenter - (sglMaxWidth / 2), iTop)-(iCenter + (sglMaxWidth / 2), iNextLine), vbBlack, B
    
    'draw dividing lines
    iLinePos = iTop + iLineSpacing - 25
    Do While iLinePos < iNextLine
        PrintDev.Line (iCenter - (sglMaxWidth / 2), iLinePos)-(iCenter + (sglMaxWidth / 2), iLinePos), vbBlack, B
        iLinePos = iLinePos + iLineSpacing
    Loop
    
    PrintPolyPicture = iNextLine + iLineSpacing

    PrintDev.FontSize = iOldFontSize
    
End Function

Public Sub ParseBuildNoteToLines(ByVal strNote As String, ByRef strLines() As String)
    '------------------------------------------------------------------------------------------------
    '   Created:   10/31/03   By:Ray Raf
    '   Purpose:   Parse the single field Build note into distict lines for printing.
    '                   This routine ASSUMES! that lines are separated with only a LF
    '   Inputs:     strNote - the text to parse into lines
    '   Outputs:   strLines() - an array containing the individual lines
    '-------------------------------------------------------------------------------------------------
        Dim iNotePos As Integer
        Dim iNextNotePos As Integer
        Dim strNoteLines() As String
        Dim iNoteLine As Integer
        Dim str4Char As String
        
        iNoteLine = 1
        iNotePos = 1
        iNextNotePos = InStr(1, strNote, vbLf)
        
        Do While iNextNotePos > 0
            ReDim Preserve strNoteLines(iNoteLine)
            strNoteLines(iNoteLine) = Mid(strNote, iNotePos, iNextNotePos - iNotePos)
            str4Char = Asc(Mid(strNote, iNextNotePos - 1, 1))
            str4Char = str4Char & "," & Asc(Mid(strNote, iNextNotePos, 1))
            str4Char = str4Char & "," & Asc(Mid(strNote, iNextNotePos + 1, 1))
            str4Char = str4Char & "," & Asc(Mid(strNote, iNextNotePos + 2, 1))
'            Debug.Print "Line #" & iNoteLine & " " & str4Char
'            Debug.Print strNoteLines(iNoteLine)
            iNotePos = iNextNotePos + 1
            iNoteLine = iNoteLine + 1
            iNextNotePos = InStr(iNotePos, strNote, vbLf)
        Loop
        
        'get last line
        ReDim Preserve strNoteLines(iNoteLine)
        strNoteLines(iNoteLine) = Mid(strNote, iNotePos, Len(strNote) - iNotePos + 1)

        'return the array of lines
        strLines = strNoteLines
        
End Sub

Public Function PorcUnitPosition(iPos As Integer, iTotal As Integer) As String
'--------------------------------------------------------------------------------------------------------------
'   This function returns a descriptive string of the position of the unit (iPos)
'   relative to the total (iTotal) number of units in an arrester.
'--------------------------------------------------------------------------------------------------------------

    PorcUnitPosition = "unknown"
    
    ' there are only 4 units max in porc arresters
    Select Case iPos
        Case 1: PorcUnitPosition = "Bot."
        Case 2
            Select Case iTotal
                Case 2: PorcUnitPosition = "Top"
                Case 3, 4: PorcUnitPosition = "2nd"
            End Select
        Case 3
            Select Case iTotal
                Case 3: PorcUnitPosition = "Top"
                Case 4: PorcUnitPosition = "3rd"
            End Select
        Case 4: PorcUnitPosition = "Top"
    End Select
End Function
