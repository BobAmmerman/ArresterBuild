VERSION 5.00
Begin VB.Form frmUnitReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6552
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13116
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6552
   ScaleWidth      =   13116
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
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
      Left            =   1500
      TabIndex        =   4
      Top             =   5850
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "Finish"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
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
      ScaleHeight     =   4956
      ScaleWidth      =   12360
      TabIndex        =   1
      Top             =   500
      Width           =   12408
      Begin VB.PictureBox pictReport 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   20000
         Left            =   0
         ScaleHeight     =   20004
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
Private dbconBuilt As New ADODB.Connection


Private Sub btnNext_Click()

    'this will take the user back to the begining
    MsgBox "This unit is Complete!"
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

Private Sub PrintUnitReport(PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    '   Created:   04/15/03    By:Ray Raf
    '   Purpose:   prints a Unit Report to the passed object.
    '-------------------------------------------------------------------------------------------------
    Dim rstDisks As New ADODB.Recordset
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
    
    PrintIt i1stColPos, iNextLine, "Serial Number:", PrintDev
    PrintIt i2ndColPos, iNextLine, globCurrentUnit.SN, PrintDev
    iNextLine = iNextLine + iLineSpacing
    
    PrintIt i1stColPos, iNextLine, "Unit:", PrintDev
    PrintIt i2ndColPos, iNextLine, Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    
    'print disk selection info
    PrintIt i1stColPos, iNextLine, "Disk Selection - Unit " & Chr(globCurrentUnit.UnitNumber + 64), PrintDev
    iNextLine = iNextLine + iLineSpacing
    PrintIt i1stColPos, iNextLine, "Disk Type: " & globCurrentUnit.DiskType, PrintDev
    iNextLine = iNextLine + (2 * iLineSpacing)
    
    PrintIt i1stColPos, iNextLine, "Quantity", PrintDev
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
    
    If globCurrentWO.strType = 1 Then
        iNextLine = PrintPorcAssyDetail(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
    Else
        iNextLine = PrintPolyAssyDetail(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iNextLine = PrintNameplateData(PrintDev, i1stColPos, i2ndColPos, iNextLine)
        iLastPictureLine = PrintPolyPicture(PrintDev)
    End If
    
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
            strBuildTable = "BuildTableCube"
        Case Else
            strBuildTable = "BuildTablePorc"
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
    
    iNextLine = iNextLine + iLineSpacing
    SQL = "SELECT * FROM BuildNotesPorc WHERE HousingCode = '" & strHousing & "' "
    rstNotes.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    If rstNotes.EOF = False Then
        'print the buld notes
        '------temp trial code
        ParseBuildNoteToLines rstNotes!BuildNotes, strBuildNoteLines
        For iBuildNoteLoop = 1 To UBound(strBuildNoteLines)
            PrintIt i1stColPos, iNextLine, strBuildNoteLines(iBuildNoteLoop), PrintDev
            iNextLine = iNextLine + iLineSpacing
        Next iBuildNoteLoop
        '---------------------end trial code
        'PrintIt i1stColPos, iNextLine, rstNotes!BuildNotes, PrintDev
    End If
    iNextLine = iNextLine + iLineSpacing
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

Private Function PrintNameplateData(PrintDev As Object, i1stColPos As Integer, _
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
    '   Rev'd:      05/29/03        By: Ray Raf
    '   Note:       Added logic to use a new cubicle mount nameplate table for porc.
    '   Rev'd:      10/21/03        By: Ray Raf
    '   Note:       Added logic to use a new cubicle mount nameplate table for polymer
    '-------------------------------------------------------------------------------------------------
    Dim iLineSpacing As Integer
    Dim rstNameplate As New ADODB.Recordset
    Dim SQL As String
    Dim strNameplateTableName As String
    Dim strArresterMCOV  As String
    
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
            SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family4Lookup & "' "
        Case "aze/azg"
            'porelain arresters
            strNameplateTableName = "NameplatePorc"
            SQL = "SELECT * FROM " & strNameplateTableName & " "
            SQL = SQL & "WHERE Family = '" & globCurrentUnit.Family4Lookup & "' "
            SQL = SQL & "AND MinRating <= " & globCurrentUnit.Rating
            SQL = SQL & " AND MaxRating >= " & globCurrentUnit.Rating & " "
        Case "12mm", "1in"
            'polymer arresters
            If Mid(globCurrentWO.strCatNumber, 13, 1) = "B" Then
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
    PrintIt i2ndColPos, iNextLine, Format(Now(), "yy"), PrintDev
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
    
    PrintNameplateData = iNextLine
    
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
    
    SQL = "SELECT * FROM " & strComponentsTable & " "
    SQL = SQL & "WHERE DiskType = '" & globCurrentUnit.DiskType & "' "
    SQL = SQL & "AND DiskQuantity = " & globNumDisksToUse & " "
    SQL = SQL & "AND Sections = " & globCurrentUnit.UnitSections & " "
    rstCompCount.Open SQL, dbconSetupData, adOpenDynamic, adLockBatchOptimistic
    
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
        varParts = CreateModBuild(.UnitSections, .UnitHeatSinkCount, _
                            .UnitMidElectrodeCount, globNumDisksToUse, _
                            .UnitORingCount)
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
            Debug.Print "Line #" & iNoteLine & " " & str4Char
            Debug.Print strNoteLines(iNoteLine)
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
