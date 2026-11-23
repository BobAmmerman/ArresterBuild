Attribute VB_Name = "modPrintFunctions"
Option Explicit
'variables needed for printer selection form
Global globSelectedPrinter
Global globDefaultReportPrinter
Global globSelectPrintResult

Public Sub PrintIt(LMarg As Integer, TMarg As Integer, InptStr As Variant, PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    'Created:   11/20/00    By:Ray Raf
    'Purpose:   print a string to a passed location on the page
    'Inputs:    LMarg - Left position of text
    '              TMarg - Top Position of Text
    '              InptStr - String to print
    '              PrintDev - either a printdev object or a PictureBox object that can be drawn on.
    '-------------------------------------------------------------------------------------------------
    
    PrintDev.CurrentX = LMarg
    PrintDev.CurrentY = TMarg
    PrintDev.Print InptStr
    
End Sub

Public Sub PrintItRight(iRightEdge As Integer, iTopEdge As Integer, strLine As String, PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    'Created:   03/05/02    By:Ray Raf
    'Purpose:   print a string right aligned to a passed location on the page
    'Inputs:    iRightEdge - the right side edge to align to
    '              iTopEdge- Top Position of Text
    '              strLine - String to print
    '              PrintDev - either a printdev object or a PictureBox object that can be drawn on.
    '-------------------------------------------------------------------------------------------------
    Dim iWidth As Integer, iStartPos As Integer
    
    iWidth = PrintDev.TextWidth(strLine)
    iStartPos = iRightEdge - iWidth
    PrintIt iStartPos, iTopEdge, strLine, PrintDev

End Sub

Public Sub PrintItCenter(iCenter As Integer, iTopEdge As Integer, strLine As String, PrintDev As Object)
    '------------------------------------------------------------------------------------------------
    'Created:   03/05/02    By:Ray Raf
    'Purpose:   print a string Center aligned to a passed location on the page
    'Inputs:    iCenter - the desired center of the text
    '              iTopEdge- Top Position of Text
    '              strLine - String to print
    '              PrintDev - either a printdev object or a PictureBox object that can be drawn on.
    '-------------------------------------------------------------------------------------------------
    Dim iWidth As Integer, iStartPos As Integer
    
    iWidth = PrintDev.TextWidth(strLine)
    iStartPos = iCenter - (iWidth / 2)
    PrintIt iStartPos, iTopEdge, strLine, PrintDev

End Sub

