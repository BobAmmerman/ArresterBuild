Attribute VB_Name = "modBarGraphFunctions"
Option Explicit

Type POINTAPI
    x As Long
    Y As Long
End Type

Type POINTDOUBLE
    x As Double
    Y As Double
End Type

Type TWOPOINT
    Pt(2) As POINTAPI
End Type

Type MINMAX
    Min As Double
    Max As Double
End Type

'SPI function for drawing a rectangle
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'windows API call for drawing polylines
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, _
        lpPoint As POINTAPI, ByVal nCount As Long) As Long

'API function for drawing from current point to point specified
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
'API Call for moving to a point without drawing a line
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Public Function DrawBarGraph(p As PictureBox, DATA() As Integer) As Integer
    '**************************************************************88
    'Created:04/14/00        By:Ray Raf.
    'Purpose:   Draw bar graph using Windows API function call in a picturebox
    '****************************************************************
    Dim SoldMode, SoldTop, SoldLeft, SoldWidth, SoldHeight ' used to restore settings when complete
    Dim iWindowWidth As Integer, iWindowHeight As Integer
    Dim iMaxHeight As Integer      'maximum value in data
    Dim iNumPoints As Integer      'number of data points in DATA
    Dim iLoop As Integer
    Dim vReturn As Variant
    Dim scData() As Integer           'data scaled to fit in box (in Pixels)
    Dim sglScaleFactor As Single    'used to scale DATA values to pixel heights
    Dim iBarInt As Integer              'interval between bars
    Dim iBarOffset As Integer           'offset of bar from begining of interval
    Dim iBarWidth As Integer            'width of bar (nominal = 75% of interval)
    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer      'values for rectangle call
    
    'save current mode
    SoldMode = p.ScaleMode
    SoldTop = p.ScaleTop
    SoldLeft = p.ScaleLeft
    SoldWidth = p.ScaleWidth
    SoldHeight = p.ScaleHeight
    
    'make sure to setup for api call
    p.ScaleMode = 3 '3 = pixel mode, required by API call
    p.FillStyle = 0 '= solid , We want to fill rect.
    p.DrawWidth = 1
        
    'resize scaled data array to match passed array
    iNumPoints = UBound(DATA) + 1
    ReDim scData(iNumPoints)
    
    'find maximum value in DATA array
    For iLoop = 0 To iNumPoints - 1
        If DATA(iLoop) > iMaxHeight Then
            iMaxHeight = DATA(iLoop)
        End If
    Next iLoop
    
    'figure scale factor to make max value fit in picbox
    If iMaxHeight > 0 Then
        sglScaleFactor = p.ScaleHeight / iMaxHeight
    Else
        sglScaleFactor = p.ScaleHeight
    End If
    
    'scale data
    For iLoop = 0 To iNumPoints - 1
        scData(iLoop) = Int(DATA(iLoop) * sglScaleFactor)
    Next iLoop
    
    'calc width and offset for bars
    iBarInt = p.ScaleWidth / iNumPoints
    iBarWidth = iBarInt * 0.75
    iBarOffset = iBarInt * 0.125 ' half of the remaining 25%
    
    'draw bars
    For iLoop = 0 To iNumPoints - 1
        X1 = iLoop * iBarInt + iBarOffset
        Y1 = p.ScaleHeight
        X2 = X1 + iBarWidth
        Y2 = Y1 - scData(iLoop)
        vReturn = Rectangle(p.hdc, X1, Y1, X2, Y2)
    Next iLoop
    
    DrawBarGraph = iMaxHeight
End Function

Function DrawLineGraph(p As Control, DATA() As POINTDOUBLE, FColor As Long, nPoints As Integer) As MINMAX
    '**************************************************************88
    'Created:9/25/99        By:Ray Raf.
    'Purpose:   Draw line graph using Windows API function call in a picturebox
    '****************************************************************
    Dim SoldMode, SoldTop, SoldLeft, SoldWidth, SoldHeight ' used to restore settings when complete
    Dim iWindowWidth As Integer, iWindowHeight As Integer
    Dim ptsLine() As POINTAPI, ptReturn As POINTAPI
    Dim dblMaxHeight As Double, dblMinHeight As Double
    Dim dblMaxWidth As Double, dblMinWidth As Double
    Dim dblYDataRange As Double, dblXDataRange As Double
    Dim dblYScaleFactor As Double, dblYOffset As Integer
    Dim dblXScaleFactor As Double, dblXOffset As Integer
    Dim iPointLoop As Integer
    Dim vReturn As Variant
    Dim intYMargin As Integer, intXMargin As Integer
    Dim intXAxisPos As Integer
    
    
    'save current mode
    SoldMode = p.ScaleMode
    SoldTop = p.ScaleTop
    SoldLeft = p.ScaleLeft
    SoldWidth = p.ScaleWidth
    SoldHeight = p.ScaleHeight
    
    'make sure to setup for api call
    p.ScaleMode = 3 '3 = pixel mode, required by API call
    p.FillStyle = 1 'transparent (don't want to fill
    p.ForeColor = FColor
    p.DrawWidth = 1
    
    ReDim ptsLine(UBound(DATA))  'resize to match passed data array
    'note that nPoints does not have to be the same as the array size.
    If nPoints > UBound(DATA) Then
        'checking just in case someone makes an error in the function call
        nPoints = UBound(DATA)
    End If
    
    
    'the following loop checks the data to get the min and max Y
    'values in the data.  This is done because all of the data
    'has to be converted to integers to represent pixels as
    'required by the API call
    For iPointLoop = 1 To nPoints
        With DATA(iPointLoop)
            If .Y > dblMaxHeight Then
               dblMaxHeight = .Y
            End If
                
            If .Y < dblMinHeight Then
                dblMinHeight = .Y
            End If
            
            
            If iPointLoop = 1 Then
                'the first x value should be our left side
                dblMinWidth = .x
                'minheight, maxheight areset to the first value to aviod a false zero
                dblMinHeight = .Y
                dblMaxHeight = .Y
            End If
            'the last x value should be the right side
            If iPointLoop = nPoints Then dblMaxWidth = .x
            
        End With
        
    Next iPointLoop
            
    'Debug.Print p.ScaleHeight, p.ScaleWidth
    dblYDataRange = dblMaxHeight - dblMinHeight
    If dblYDataRange < 0.01 Then   'the data range is nearly zero, set some resonable values
        'pass back min and max heights for display
        DrawLineGraph.Min = dblMinHeight
        DrawLineGraph.Max = dblMaxHeight
        'set dblMinHeight and dblMaxHeight above and below the data
        dblMinHeight = dblMinHeight - 0.01
        dblMaxHeight = dblMaxHeight + 0.01
        dblYDataRange = dblMaxHeight - dblMinHeight
    Else
        'pass back min and max heights for display
        DrawLineGraph.Min = dblMinHeight
        DrawLineGraph.Max = dblMaxHeight
    End If
    dblXDataRange = dblMaxWidth - dblMinWidth
    
    'Want to leave a 10 percent margin (5 top and botom) on y axis
    intYMargin = p.ScaleHeight * 0.1
    dblYScaleFactor = (p.ScaleHeight - intYMargin) / dblYDataRange
    dblYOffset = dblMinHeight * dblYScaleFactor
    
    'the x axis margin is fixed at xxx pixels (for labels on left side
    intXMargin = 0
    dblXScaleFactor = (p.ScaleWidth - intXMargin) / dblXDataRange
    dblXOffset = dblMinWidth * dblXScaleFactor
    'dblXOffset = intXMargin
    
    For iPointLoop = 1 To nPoints
        'still upside down
        With DATA(iPointLoop)
            ptsLine(iPointLoop).x = CInt((dblXScaleFactor * .x) - dblXOffset) + intXMargin
            ptsLine(iPointLoop).Y = p.ScaleHeight - CInt((dblYScaleFactor * .Y) - dblYOffset) - (intYMargin / 2)
            'Debug.Print iPointLoop, ptsLine(iPointLoop).x, ptsLine(iPointLoop).y, .x, .y
            
        End With
    Next iPointLoop


    'draw X axis first if needed
    If dblMinHeight <= 0 Then
        'display X axis only if the graph crosses zero
        'figure where the axis should be drawn
        intXAxisPos = p.ScaleHeight - CInt((dblYScaleFactor * 0) - dblYOffset) - (intYMargin / 2)
        'move to left side at corrent height in picturebox
        vReturn = MoveToEx(p.hdc, 0, intXAxisPos, ptReturn)
        'set color for axis to be drawn in
        p.ForeColor = vbBlack
        'draw line to right side
        vReturn = LineTo(p.hdc, p.ScaleWidth, intXAxisPos)
    End If
    
    'draw line representing data
    p.ForeColor = FColor
    vReturn = Polyline(p.hdc, ptsLine(1), CLng(nPoints))
    
    
    
    p.ScaleMode = SoldMode
    p.ScaleTop = SoldTop
    p.ScaleLeft = SoldLeft
    p.ScaleHeight = SoldHeight
    p.ScaleWidth = SoldWidth
End Function


