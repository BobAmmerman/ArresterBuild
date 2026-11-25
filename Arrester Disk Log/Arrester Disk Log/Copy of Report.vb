Option Strict On
Public Class Report


    Private m_workOrder As WorkOrder
    Private m_catNum As CatNum
    Private m_report As String = Nothing
    Private m_qtyMap As List(Of CatQtyMapping)

    '   Public Sub New(ByRef moduleIn As ArresterModule, ByRef qtyMapIn As List(Of CatQtyMapping))
    '        m_workOrder = moduleIn.WorkOrder
    '       m_catNum = m_workOrder.CatNum
    '       m_qtyMap = qtyMapIn
    '   End Sub

    Public Sub New(ByRef priWO As WorkOrder)
        m_workOrder = priWO
        'm_workOrder.Lookup(m_workOrder.CatNum.SerialNumber)
        m_catNum = m_workOrder.CatNum
    End Sub

    Public Sub New(ByVal serNumIn As Int64)
        m_workOrder = New WorkOrder()
        m_workOrder.Lookup(serNumIn, Nothing)
        m_catNum = m_workOrder.CatNum
        m_workOrder.CatNum.SerialNumber = serNumIn
    End Sub

    Function GetReport() As String
        addHeader()
        addBasics()

        If (m_catNum.Type = CatNum.Types.PORC) Then
            addArresterNameplate()
            addUnitNameplates()
            addAssyDetails()
            addBuildNotes()
        Else
            addAssyDetails()
            addArresterNameplate()
            addStackedNameplate()
            addCoverPlateEtching()
            addBuildNotes()

        End If
        'endReport()
        Return m_report
    End Function

    Private Sub addHeader()


        m_report = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" & vbCr & _
"{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\sl240\slmult1\qc\f0\fs24  \lang9\b\fs32 Report For " & m_workOrder.WorkOrderNum & ", Serial: " & m_catNum.SerialNumber & "\par" & vbCr & _
"\b0\fs22\par" & vbCr

        'Changed to single space -8May2012 RPost
        '        m_report = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Calibri;}}" & vbCr & _
        '"{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\sa200\sl240\slmult1\qc\lang9\b\f0\fs32 Report For " & m_workOrder.WorkOrderNum & ", Serial: " & m_catNum.SerialNumber & "\par" & vbCr & _
        '"\b0\fs22\par" & vbCr
    End Sub

    Private Sub addBasics()
        If m_workOrder.CatNum.Type = CatNum.Types.POLY Then
            m_report = m_report & "\pard\sa200\sl240\slmult1Arrester Mounting:\tab " & m_catNum.Mounting & "\par" & vbCr
        End If

        m_report = m_report & "\pard\sa200\sl240\slmult1Catalog Number:\tab " & m_catNum.CatNumString & "\par" & vbCr & _
"WO Number:\tab\tab " & m_workOrder.WorkOrderNum & "\par" & vbCr & _
"Serial Number:\tab\tab " & m_catNum.SerialNumber & " \par" & vbCr
        '"Unit:\tab\tab\tab " & m_module.UnitLetter & "\par" & vbCr & _
        '"Module:\tab\tab\tab " & m_module.ModuleNumber & "\par" & vbCr
        '& _"\par"

    End Sub

    Private Sub addDiskQtys()
        '        m_report = m_report & "\pard\sa200\sl240\slmult1\b Disk Selection - Unit " & m_module.UnitLetter & ":\b0\par" & vbCr & _
        '       "\pard\li720\sa200\sl240\slmult1\b Disk Type\b0\tab\b Qty\b0\tab\tab\b Category\b0\par" & vbCr
        Dim row As Object
        Dim tmpCatMap As CatQtyMapping
        For Each row In m_qtyMap
            tmpCatMap = DirectCast(row, CatQtyMapping)
            m_report = m_report & tmpCatMap.diskName & "\tab\tab " & tmpCatMap.assertedQty & "\tab\tab " & tmpCatMap.category & "\par" & vbCr
        Next row
    End Sub

    Private Sub addStackedNameplate()
        Dim ringShed As Integer = 0
        If (m_workOrder.CatNum.Type = CatNum.Types.POLY AndAlso (Not Integer.TryParse(m_catNum.RingShedCode, ringShed) Or ringShed > 62)) Then
            m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
    "\pard\sa200\sl240\slmult1\b Stacked Nameplate\par" & vbCr & _
    "\pard\li720\sa200\sl240\slmult1\b0 Catalog Number:\tab " & m_catNum.CatNumString & "\tab\tab\par" & vbCr & _
    "Serial Number:\tab " & m_catNum.SerialNumber & "\par" & vbCr & _
    "\par" & vbCr


            m_catNum.Units.Reverse()
            Dim unit As ArresterUnit
            For Each unit In m_catNum.Units
                m_report = m_report & "Unit:\tab " & unit.UnitLetter & "\tab " & unit.UnitMCOV & "kV rms\par" & vbCr
            Next unit
            m_catNum.Units.Reverse()
        End If
    End Sub

    Private Sub addCoverPlateEtching()

        If (m_catNum.NumModules >= 3 And m_catNum.Stud = "1in") Then
            m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
    "\pard\sa200\sl240\slmult1\b Cover Plate Etching:\par" & vbCr & _
    "\pard\li720\sa200\sl240\slmult1\b0 Serial Number:\tab\tab " & m_catNum.SerialNumber & "\tab\tab\par" & vbCr & _
    "\par"
        End If
    End Sub

    Private Sub addBuildNotes()
        Dim unit As ArresterUnit
        For Each unit In m_catNum.Units
            If (unit.Modules(0).BuildNotes <> "NA") Then
                m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
        "\pard\sa200\sl240\slmult1\b Build Notes:\par" & vbCr & _
        "Unit:\tab " & unit.UnitLetter & "\par" & vbCr & _
        "\pard\li720\sa200\sl240\slmult1\b0 " & unit.Modules(0).BuildNotes.Replace(vbLf, "\par" & vbCr) & "\par" & vbCr & _
        "\par"
            End If
        Next unit
    End Sub

    Private Sub addAssyDetails()
        Dim unit As ArresterUnit
        For Each unit In m_catNum.Units
            If (m_workOrder.CatNum.Type = CatNum.Types.PORC) Then
                m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
        "\pard\sa200\sl240\slmult1\b Arrester Unit " & unit.UnitLetter & " Assembly Details:\par" & vbCr & _
        "\pard\li720\sa200\sl240\slmult1\b0 Porcelain Housing:\tab " & unit.Modules(0).Housing & "\tab\tab\par" & vbCr & _
        "MCOV/COV(kV):\tab " & unit.UnitMCOV & "\par" & vbCr & _
        "\par" & vbCr
            ElseIf (m_workOrder.CatNum.Type = CatNum.Types.POLY) Then
                m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
        "\pard\sa200\sl240\slmult1\b Arrester Unit " & unit.UnitLetter & " Assembly Details:\par" & vbCr & _
        "\pard\li720\sa200\sl240\slmult1\b0 Shed Sections:\tab\tab " & unit.UnitShed & " \tab\tab\par" & vbCr & _
        "MCOV/COV(kV):\tab " & unit.UnitMCOV & "\par" & vbCr & _
        "\par" & vbCr
            End If
        Next unit
    End Sub

    Private Sub addArresterNameplate()
        m_report = m_report & "\pard\sa200\sl240\slmult1\b Arrester Nameplate:\b0\par" & vbCr & _
          "\pard\li720\sa200\sl240\slmult1 Catalog Number:\tab " & m_catNum.CatNumString & "\par" & vbCr & _
          "Serial Number:\tab\tab " & m_catNum.SerialNumber & "\par" & vbCr & _
          "Year Of Manuf:\tab\tab " & m_catNum.Units(0).Modules(0).BuildDate.Year & "\par" & vbCr & _
          "Rating:\tab\tab\tab " & m_catNum.Nameplate.Rating & "\par" & vbCr & _
          "MCOV/COV\tab\tab " & m_catNum.Nameplate.MCOV & "\par" & vbCr & _
          "Pressure Relief:\tab " & m_catNum.Nameplate.PressureRelief & " \par" & vbCr & _
          "Class:\tab\tab\tab " & m_catNum.Nameplate.ArresterClass & "\par" & vbCr & _
          "Certification:\tab\tab " & m_catNum.Nameplate.Certification & "\par" & vbCr
    End Sub

    Private Sub addUnitNameplates()
        Dim unit As ArresterUnit
        For Each unit In m_catNum.Units
            If (m_workOrder.CatNum.Type = CatNum.Types.PORC) Then
                m_report = m_report & "\pard\li1440\sa200\sl240\slmult1\par" & vbCr & _
        "\pard\sa200\sl240\slmult1\b Unit " & unit.UnitLetter & "  Nameplate:\par" & vbCr & _
        "\pard\li720\sa200\sl240\slmult1\b0 Catalog Number:\tab " & m_catNum.CatNumString & "\tab\tab\par" & vbCr & _
        "Serial Number:\tab\tab " & m_catNum.SerialNumber & "\par" & vbCr & _
        "MCOV/COV:\tab\tab " & unit.UnitMCOV & "\par" & vbCr & _
        "Unit:\tab\tab\tab " & unit.UnitNumber & " of " & m_catNum.NumUnits & "\par" & vbCr & _
        "Position:\tab\tab " & PorcUnitPosition(unit.UnitNumber, m_catNum.NumUnits) & "\par" & vbCr & _
        "\par" & vbCr
            End If

        Next unit
    End Sub

    Public Function PorcUnitPosition(ByVal iPos As Integer, ByVal iTotal As Integer) As String
        '--------------------------------------------------------------------------------------------------------------
        '   This function returns a descriptive string of the position of the unit (iPos)
        '   relative to the total (iTotal) number of units in an arrester.
        '--------------------------------------------------------------------------------------------------------------

        PorcUnitPosition = "unknown"

        ' there are only 4 units max in porc arresters
        Select Case iPos
            Case 1 : PorcUnitPosition = "Bot."
            Case 2
                Select Case iTotal
                    Case 2 : PorcUnitPosition = "Top"
                    Case 3, 4 : PorcUnitPosition = "2nd"
                End Select
            Case 3
                Select Case iTotal
                    Case 3 : PorcUnitPosition = "Top"
                    Case 4 : PorcUnitPosition = "3rd"
                End Select
            Case 4 : PorcUnitPosition = "Top"
        End Select
    End Function

    Private Sub endReport()
        m_report = m_report & "}"
    End Sub


End Class
