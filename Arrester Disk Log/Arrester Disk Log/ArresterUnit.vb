Option Strict On
Public Class ArresterUnit

    Private m_unitNumber As Integer

    Public Sub New(ByVal unitNumIn As Integer)
        m_unitNumber = unitNumIn
    End Sub

    ReadOnly Property UnitNumber As Integer
        Get
            Return m_unitNumber
        End Get
    End Property

    ReadOnly Property UnitLetter() As String
        Get
            Return Chr(m_unitNumber + 64)
        End Get
    End Property

    Public m_modules As New List(Of ArresterModule)

    ReadOnly Property UnitMCOV As Double
        Get
            Dim aMod As ArresterModule
            Dim sumMCOV As Double = 0.0
            For Each aMod In m_modules
                sumMCOV += CDbl(aMod.ModuleMCOV)
            Next
            Return sumMCOV
        End Get
    End Property

    ReadOnly Property UnitShed As Integer
        Get
            Dim aMod As ArresterModule
            Dim shedSum As Integer = 0
            For Each aMod In m_modules
                shedSum += aMod.Shed
            Next
            Return shedSum
        End Get
    End Property

    ReadOnly Property Modules As List(Of ArresterModule)
        Get
            Return m_modules
        End Get
    End Property


    Public Sub AddModule(ByVal ModuleNumIn As Integer, ByVal catNumIn As CatNum, ByVal MCOVIn As Double)
        m_modules.Add(New ArresterModule(Me, catNumIn, ModuleNumIn, CStr(MCOVIn)))
    End Sub



End Class
