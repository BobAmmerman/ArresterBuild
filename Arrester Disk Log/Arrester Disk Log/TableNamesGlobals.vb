Option Strict On
Imports System.Data.SqlClient
Module TableNamesGlobals
    Friend ACTUAL_MCOV As String = "Actual_MCOV"
    Friend NAMEPLATE_POLY As String = "NameplatePoly"
    Friend NAMEPLATE_PORC As String = "NameplatePorc"
    Friend NAMEPLATE_CUBE_POLY As String = "NameplateCubePoly"
    Friend NAMEPLATE_DR_POLY As String = "NameplateDrPoly"
    Friend NAMEPLATE_DR_PORC As String = "NameplateDrPorc"   'added this line for Dual rated Porc 1/21/26 RVA
    Friend NAMEPLATE_CUBE_PORC As String = "NameplateCubePorc"
    Friend MCOV_DR_PORC As String = "MCOV_Dr_Porc"
    Friend MCOV_DR_POLY As String = "MCOV_Dr_Poly"
    Friend MCOV_CUBE As String = "MCOV_Cube"
    Friend MCOV_PORC_CALCINED As String = "MCOV_Porc_Rev1"
    Friend DISK_QTY_PORC_CALCINED As String = "DiskQuanPorc_Rev1"
    Friend DISK_QTY_POLY As String = "DiskQuanPoly_Rev2"
    Friend DISK_QTY_ULTRASIL As String = "Poly_DiskQuantity_Rev0"
    ' 2017-07-07 RVA - Replaced following line as per Mike R
    'Friend BUILD_TABLE_DR_PORC As String = "BuildDrPorc"
    Friend BUILD_TABLE_DR_PORC As String = "BuildTablePorc_Rev1"
    Friend BUILD_TABLE_CUBE As String = "BuildTableCube_Rev1"
    Friend BUILD_TABLE_PORC_CALCINED As String = "BuildTablePorc_Rev1"
    Friend HOUSING_TABLE_DR_PORC As String = "HousingDrPorc"
    Friend HOUSING_TABLE_PORC As String = "HousingPorc"
    Friend MCOV_12MM_NEW As String = "MCOV_12mm_new"
    Friend MCOV_12MM As String = "MCOV_12mm"
    Friend MCOV_1IN As String = "MCOV_1in"
    Friend COMPONENTS_12MM_NEW As String = "components_12mm_new_Rev1"
    Friend COMPONENTS_12MM_OLD As String = "components_12mm_old"
    Friend COMPONENTS_1IN As String = "components_1in_Rev1"
    Friend HOUSING_DR_POLY As String = "HousingDrPoly"
    Friend HOUSING_POLY As String = "HousingPoly"
    Friend MODULE_COMPONENTS_POLY As String = "ModuleComponents_Poly"
    Friend MCOV_HOUSING_ULTRASIL As String = "Poly_MCOV_Housing_Rev0"
    Friend BUILD_NOTES_POLY As String = "BuildNotesPoly"
    Friend DISK_PARTNERS As String = "DiskPartners_Rev3"
    Friend MCOV_HOUSING_PARTNUM As String = "Poly_Housing_PartNumber_Rev0"
    Friend GRoomConnection As New SqlConnection(My.Settings.GRoomConnectionString)
    Friend ngd_ProductionConnection As New SqlConnection(My.Settings.ngd_productionConnectionString)
    Friend PolymerStationConnection As New SqlConnection(My.Settings.PolymerStationConnectionString)
    Friend BUILT_UNITS As String = "BuiltUnits"
    Friend ULTRASIL_BASIC_ARRESTER_DATA As String = "Poly_BasicArresterData_Rev0"


    Function dbMatch(ByVal query As String, ByRef trans As SqlTransaction) As Boolean
        Return dbMatch(query, GRoomConnection, trans)
    End Function

        Function dbMatch(ByVal query As String, ByRef Connection As SqlConnection ,ByRef trans As SqlTransaction) As Boolean
        Dim retVal As Boolean = False
        Dim command As New SqlCommand(query, Connection)
        If (Not trans Is Nothing) Then
            command.Transaction = trans
        End If
        Dim reader As SqlDataReader = command.ExecuteReader
        Try
            retVal = reader.HasRows()
        Finally
			reader.Close()
        End Try
        Return retVal
    End Function
End Module
