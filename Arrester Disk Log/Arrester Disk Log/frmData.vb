Imports System.Data.SqlClient

Public Class frmData

    Public Sub FillForm()

        ResetForm()

        If isValidWO(tbWO.Text) Then
            If True Then

                Dim sSQL As String
                Dim cat As New CatNum
                Dim totDBUnits As Integer
                Dim UnitsBuilt() As Integer = {}

                'Get WO Info
                sSQL = "SELECT CatalogNumber, Unit1NumBuilt AS U1, Unit2NumBuilt AS U2, Unit3NumBuilt AS U3, Unit4NumBuilt AS U4, Unit5NumBuilt AS U5, Unit6NumBuilt AS U6, NumArrestersRequired AS UR " & _
                        "FROM WO_Info " & _
                        "WHERE WO_Number='" & tbWO.Text & "' "
                Using dr As SqlDataReader = New SqlCommand(sSQL, GRoomConnection).ExecuteReader()
                    If dr.HasRows() Then
                        dr.Read()
                        UnitsBuilt = {dr!UR, dr!U1, dr!U2, dr!U3, dr!U4, dr!U5, dr!U6}
                        cat.CatNumString = dr!CatalogNumber
                        lblCatNum.Text = cat.CatNumString
                    End If

                End Using

                If lblCatNum.Text <> "" Then
                    totDBUnits = 2 * (cat.Units.Count - 1) + cat.Units(0).m_modules.Count
                    sSQL = ""
                    For i As Integer = 1 To totDBUnits
                        '    sSQL = sSQL & UnitsBuilt(i) & IIf(i < totDBUnits, ",", "")
                        sSQL = String.Format("{0}{1}: {2} of {3}{4}", sSQL, dbUnit2ModUnit(i, totDBUnits), UnitsBuilt(i), UnitsBuilt(0), vbCrLf)
                    Next
                    'sSQL = sSQL & " of " & UnitsBuilt(0) & " complete"

                    lblCompleted.Text = sSQL

                    tvData.Top = Math.Max(lblCompleted.Bottom, lblCatNum.Bottom) + 5


                    'Populate TreeView
                    sSQL = "SELECT     WO_SerialNumbers.SerialNumber SN, " & _
                                    "isNull(ConsumedDisks.UnitNumber,0) UnitNumber, isNull(ConsumedDisks.Category,0) Category, isnull(ConsumedDisks.Disk_Number,0) Disk_Number, isnull(ConsumedDisks.Batch_Number,0) Batch_Number " & _
                            "FROM       WO_SerialNumbers LEFT OUTER JOIN ConsumedDisks ON WO_SerialNumbers.SerialNumber = ConsumedDisks.SerialNumber " & _
                            "WHERE      WO_SerialNumbers.WO_Number='" & tbWO.Text & "' " & _
                            "ORDER BY   SN,UnitNumber,Category,Batch_Number,Disk_Number"
                    Dim SN As String = ""
                    Dim dbUnit As Integer = 0
                    Dim DiskNum As Integer = 0
                    Dim currNode As TreeNode = Nothing
                    Dim tmpNode As TreeNode



                    Using dr As SqlDataReader = New SqlCommand(sSQL, GRoomConnection).ExecuteReader()
                        'Getting static data from 1st row
                        If dr.HasRows() Then
                            While dr.Read()

                                If (dr!SN = SN) Then
                                Else
                                    SN = dr!SN
                                    dbUnit = 0
                                    DiskNum = 0
                                    currNode = New TreeNode(SN)
                                    tvData.Nodes.Add(currNode)
                                End If

                                If (dr!UnitNumber = dbUnit) Then
                                Else
                                    dbUnit = dr!UnitNumber
                                    tmpNode = New TreeNode(dbUnit2ModUnit(dbUnit, totDBUnits))
                                    If DiskNum = 0 Then
                                        'New SN (currNode=parent) just created
                                        currNode.Nodes.Add(tmpNode)
                                    Else
                                        'Just New unit (currNode=sibling)
                                        'currNode.Parent = here.parent
                                        currNode.Parent.Nodes.Add(tmpNode)
                                    End If
                                    currNode = tmpNode
                                    'creating a label
                                    currNode.Nodes.Add("Cat, Batch, Disk")
                                End If


                                'If (dr!Disk_Number = DiskNum) Then  'No need, always will be false
                                'Else

                                'currNode will always be a parent
                                DiskNum = dr!Disk_Number
                                currNode.Nodes.Add(String.Format("{0}, {1}, {2}", dr!Category, dr!Batch_Number, DiskNum))

                                'End If


                            End While
                        End If
                    End Using

                    'Color SNs that are incomplete red
                    For Each currNode In tvData.Nodes
                        If currNode.Nodes.Count < totDBUnits Then
                            currNode.ForeColor = Color.Red
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Private Function dbUnit2ModUnit(ByVal dbUnit As Integer, ByVal totalUnits As Integer) As String
        Const letters As String = "AABBCC"
        Const numbers As String = "1121212"
        Dim letter As Char
        Dim number As Char

        letter = letters.Chars(totalUnits Mod 2 + dbUnit - 1)
        number = numbers.Chars((totalUnits + 1) Mod 2 + dbUnit - 1)
        Return letter & number
    End Function


    Private Function isValidWO(ByVal WONum As String)
        'Is WO an 8digit number?
        If (WONum.Length = 8) Then
            If (IsNumeric(WONum)) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub ResetForm()
        tvData.Nodes.Clear()
        lblCatNum.Text = ""
        lblCompleted.Text = "No WO Loaded"
    End Sub

    Private Sub tbWO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbWO.TextChanged
        If isValidWO(tbWO.Text) Then FillForm()
    End Sub
End Class