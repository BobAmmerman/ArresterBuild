Option Strict On
Public Class UnbuildModule

    Dim modworkOrder As New WorkOrder
    Private Sub CancelUnbuildButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelUnbuildButton.Click
        Me.Close()
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        If (ModulesListBox.Items.Count = 0) Then
            FillModulesBox()
			 If (ModulesListBox.Items.Count = 0) Then
				MsgBox("No modules found for this serial number.")
				End If
		Else

			Dim serNum As Long
			Dim success As Boolean = True
			' If (Not UnitTextBox.Text = "") Then
			' unitNum = Asc(UnitTextBox.Text.ToUpper().Chars(0)) - 64
			' End If
			Dim modToDelete As ArresterModule = GetModuleFromString(ModulesListBox.Text)

			serNum = CLng(SerialNumberTextBox.Text)
			If (Not modToDelete Is Nothing) Then
				If (modworkOrder.UnbuildUnit(serNum, modToDelete.OldModNumber, Nothing)) Then
					MsgBox("Successfully unbuilt SN: " & serNum & " Unit: " & modToDelete.Unit.UnitLetter & " Module: " & modToDelete.ModuleNumber)
				End If
				Me.Close()
			End If

		End If
    End Sub

    Private Sub FillModulesBox()
		Dim serNum As Long
		ModulesListBox.Items.Clear()
        If (Not Long.TryParse(SerialNumberTextBox.Text, serNum)) Then
            MsgBox("Serial number must be a number!")
        End If
        If (modworkOrder.Lookup(serNum, Nothing)) Then
            Dim Units As List(Of ArresterUnit) = modworkOrder.CatNum.Units
            Dim unit As ArresterUnit
            For Each unit In Units
                Dim arresterMod As ArresterModule
                For Each arresterMod In unit.Modules
                    If (arresterMod.IsBuilt()) Then
                        ModulesListBox.Items.Add(arresterMod.moduleString())
                    End If
                Next
            Next unit

            If (ModulesListBox.Items.Count > 0) Then
                ModulesListBox.SelectedItem = ModulesListBox.Items(0)
            ElseIf (Units.Count <> 0) Then
                MsgBox("There are no modules left to unbuild for this serial number.")
            End If

        End If
    End Sub

    Private Function GetModuleFromString(ByVal modString As String) As ArresterModule
        Dim retVal As ArresterModule = Nothing
        Dim unit As ArresterUnit
        Dim Units As List(Of ArresterUnit) = modworkOrder.CatNum.Units
        Dim arresterMod As ArresterModule
        For Each unit In Units
            For Each arresterMod In unit.Modules
                If modString = arresterMod.moduleString() Then
                    retVal = arresterMod
                End If
            Next
        Next

        Return retVal
    End Function

    Private Sub SerialNumberTextBox_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SerialNumberTextBox.TextChanged
        ' 2019-01-22 RVA - Allow 12 digit serial numbers
        If (SerialNumberTextBox.Text.Length = 12) Then
            FillModulesBox()
        End If
    End Sub
End Class