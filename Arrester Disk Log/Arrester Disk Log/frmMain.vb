Option Strict On
Imports System.Data.SqlClient

Public Structure CatQtyMapping : Implements IComparable
	Public category As Integer
	Public assertedQty As Integer
	Public actualQty As Integer
	Public diskName As String
	Public Sub New(ByVal cat As Integer, ByVal asserted As Integer, ByVal actual As Integer, ByVal dsk As String)
		category = cat
		assertedQty = asserted
		actualQty = actual
		diskName = dsk
	End Sub

	Public Overrides Function Equals(ByVal obj As Object) As Boolean
		Dim compMap As CatQtyMapping = DirectCast(obj, CatQtyMapping)

		If ( _
			category = compMap.category _
			) Then
			Return True
		Else
			Return False
		End If


	End Function

	Public Function Compare(ByVal o1 As Object) As Integer Implements IComparable.CompareTo
		Dim map1 As CatQtyMapping
		map1 = DirectCast(o1, CatQtyMapping)
		If (map1.category < Me.category) Then
			Return 1
		ElseIf (map1.category > Me.category) Then
			Return -1
		Else
			Return 0
		End If
	End Function
End Structure

Public Class frmMain

#Region "Internal Data Structures"
	Private loginForm As frmLogin
	Private loginObject As Login
	Private ManualWorkOrderMode As Boolean = False
	Private LoginAlive As Boolean = True

	Private diskCatsMode As Boolean = True
	Private primaryWorkOrder As New WorkOrder
	Private activeModule As ArresterModule
	Private justFailedDiskCheck As Boolean = False
	Private qtysList As New List(Of CatQtyMapping)

	Enum States
		woSelect = 0
		unitSel = 1
		numDisks = 2
		diskInfo = 3
		unitReport = 4
	End Enum

	Dim currState As Integer
	Public Sub New(ByRef loginObjIn As Login, ByRef frmLoginIn As frmLogin)
		Me.InitializeComponent()
		loginObject = loginObjIn
		currState = States.woSelect
		woSelectLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
		prevButton.Enabled = False
		SerNumBottom.Text = ""
		WorkOrderBottom.Text = ""
		UnitBottom.Text = ""
		ModuleBottom.Text = ""
		DateBottom.Text = DateTime.Now.ToString()
		loginForm = frmLoginIn
		WorkOrderContainer.Show()
		SelectWOPreviousWorkOrder.Text = My.Settings.LastWorkOrder
		PagesTabControl.DrawMode = TabDrawMode.OwnerDrawFixed
		PagesTabControl.ItemSize = New System.Drawing.Size(0, 1)
		PagesTabControl.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
		EntryCancelButton.Visible = False
		Dim objArray(3) As Object
		objArray(0) = 0
		objArray(1) = ""
		objArray(2) = ""

		Try
			If (GRoomConnection.State <> ConnectionState.Open) Then
				GRoomConnection.Open()
			End If
			If (PolymerStationConnection.State <> ConnectionState.Open) Then
				PolymerStationConnection.Open()
			End If

			If (ngd_ProductionConnection.State <> ConnectionState.Open) Then
				ngd_ProductionConnection.Open()
			End If
		Catch e As Exception
			Dim setForm As New Settings(frmLoginIn)
			loginForm.Hide()
			setForm.Show()
			setForm.BringToFront()
			MsgBox("There was an error connecting to the database. Please check your settings and try again.")
		End Try




		DiskDetailPanel.Hide()
		SelectWOWorkOrderTextBox.Select()
		primaryWorkOrder.CatNum.WorkOrder = primaryWorkOrder
		If (My.Settings.LoginTimeoutMs > 1000) Then
			LoginTimer.Start()
		End If
		ClockTickTimer.Start()
	End Sub

	Private Sub DiskQtyCategory_Init()
		DiskQtyCategoryGrid.Rows.Clear()
		DiskQtyCategoryGrid.Rows.Insert(0, 1, "", "")
		DiskQtyCategoryGrid.Rows.Insert(1, 2, "", "")
		DiskQtyCategoryGrid.Rows.Insert(2, 3, "", "")
		DiskQtyCategoryGrid.Rows.Insert(3, 4, "", "")
		DiskQtyCategoryGrid.Rows.Insert(4, 5, "", "")
		DiskQtyCategoryGrid.Rows.Insert(5, 6, "", "")
		DiskQtyCategoryGrid.CurrentCell = DiskQtyCategoryGrid.Rows(0).Cells(1)
	End Sub

#End Region

#Region "Event Handlers"
	Private Sub nextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nextButton.Click
		LoginAlive = True
		If (currState = States.woSelect AndAlso ManualWorkOrderMode = True) Then
			If (CommitManualWO()) Then

				ManualWOContainer.Hide()
				SelectWOPageValidation()

				ChangeState(States.unitSel)
			End If
		ElseIf (currState = States.woSelect AndAlso ManualWorkOrderMode = False AndAlso SelectWOPageValidation()) Then
			If (primaryWorkOrder.numRequired * primaryWorkOrder.CatNum.NumModules <= primaryWorkOrder.GetNumBuilt) Then
				MsgBox("All of the modules for this work order have already been built.")
				ChangeState(States.woSelect)
			Else
				ChangeState(States.unitSel)
			End If
		ElseIf currState = States.unitSel Then
			ChangeState(States.numDisks)
		ElseIf (currState = States.numDisks AndAlso NumDisksPageValidation()) Then
			ChangeState(States.diskInfo)
		ElseIf (currState = States.diskInfo) Then
			Dim ddValidSuccess As Boolean = DiskDetailValidation()
			If (ddValidSuccess) Then
				If (activeModule.ModuleNumber = 1 AndAlso activeModule.Unit.UnitNumber = 1) Then
					ChangeState(States.unitReport)
					GenerateReport()
				ElseIf (currState = States.diskInfo AndAlso activeModule.NumberBuilt < primaryWorkOrder.numRequired()) Then
					DiskInfoQtysDoneButton_Click(sender, e)
					FillSNBox()
					ChangeState(States.diskInfo)
				ElseIf (primaryWorkOrder.numRequired * primaryWorkOrder.CatNum.NumModules = primaryWorkOrder.GetNumBuilt) Then
					DiskInfoQtysDoneButton_Click(sender, e)
					MsgBox("Successfully built the module.")
					ChangeState(States.woSelect)
				Else
					DiskInfoQtysDoneButton_Click(sender, e)
					MsgBox("Successfully built the module.")
					ChangeState(States.unitSel)
				End If
			End If

		ElseIf (currState = States.unitReport AndAlso Not activeModule Is Nothing AndAlso primaryWorkOrder.numRequired * primaryWorkOrder.CatNum.NumModules = primaryWorkOrder.GetNumBuilt) Then
		'Module complete
			DiskInfoQtysDoneButton_Click(sender, e)
			ChangeState(States.woSelect)
		ElseIf (currState = States.unitReport AndAlso Not activeModule Is Nothing AndAlso primaryWorkOrder.GetSerialNumbersUnBuilt(activeModule.UnitNumber).Length = 0) Then
			activeModule.ClearDisks()
			DiskInfoQtysDoneButton_Click(sender, e)
			ChangeState(States.unitSel)
		ElseIf (currState = States.unitReport) Then
			If (Not activeModule Is Nothing AndAlso primaryWorkOrder.GetSerialNumbersUnBuilt(activeModule.UnitNumber).Length > 0) Then
				DiskInfoQtysDoneButton_Click(sender, e)
				ChangeState(States.diskInfo)
			Else
				ChangeState(States.woSelect)
			End If
		End If

	End Sub

	Private Sub prevButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prevButton.Click
		LoginAlive = True
		If (currState = States.unitSel) Then
			primaryWorkOrder = New WorkOrder()
			SelectWOPreviousWorkOrder.Text = My.Settings.LastWorkOrder
			ChangeState(States.woSelect)
			primaryWorkOrder.CatNum.ClearModules()
		ElseIf (currState = States.numDisks) Then
			ChangeState(States.unitSel)
			UnitBottom.Text = ""
		ElseIf (currState = States.diskInfo) Then
			ChangeState(States.numDisks)
			activeModule.ClearDisks()
			If (Not diskCatsMode) Then
				DiskInfoQtysDoneButton_Click(sender, e)
			End If
			DiskQtyCategory_Init()
			SerNumBottom.Text = ""
		ElseIf (currState = States.unitReport) Then
			ChangeState(States.diskInfo)
		End If
	End Sub

	Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
		loginObject.Logout()
		loginForm.Show()
		Me.Enabled = False
	End Sub

	Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
		loginObject.Logout()
		Me.Hide()
		loginForm.Show()
	End Sub

	Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
		frmLogin.Close()
	End Sub

	Private Sub DiskQtyCategoryGrid_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiskQtyCategoryGrid.CurrentCellChanged
		Dim grid As DataGridView = DirectCast(sender, DataGridView)
		If (Not grid Is Nothing And Not grid.CurrentCell Is Nothing) Then
			If grid.Focused = True And grid.CurrentCell.ColumnIndex = 0 Then
				SendKeys.Send("{TAB}")
			End If
		End If
	End Sub

	Private Sub ManualWorkOrderContainer_onVisible(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManualWOContainer.VisibleChanged
		If (ManualWOContainer.Visible = True) Then
			ManualWorkOrderMode = True
		Else
			ManualWorkOrderMode = False
		End If
	End Sub

	Private Sub UsePreviousWorkOrderButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UsePreviousWorkOrderButton.Click
		SelectWOWorkOrderTextBox.Text = SelectWOPreviousWorkOrder.Text
	End Sub

	Private Sub SelectUnitSelectUnitBox_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectUnitSelectUnitBox.SelectedIndexChanged
		If Not CStr(SelectUnitSelectUnitBox.SelectedItem) Is Nothing Then
			SelectUnitNumBuiltTextBox.Text = CStr(primaryWorkOrder.GetUnitsBuilt(GetModuleFromString(CStr(SelectUnitSelectUnitBox.SelectedItem)).UnitNumber))

		End If
	End Sub

	Private Sub DiskInfoQtysDoneButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiskInfoQtysDoneButton.Click

		If (diskCatsMode) Then
			DiskQtyCategoryGrid.SelectionMode = CType(SelectionMode.None, DataGridViewSelectionMode)
			Dim catSum As Integer = 0
			Dim numDisks As Integer = 0
			Dim i As Integer
			Dim inputValid As Boolean = True

			If (currState = States.diskInfo And Not DiskInfoSelectSNBox.SelectedItem Is Nothing) Then
				Dim serNumTmp As Long
				Long.TryParse(DiskInfoSelectSNBox.SelectedItem.ToString, serNumTmp)
				primaryWorkOrder.CatNum.SerialNumber = serNumTmp
			Else
				MsgBox("You must select a serial number to build!")
				inputValid = False
			End If

			For i = 0 To 5
				If (inputValid And CStr(DiskQtyCategoryGrid.Item("Quantity", i).Value) <> "" And _
					CStr(DiskQtyCategoryGrid.Item("Category", i).Value) <> "") Then
					Dim currDiskQty As Integer
					Dim currDiskCat As Integer
					If (Integer.TryParse(CStr(DiskQtyCategoryGrid.Item("Quantity", i).Value), currDiskQty) And _
						Integer.TryParse(CStr(DiskQtyCategoryGrid.Item("Category", i).Value), currDiskCat)) Then
						' removed this if statement so operators can put in correct category    16Apr2012 DLimburg
						'   If (activeModule.Num5kVDisks > 0 And (activeModule.CatNumber.DiskType.Substring(0, 3) = "AT5" _
						'           Or activeModule.CatNumber.PartnerDiskType.Substring(0, 3) = "AT5") And currDiskCat < 1000) Then ' Or activeModule.CatNumber.PartnerDiskType.Substring(0, 4) = "AT52") Then
						'   currDiskCat = currDiskCat + 1000
						' ElseIf (currDiskCat = 100 Or currDiskCat = 110 Or currDiskCat = 120) Then
						'   currDiskCat = 1000
						'End If
						catSum = catSum + (currDiskQty * currDiskCat)
						numDisks = numDisks + currDiskQty
					Else
						MsgBox("All disk categories and quantities must be numbers!", MsgBoxStyle.Critical)
						inputValid = False
					End If
				End If
			Next


			If (inputValid) Then
				If (numDisks <> activeModule.TotalDisks) Then
					MsgBox("Total Disk Quantity must be equal to " & activeModule.TotalDisks & "!", MsgBoxStyle.Critical)
					inputValid = False
				End If
			End If

			If (inputValid) Then
				If (catSum > activeModule.Cum_V10CATMax Or catSum < activeModule.Cum_V10CATMin) Then
					MsgBox("Sum of category values (currently " & catSum & ") must be between " & activeModule.Cum_V10CATMin & " And " & activeModule.Cum_V10CATMax & "!", MsgBoxStyle.Critical)
					inputValid = False
				End If
				DiskInfoDiskDetailDataGrid.Rows.Clear()
				DiskInfoDiskDetailDataGrid.Rows.Add()
				DiskInfoDiskDetailDataGrid.Rows.Item(0).ReadOnly = False
				DiskInfoDiskDetailDataGrid.Rows.Item(0).Resizable = DataGridViewTriState.False

				BatchGridColumn.ReadOnly = False
				DiskNumberGridColumn.ReadOnly = False
				If numDisks > 1 Then						   '24Apr2012 added test to make sure more than one row was needed MMR
					DiskInfoDiskDetailDataGrid.Rows.AddCopies(0, numDisks - 1)
				End If
				If (inputValid) Then
					For i = 0 To numDisks - 1

						DiskInfoDiskDetailDataGrid.Item(0, i).Value = (i + 1)
					Next
					DiskDetailPanel.Show()
					DiskInfoDiskDetailDataGrid.Rows(0).Cells(0).Selected = True
					'Appears to be the easiest way to get to cell 1 and enable editing on barcode scan. Programmatic methods don't seem to work.
					SendKeys.Send("{TAB}")
					DiskQtyCategoryGrid.Enabled = False
					DiskInfoSelectSNBox.Enabled = False
					DiskInfoQtysDoneButton.Text = "Re-Enter"
				End If
			End If
			If (inputValid) Then
				diskCatsMode = False
				SerNumBottom.Text = CStr(activeModule.CatNumber.SerialNumber)
			Else
				diskCatsMode = True
				SerNumBottom.Text = ""
			End If
		Else
			diskCatsMode = True
			DiskInfoDiskDetailDataGrid.Rows.Clear()
			DiskInfoUpdateSums()
			DiskDetailPanel.Hide()
			DiskQtyCategoryGrid.Enabled = True
			DiskInfoSelectSNBox.Enabled = True
			DiskInfoQtysDoneButton.Text = "Done"
		End If



	End Sub

	Private Sub DiskInfoDiskDetailDataGrid_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiskInfoDiskDetailDataGrid.CurrentCellChanged
		Dim grid As DataGridView = DirectCast(sender, DataGridView)
		If Not grid.CurrentCell Is Nothing And Not justFailedDiskCheck Then
			If grid.CurrentCell.ColumnIndex = 0 Or grid.CurrentCell.ColumnIndex >= 3 Then
				SendKeys.Send("{TAB}")
			End If
			justFailedDiskCheck = False
		Else
			justFailedDiskCheck = False
		End If
	End Sub

	Private Sub DiskInfoDiskDetailDataGrid_CellLeave(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DiskInfoDiskDetailDataGrid.CellEnter, DiskInfoDiskDetailDataGrid.CellEndEdit
		If (Not diskCatsMode) Then
			Dim grid As CustomDataGridView = DirectCast(sender, CustomDataGridView)
			'If (grid.CurrentCell.ColumnIndex = grid.Columns.IndexOf(DiskTypeGridColumn) OrElse grid.CurrentCell.ColumnIndex = grid.Columns.IndexOf(BatchGridColumn)) Then

			'8May2012 RPost 
			'' '''Dim foundBarCode As Boolean
			'' ''foundBarCode = FillBarCodes(grid)
			' '' '' If grid.CurrentCellAddress.X <> grid.Columns.IndexOf(DiskNumberGridColumn) And grid.CurrentCellAddress.X <> grid.Columns.IndexOf(BatchGridColumn) Then    ' 23Apr2012 DLimburg commented out this test and replaced with test of foundBarCode = true, the next two lines had already been commented out previously by Forrest
			' '' ''  grid.CurrentRow.Cells("DiskNumberGridColumn").Value = ""
			' '' ''End If


			'' ''If foundBarCode = True Then ' 23Apr2012 DLimburg - added test of foundBarCode instead of test above, seems more logical
			'' ''    UpdateDiskDetails(grid, foundBarCode)
			'' ''End If

			ParseCode(grid.CurrentRow)
			UpdateDiskDetails(grid, True)


		End If
	End Sub

	Private Sub SelectWOWorkOrderTextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SelectWOWorkOrderTextBox.KeyPress
	'14 May 2012 RPost -Reset Manual entry is WO is changed
	ManualWorkOrderMode = False
		If (e.KeyChar = vbCr) Then
			nextButton_Click(sender, e)
		End If

	End Sub

	Private Sub SelectUnitSelectUnitBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SelectUnitSelectUnitBox.KeyPress
		If (e.KeyChar = vbCr) Then
			nextButton_Click(sender, e)
		End If
	End Sub

	Private Sub NumDisksNumDisksAT32Gray_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles NumDisksNumDisksAT32Gray.KeyPress, NumDisksNumDisksAT52Gray.KeyPress, NumDisksNumDisksAT52Calcined.KeyPress, NumDisksNumDisksAT32Calcined.KeyPress, NumDisksTotalDisks.KeyPress, UnitReportReportRichTextBox.KeyPress, ManualWONumPartsTextBox.KeyPress
		If (e.KeyChar = vbCr) Then
			nextButton_Click(sender, e)
		End If
	End Sub

	Private Sub DiskQtyCategoryGrid_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DiskQtyCategoryGrid.KeyPress
		If (e.KeyChar = vbCr) Then
			DiskInfoQtysDoneButton_Click(sender, e)
		End If
	End Sub


	Private Sub CancelManualWO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelManualWO.Click
		SelectWOWorkOrderTextBox.Clear()
		ManualWOContainer.Hide()
	End Sub

	Private Sub LoginTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginTimer.Tick
		DateBottom.Text = DateTime.Now.ToString()
		If (LoginAlive = True) Then
			LoginAlive = False
		Else
			LogoutToolStripMenuItem_Click(sender, e)
		End If
	End Sub

	Private Sub EntryCancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EntryCancelButton.Click
		primaryWorkOrder = New WorkOrder()
		activeModule = Nothing
		If (diskCatsMode = False) Then
			DiskInfoQtysDoneButton_Click(sender, e)
		End If
		ChangeState(States.woSelect)
	End Sub

	<DebuggerHidden()>
	Private Sub ClockTickTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClockTickTimer.Tick
		DateBottom.Text = DateTime.Now.ToString()
	End Sub

	Private Sub frmMain_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
		loginForm.Close()
	End Sub

	Private Sub DiskInfoSelectSNBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiskInfoSelectSNBox.SelectedIndexChanged
		DiskQtyCategoryGrid.Focus()
	End Sub

	Private Sub DiskQtyCategoryGrid_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DiskQtyCategoryGrid.Enter
		Dim grid As DataGridView = DirectCast(sender, DataGridView)
		If Not grid.CurrentCell Is Nothing Then
			If grid.CurrentCell.ColumnIndex = 0 Or grid.CurrentCell.ColumnIndex >= 3 Then
				SendKeys.Send("{TAB}")
			End If
		End If
	End Sub

	Private checkPrint As Integer

	Private Sub BuildReportDocPrinter_BeginPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles BuildReportDocPrinter.BeginPrint
		checkPrint = 0
	End Sub

	Private Sub BuildReportDocPrinter_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles BuildReportDocPrinter.PrintPage
		' Print the content of the RichTextBox. Store the last character printed.
		checkPrint = UnitReportReportRichTextBox.Print(checkPrint, UnitReportReportRichTextBox.TextLength, e)
		' Look for more pages
		If checkPrint < UnitReportReportRichTextBox.TextLength Then
			e.HasMorePages = True
		Else
			e.HasMorePages = False
		End If
	End Sub

	Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
		Dim settingsForm As New Settings
		settingsForm.Show()
	End Sub

	Private Sub DeleteWorkOrderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteWorkOrderToolStripMenuItem.Click
		Dim deleteWoForm As New DeleteWorkOrder
		deleteWoForm.Show()
	End Sub

	Private Sub UnbuildUnitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnbuildUnitToolStripMenuItem.Click
		Dim UnbuildModuleForm As New UnbuildModule
		UnbuildModuleForm.Show()
	End Sub

	Private Sub miPrintUnitReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles miPrintUnitReport.Click
		Dim PrintUnitForm As New PrintUnitReport(Me)
		PrintUnitForm.Show()
	End Sub

	Private Sub miPrintMultipleReports_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miPrintMultipleReports.Click
		Dim PrintMultUnitReport As New PrintMultReport(Me)
		PrintMultUnitReport.Show()
	End Sub

#End Region

#Region "Validation Functions"
	Private Function SelectWOPageValidation() As Boolean
		Dim woNum As Integer
		primaryWorkOrder = New WorkOrder()
		If (Integer.TryParse(SelectWOWorkOrderTextBox.Text, woNum)) Then
			primaryWorkOrder.WorkOrderNum = SelectWOWorkOrderTextBox.Text
		Else
			MsgBox("Work order must be numeric!")
			Return False
		End If

		If (primaryWorkOrder.Validate() And Not primaryWorkOrder.Exists()) Then
			ManualWOContainer.Show()

			ManualWOCatNumTextBox.Clear()
			ManualWOFirstSNTextBox.Clear()
			ManualWONumPartsTextBox.Clear()

			ManualWOCatNumTextBox.Focus()
			ManualWorkOrderMode = True

			Return False
		End If

		
		SelectUnitWorkOrder.Text = primaryWorkOrder.WorkOrderNum
		WorkOrderBottom.Text = primaryWorkOrder.WorkOrderNum
		SelectUnitType.Text = primaryWorkOrder.CatNum.TypeString
		SelectUnitFamily.Text = CStr(primaryWorkOrder.CatNum.Family)
		SelectUnitCatNum.Text = primaryWorkOrder.CatNum.CatNumString
		SelectUnitNumRequiredTextBox.Text = CStr(primaryWorkOrder.NumUnits)
		NumDisksType.Text = primaryWorkOrder.CatNum.TypeString
		NumDisksFamily.Text = CStr(primaryWorkOrder.CatNum.Family)
		NumDisksCatNum.Text = primaryWorkOrder.CatNum.CatNumString


		Return True
	End Function

	Private Function NumDisksPageValidation() As Boolean
		Dim totDisksTmp As Integer = 0
		Dim at32grTmp As Integer = 0
		Dim at32ccTmp As Integer = 0
		Dim at52grTmp As Integer = 0
		Dim at52ccTmp As Integer = 0

		Integer.TryParse(NumDisksTotalDisks.Text, totDisksTmp)
		Integer.TryParse(NumDisksNumDisksAT32Gray.Text, at32grTmp)
		Integer.TryParse(NumDisksNumDisksAT32Calcined.Text, at32ccTmp)
		Integer.TryParse(NumDisksNumDisksAT52Gray.Text, at52grTmp)
		Integer.TryParse(NumDisksNumDisksAT52Calcined.Text, at52ccTmp)

		activeModule.TotalDisks = totDisksTmp
		activeModule.Num3kVDisksGray = at32grTmp
		activeModule.Num3kVDisksCalcined = at32ccTmp
		activeModule.Num5kVDisksGray = at52grTmp
		activeModule.Num5kVDisksCalcined = at52ccTmp
		Return activeModule.verifyDiskCounts(primaryWorkOrder)
	End Function

	Private Function DiskDetailValidation() As Boolean
		DiskInfoUpdateSums()
		Dim retVal As Boolean = True
		Dim row As DataGridViewRow
		qtysList.Clear()
		For Each row In DiskQtyCategoryGrid.Rows
			If (IsNumeric(row.Cells.Item("Quantity").Value) And IsNumeric(row.Cells("Category").Value)) Then
				Dim mapping As New CatQtyMapping(CInt(row.Cells("Category").Value), CInt(row.Cells.Item("Quantity").Value), 0, Nothing)
				If (qtysList.Contains(mapping)) Then
					Dim foundMapping As CatQtyMapping = qtysList.Item(qtysList.IndexOf(mapping))
					foundMapping.assertedQty += mapping.assertedQty

					qtysList.RemoveAt(qtysList.IndexOf(mapping))
					qtysList.Add(foundMapping)
				Else
					qtysList.Add(mapping)
				End If
			End If
		Next


		For Each row In DiskInfoDiskDetailDataGrid.Rows
			If (Not row.Cells.Item(1).Value Is Nothing) Then

				Dim tmpBatch As String = row.Cells.Item(1).Value.ToString()
				Dim tmpDiskNum As Integer
				Dim catToFind As CatQtyMapping
				If (Not (row.Cells.Item(1).Value Is Nothing Or row.Cells.Item(2).Value Is Nothing)) Then
					If (Integer.TryParse(row.Cells.Item(2).Value.ToString, tmpDiskNum)) Then
						catToFind = New CatQtyMapping(CInt(row.Cells("DiskCategoryGridColumn").Value), 0, 0, CStr(row.Cells("DiskTypeGridColumn").Value))
						If (Not qtysList.Contains(catToFind)) Then
							MsgBox("Category not specified above: " & CStr(row.Cells("DiskCategoryGridColumn").Value), MsgBoxStyle.Critical)
							retVal = False
						Else
							Dim tmpItem As CatQtyMapping
							tmpItem = qtysList.Item(qtysList.IndexOf(catToFind))
							qtysList.Remove(tmpItem)
							tmpItem.actualQty = tmpItem.actualQty + 1
							tmpItem.diskName = CStr(row.Cells("DiskTypeGridColumn").Value)
							qtysList.Add(tmpItem)
						End If
					ElseIf (retVal = True) Then
						MsgBox("Must enter a valid numeric value for Disk Number!", MsgBoxStyle.Critical)
						retVal = False
					End If
				ElseIf (retVal = True) Then
					MsgBox("Must enter data for all disks!", MsgBoxStyle.Critical)
					retVal = False
				End If
			Else
				MsgBox("You must enter disk information for every row!")
				retVal = False
			End If
		Next row

		'Make sure that the summation of the disks entered in the disk detail equals the number of disks for that category
		'As given in the DiskQtyCategoryGrid At the top of the page.
		If (retVal = True) Then
			Dim item As CatQtyMapping
			For Each item In qtysList
				If (retVal And item.actualQty <> item.assertedQty) Then
					If (retVal) Then
						MsgBox("Category " & item.category & ":Expected quantity " & item.assertedQty & " does not match actual quantity: " & item.actualQty)
						retVal = False
					End If
				End If
			Next
		End If


        If (retVal = True) Then
            ' DLL & MMR added if statement to test for new polymer 3Nov2014
            ' Code was duplicated in the Then and Else portions of If statement, this could be cleaned up at some point.

            If (primaryWorkOrder.CatNum.BuildPath = CatNum.BuildPaths.UltraSIL Or primaryWorkOrder.CatNum.BuildPath = CatNum.BuildPaths.DualRatedPolymerNew) Then
				Dim testParamsQuery As String = "SELECT * FROM ArresterTestParameters " &
				"WHERE ArresterFamily = '" & primaryWorkOrder.CatNum.Family & "' " &
				"AND IR_Code = '" & primaryWorkOrder.CatNum.IRCode & "' " &
				"AND CantileverVentingCode = '" & primaryWorkOrder.CatNum.CantileverStrength & "' " & _    'MMR 3/24/26 need to add cantilever venting code
				"AND MinRating <= " & primaryWorkOrder.CatNum.Rating & " " &
				"AND MaxRating >= " & primaryWorkOrder.CatNum.Rating & " "

				Dim testParamsCommand As New SqlCommand(testParamsQuery, PolymerStationConnection)
                Dim testParamsResult As SqlDataReader = testParamsCommand.ExecuteReader()
                Try
                    If (testParamsResult.HasRows()) Then
                        testParamsResult.Read()
                        Dim DesignMCOV As Double = CDbl(activeModule.ModuleMCOV) * 1.0
                        Dim VrefLow As Double = CDbl(testParamsResult!MinVrefFactor) * DesignMCOV
                        Dim VrefHigh As Double = CDbl(testParamsResult!MAxVrefFactor) * DesignMCOV
                        Dim V1MASum As Double
                        Double.TryParse(DiskInfoVRefSum.Text, V1MASum)

                        'If (V1MASum < VrefLow Or V1MASum > VrefHigh) Then
                        'MsgBox("The Vref:" & V1MASum & " is not between the limits of " & VrefLow & " And " & _
                        'VrefHigh & ". This module cannot be built.", MsgBoxStyle.Critical)
                        'retVal = False
                        'End If
                        'Per MMR removed Vref sanity check due to disk and arrester difference in current magnitude for iRef


                    Else
                        MsgBox("Could not find the vRef limits for this part in ArresterTestParameters", MsgBoxStyle.Critical)
                        retVal = False
                    End If
                Finally
                    testParamsResult.Close()
                End Try
            Else
                Dim testParamsQuery As String = "SELECT * FROM ArresterTestParameters " & _
                "WHERE ArresterFamily = '" & primaryWorkOrder.CatNum.Family & "' " & _
                "AND MinRating <= " & primaryWorkOrder.CatNum.Rating & " " & _
                "AND MaxRating >= " & primaryWorkOrder.CatNum.Rating & " "

                Dim testParamsCommand As New SqlCommand(testParamsQuery, GRoomConnection)
                Dim testParamsResult As SqlDataReader = testParamsCommand.ExecuteReader()

                Try
                    If (testParamsResult.HasRows()) Then
                        testParamsResult.Read()
                        Dim DesignMCOV As Double = CDbl(activeModule.ModuleMCOV) * 1.0
                        Dim VrefLow As Double = CDbl(testParamsResult!MinVrefFactor) * DesignMCOV
                        Dim VrefHigh As Double = CDbl(testParamsResult!MAxVrefFactor) * DesignMCOV
                        Dim V1MASum As Double
                        Double.TryParse(DiskInfoVRefSum.Text, V1MASum)

                        'If (V1MASum < VrefLow Or V1MASum > VrefHigh) Then
                        'MsgBox("The Vref:" & V1MASum & " is not between the limits of " & VrefLow & " And " & _
                        'VrefHigh & ". This module cannot be built.", MsgBoxStyle.Critical)
                        'retVal = False
                        'End If
                        'Per MMR removed Vref sanity check due to disk and arrester difference in current magnitude for iRef


                    Else
                        MsgBox("Could not find the vRef limits for this part in ArresterTestParameters", MsgBoxStyle.Critical)
                        retVal = False
                    End If
                Finally
                    testParamsResult.Close()
                End Try
            End If
        End If

            If (retVal AndAlso primaryWorkOrder.CatNum.Type = CatNum.Types.PORC) Then
                Dim GlassDisks As New List(Of Disk)
                Dim msgBoxString As String = "The following disks are glass disks in the database and " & _
                 "should only be used if they are epoxy disks. Click OK to confirm or Cancel to not use these disks."
                For Each diskElement As Disk In activeModule.Disks
                    If (diskElement.EdgeType = Disk.EdgeTypes.Glass) Then
                        GlassDisks.Add(diskElement)
                        msgBoxString &= vbLf & diskElement.Batch & "    " & diskElement.DiskNum
                    End If
                Next diskElement
                If (GlassDisks.Count > 0) Then
                    Dim result As MsgBoxResult = MsgBox(msgBoxString, vbYesNo)
                    If (result = vbYes) Then

                    Else
                        retVal = False
                    End If

                End If
            End If


            If (retVal = True) Then
                retVal = activeModule.Build(loginObject)
            End If


            Return retVal
	End Function

#End Region

#Region "Action Functions"

	Public Sub ChangeState(ByVal newState As Integer)
		woSelectLabel.Font = New Font(woSelectLabel.Font, FontStyle.Regular)
		unitSelLabel.Font = New Font(woSelectLabel.Font, FontStyle.Regular)
		numDisksLabel.Font = New Font(woSelectLabel.Font, FontStyle.Regular)
		diskInfoLabel.Font = New Font(woSelectLabel.Font, FontStyle.Regular)
		unitReportLabel.Font = New Font(woSelectLabel.Font, FontStyle.Regular)
		nextButton.Enabled = True
		prevButton.Enabled = True
		nextButton.Text = "Next"
		If (newState = States.unitReport) Then
			unitReportLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
			nextButton.Enabled = True
			PagesTabControl.SelectedTab = UnitReportPage
			nextButton.Text = "Finish"
			prevButton.Enabled = False
			EntryCancelButton.Visible = False
			'RPost Why is this being set?
			NumDisksTotalDisks.Text = ""
			NumDisksNumDisksAT32Gray.Text = ""
			NumDisksNumDisksAT32Calcined.Text = ""
			NumDisksNumDisksAT52Gray.Text = ""
			NumDisksNumDisksAT52Calcined.Text = ""

		ElseIf (newState = States.diskInfo) Then
			diskInfoLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
			PagesTabControl.SelectedTab = DiskInfoPage
			FillSNBox()
			'Don't clear if returning from a unit report -15May2012 RPost
			If currState <> States.unitReport AndAlso currState <> States.diskInfo Then
				DiskQtyCategory_Init()
			End If
			DiskInfoSelectSNBox.Select()
			EntryCancelButton.Visible = True
			DiskDetailLabel.Text = "Provide Disk Quantities and Categories (Unit: " & activeModule.Unit.UnitLetter & ", Module: " & activeModule.ModuleNumber & ")"
		ElseIf (newState = States.numDisks) Then
			numDisksLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
			PagesTabControl.SelectedTab = NumDisksPage
			NumDisksTotalDisks.Select()
			SetActiveModule()
			NumDisksAT32Label.Text = "Number of " & primaryWorkOrder.CatNum.DiskType & " Disks"
			NumDisksAT52Label.Text = "Number of " & primaryWorkOrder.CatNum.PartnerDiskType & " Disks"
			UnitBottom.Text = activeModule.Unit.UnitLetter
			ModuleBottom.Text = CStr(activeModule.ModuleNumber)
			EntryCancelButton.Visible = True

			' 12Apr2012 DLimburg Added test for New style polymer, this is the only arrester that does not allow the use of gray disks
			NumDisksNumDisksAT52Gray.Enabled = (Len(primaryWorkOrder.CatNum.CatNumString) <> 18)
			NumDisksNumDisksAT32Gray.Enabled = NumDisksNumDisksAT52Gray.Enabled



			NumDisksTotalDisks.Text = ""		' Added line to clear disk quantity on entry (DLimburg 27Mar2012)
			NumDisksNumDisksAT52Gray.Text = ""
			NumDisksNumDisksAT52Calcined.Text = ""		  ' Added line to clear disk quantity on entry (DLimburg 27Mar2012)
			NumDisksNumDisksAT32Gray.Text = ""
			NumDisksNumDisksAT32Calcined.Text = ""		  ' Added line to clear disk 

			'14May2012 RPost AutoPopulate
			Dim sql As String = _
				String.Format("SELECT BuiltUnits.NumDisks, BuiltUnits.NumOf3kVDisks, BuiltUnits.NumOf5kVDisks, BuiltUnits.NumOf3kVDisksCalcined, BuiltUnits.NumOf5kVDisksCalcined " & _
							"FROM WO_SerialNumbers INNER JOIN BuiltUnits ON WO_SerialNumbers.SerialNumber = BuiltUnits.SerialNumber " & _
							"WHERE WO_Number='{0}' AND Unit={1}", SelectUnitWorkOrder.Text, GetModuleFromString(CStr(SelectUnitSelectUnitBox.SelectedItem)).OldModNumber)

			Using rdr As SqlDataReader = New SqlCommand(sql, GRoomConnection).ExecuteReader()
				If (rdr.HasRows()) Then
					rdr.Read()

					NumDisksTotalDisks.Text = rdr!NumDisks.ToString
					NumDisksNumDisksAT52Gray.Text = rdr!NumOf5kVDisks.ToString
					NumDisksNumDisksAT52Calcined.Text = rdr!NumOf5kVDisksCalcined.ToString
					NumDisksNumDisksAT32Gray.Text = rdr!NumOf3kVDisks.ToString
					NumDisksNumDisksAT32Calcined.Text = rdr!NumOf3kVDisksCalcined.ToString
				End If
			End Using

		ElseIf (newState = States.unitSel) Then
			PagesTabControl.SelectedTab = SelectUnitPage
			SelectUnitSelectUnitBox.Select()
			unitSelLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
			My.Settings.LastWorkOrder = SelectWOWorkOrderTextBox.Text
			FillUnitsBox()
			EntryCancelButton.Visible = True
		ElseIf (newState = States.woSelect) Then
			prevButton.Enabled = False
			SelectWOWorkOrderTextBox.Text = ""
			SelectWOPreviousWorkOrder.Text = My.Settings.LastWorkOrder
			woSelectLabel.Font = New Font(woSelectLabel.Font, FontStyle.Bold)
			PagesTabControl.SelectedTab = SelectWOPage
			EntryCancelButton.Visible = False
			primaryWorkOrder = New WorkOrder()
			activeModule = Nothing
		End If

		currState = newState
	End Sub


	Delegate Sub SetColumnIndex(ByVal i As Integer)

	Private Sub ThreadSafeCellChange(ByVal columnIndex As Integer)
		If (DiskInfoDiskDetailDataGrid.CurrentRow.Index - 1 >= 0) Then
			Me.DiskInfoDiskDetailDataGrid.CurrentCell = Me.DiskInfoDiskDetailDataGrid.Rows(DiskInfoDiskDetailDataGrid.CurrentRow.Index - 1).Cells(columnIndex)

			Me.DiskInfoDiskDetailDataGrid.BeginEdit(True)
		End If


	End Sub

	Private Sub GenerateReport()
		Dim report As New Report(primaryWorkOrder)
		UnitReportReportRichTextBox.Rtf = report.GetReport()
		If (My.Settings.AutoPrintReport) Then
			BuildReportDocPrinter.Print()
		End If

	End Sub

	Private Function FillBarCodes(ByVal grid As CustomDataGridView) As Boolean
		Dim row As DataGridViewRow
		Dim retVal As Boolean = False
		For Each row In grid.Rows
			'If Batch and Grid have something in them
			If (Not row.Cells("BatchGridColumn") Is Nothing And Not row.Cells("DiskNumberGridColumn") Is Nothing) Then

				'If ((Not row.Cells("BatchGridColumn").Value Is Nothing AndAlso row.Cells("BatchGridColumn").Value.ToString.Contains("D") AndAlso row.Cells("BatchGridColumn").Value.ToString.Length = 30) OrElse _
				'(Not row.Cells("DiskNumberGridColumn").Value Is Nothing AndAlso row.Cells("DiskNumberGridColumn").Value.ToString.Contains("D") AndAlso row.Cells("DiskNumberGridColumn").Value.ToString.Length = 30)) Then '19Apr2012 DLimburg
				If (Not row.Cells("BatchGridColumn").Value Is Nothing AndAlso row.Cells("BatchGridColumn").Value.ToString.Contains("D")) Then
					Dim cellStr As String
					'If (Not row.Cells("BatchGridColumn").Value Is Nothing AndAlso row.Cells("BatchGridColumn").Value.ToString.Contains("D")) Then  ' removed code DLimburg 1May12
					cellStr = CStr(row.Cells("BatchGridColumn").Value)
					'Else    ' removed code DLimburg 1May12
					'    cellStr = CStr(row.Cells("DiskNumberGridColumn").Value)    ' removed code DLimburg 1May12
					'End If    ' removed code DLimburg 1May12

					Dim startIndex As Integer = row.Cells("BatchGridColumn").ToString.IndexOf("D")
					Dim diskBatch As String
					Dim diskNum As Integer
					'Dim category As Integer    ' removed code DLimburg 1May12
					'Dim categoryStr As String    ' removed code DLimburg 1May12
					'Dim scpVolts As Integer    ' removed code DLimburg 1May12
					'Dim v10kA As Integer    ' removed code DLimburg 1May12
					'Dim v10kAStr As String    ' removed code DLimburg 1May12
					'Dim vRef As Integer    ' removed code DLimburg 1May12
					'Dim diskType As String    ' removed code DLimburg 1May12
					diskBatch = cellStr.Substring(startIndex + 2, 5)
					Integer.TryParse(cellStr.Substring(startIndex + 7, 4), diskNum)
					'categoryStr = cellStr.Substring(startIndex + 11, 4)    ' removed code DLimburg 1May12
					'If categoryStr.Substring(0, 1) = "L" Or categoryStr.Substring(0, 1) = "H" Then  ' 27Apr2012 DLimburg added this test to see if this will resolve disk parsing issue
					'    Integer.TryParse(cellStr.Substring(startIndex + 12, 3), category)    ' removed code DLimburg 1May12
					'Else    ' removed code DLimburg 1May12
					'    Integer.TryParse(cellStr.Substring(startIndex + 11, 4), category)    ' removed code DLimburg 1May12
					'End If    ' removed code DLimburg 1May12
					'Integer.TryParse(cellStr.Substring(startIndex + 15, 4), scpVolts)    ' removed code DLimburg 1May12
					'v10kAStr = cellStr.Substring(startIndex + 19, 3)    ' removed code DLimburg 1May12
					'If (v10kAStr.Substring(0, 1) = "X") Then    ' removed code DLimburg 1May12
					'    Integer.TryParse(v10kAStr.Substring(1), v10kA)    ' removed code DLimburg 1May12
					'    v10kA = v10kA + 1000    ' removed code DLimburg 1May12
					'Else    ' removed code DLimburg 1May12
					'    Integer.TryParse(v10kAStr, v10kA)    ' removed code DLimburg 1May12
					'End If    ' removed code DLimburg 1May12
					'Integer.TryParse(cellStr.Substring(startIndex + 22, 3), vRef)    ' removed code DLimburg 1May12
					' diskType = cellStr.Substring(startIndex + 25, 4)    ' removed code DLimburg 1May12
					grid.CurrentRow.Cells("BatchGridColumn").Value = diskBatch
					grid.CurrentRow.Cells("DiskNumberGridColumn").Value = diskNum
					SendKeys.Send("{Tab}")
					retVal = True
				End If

				'Batch and Disk num are populated but manual entry
				'No validation of manual entries?
				retVal = True 'RPost 7May2012


			End If





		Next row
		Return retVal
	End Function


'This code replaced FillBarCodes  -RPost
	Private Sub ParseCode(ByVal row As DataGridViewRow)

		If (Not row.Cells("BatchGridColumn").Value Is Nothing) Then
			Dim code As String
			code = row.Cells("BatchGridColumn").Value.ToString.ToUpper
			' Replaced the following line RVA 4/6/2023
			'If code.Contains("D") Then
			'Scanned code
			If code.Length > 25 AndAlso code.Contains("D") Then
				Dim startIndex As Integer = code.IndexOf("D")
				Dim diskNum As Integer

				row.Cells("BatchGridColumn").Value = Trim(code.Substring(startIndex + 1, 6)) 'Changed from substring(ndx,2,5) and added trim
				If (Integer.TryParse(code.Substring(startIndex + 7, 4), diskNum)) Then
					row.Cells("DiskNumberGridColumn").Value = diskNum
				Else
					'Failed to parse disknum
					row.Cells("DiskNumberGridColumn").Value = "FAIL"
				End If

				SendKeys.Send("{Tab}")
			End If
		End If
	End Sub


	Private Sub UpdateDiskDetails(ByVal grid As DataGridView, ByVal FoundBarCode As Boolean)
		activeModule.ClearDisks()
		Dim row As DataGridViewRow
		For Each row In grid.Rows
			If (Not row.Cells("BatchGridColumn") Is Nothing AndAlso Not row.Cells("DiskNumberGridColumn") Is Nothing) Then
				If (IsNumeric(row.Cells("DiskNumberGridColumn").Value)) Then
					Dim cellStr As String = CStr(grid.Rows(grid.CurrentCell.RowIndex).Cells(grid.CurrentCell.ColumnIndex).Value)
					If (Not row.Cells("BatchGridColumn").Value Is Nothing) AndAlso (Not row.Cells("DiskNumberGridColumn").Value Is Nothing) Then
						If row.Cells("DiskNumberGridColumn").Value.ToString.Length <= 10 AndAlso row.Cells("BatchGridColumn").Value.ToString.Length <= 10 Then



							Dim addedDisk As New Disk(activeModule, CStr(row.Cells("BatchGridColumn").Value), CInt(row.Cells("DiskNumberGridColumn").Value))
							row.Cells("BatchGridColumn").Value = addedDisk.Batch
							row.Cells("DiskNumberGridColumn").Value = addedDisk.DiskNum
							row.Cells("VrefGridColumn").Value = addedDisk.VRef
							row.Cells("ActV10kAGridColumn").Value = addedDisk.V10kA
							row.Cells("DiskCategoryGridColumn").Value = addedDisk.Category
							row.Cells("DiskTypeGridColumn").Value = addedDisk.DiskType
							If (Not CBool(addedDisk.ValidForUse) Or Not activeModule.AddDisk(addedDisk)) Then
								row.Cells("BatchGridColumn").Value = Nothing
								row.Cells("DiskNumberGridColumn").Value = Nothing
								row.Cells("VrefGridColumn").Value = Nothing
								row.Cells("ActV10kAGridColumn").Value = Nothing
								row.Cells("DiskCategoryGridColumn").Value = Nothing
								row.Cells("DiskTypeGridColumn").Value = Nothing
								Threading.Thread.Sleep(20)
								'  If (Not grid.CurrentCell.RowIndex = grid.Rows.Count - 1) Then
								' SendKeys.Send("{UP}")
								'End If
								'If (grid.CurrentCell.ColumnIndex = grid.Columns.IndexOf(DiskNumberGridColumn)) Then
								SendKeys.Send("{LEFT}")
								If (Not FoundBarCode) Then
									SendKeys.Send("{LEFT}")
								End If

								'End If
								justFailedDiskCheck = True
								'SendKeys.Send("{LEFT}")
								'                        grid.Rows(grid.CurrentCell.RowIndex + 1).Cells("BatchGridColumn").Selected = False
								'If Me.DiskInfoDiskDetailDataGrid.CurrentCell.ColumnIndex <> Me.DiskInfoDiskDetailDataGrid.Columns.Count - 1 Then

								'Dim nextindex As Integer = Math.Min(Me.DiskInfoDiskDetailDataGrid.Columns.Count - 1, Me.DiskInfoDiskDetailDataGrid.CurrentCell.ColumnIndex + 1)

								' Dim method As New SetColumnIndex(AddressOf ThreadSafeCellChange)
								' DiskInfoDiskDetailDataGrid.ClearSelection()
								' DiskInfoDiskDetailDataGrid.CurrentCell = DiskInfoDiskDetailDataGrid.Rows(0).Cells(0)
								' Me.DiskInfoDiskDetailDataGrid.BeginInvoke(method, 1)
								'End If
							End If

						Else
							MsgBox("Disk number or batch number too large!")
							row.Cells("BatchGridColumn").Value = ""
							row.Cells("DiskNumberGridColumn").Value = ""
							SendKeys.Send("{LEFT}")
							SendKeys.Send("{UP}")
						End If
					Else
						row.Cells("VrefGridColumn").Value = ""
						row.Cells("ActV10kAGridColumn").Value = ""
						row.Cells("DiskCategoryGridColumn").Value = ""
						row.Cells("DiskTypeGridColumn").Value = ""
					End If
				Else
					row.Cells("VrefGridColumn").Value = ""
					row.Cells("ActV10kAGridColumn").Value = ""
					row.Cells("DiskCategoryGridColumn").Value = ""
					row.Cells("DiskTypeGridColumn").Value = ""
				End If

			End If
		Next row
		DiskInfoUpdateSums()
		'DiskInfoDiskDetailDataGrid.ClearSelection()
		'Dim method As New SetColumnIndex(AddressOf ThreadSafeCellChange)
		'Me.DiskInfoDiskDetailDataGrid.BeginInvoke(method, 1)
		' DiskInfoDiskDetailDataGrid.CurrentCell = DiskInfoDiskDetailDataGrid.Rows(0).Cells(3)

	End Sub



	Private Function CommitManualWO() As Boolean
		Dim success As Boolean = True
		Dim woNumber As Integer
		Dim firstSN As Long
		Dim numUnits As Integer
		If (Not Integer.TryParse(SelectWOWorkOrderTextBox.Text, woNumber)) Then
			MsgBox("Entered work order is not a number!", MsgBoxStyle.Critical)
			success = False
		End If

        ' 2019-01-22 RVA Allow 11 or 12 digits
        If (success And ManualWOFirstSNTextBox.Text.Length <> 12 And ManualWOFirstSNTextBox.Text.Length <> 11) Then
            MsgBox("Serial number must be 11 or 12 digits!")
            success = False
        End If

		If (success And Not Long.TryParse(ManualWOFirstSNTextBox.Text, firstSN)) Then
			MsgBox("Entered first serial number is not a number!", MsgBoxStyle.Critical)
			success = False
		End If

		If (success And Not Integer.TryParse(ManualWONumPartsTextBox.Text, numUnits)) Then
			MsgBox("Entered number of units is not a number!", MsgBoxStyle.Critical)
			success = False
		End If

		If (success) Then
			primaryWorkOrder = New WorkOrder()
			primaryWorkOrder.WorkOrderNum = CStr(woNumber)
			primaryWorkOrder.CatNum.GivenCatNum = UCase(ManualWOCatNumTextBox.Text())  '   20Apr2012   DLimburg added this field to maintain the shortcut number to save to WO_Info
			primaryWorkOrder.CatNum.CatNumString = UCase(ManualWOCatNumTextBox.Text())	'   20Apr2012   Converted CatNumber to Uppercase in case of manual entry
			primaryWorkOrder.StartSerial = firstSN
			primaryWorkOrder.NumUnits = numUnits
			success = primaryWorkOrder.AddWorkOrder()
		End If

		Return success

	End Function

	Private Sub FillUnitsBox()
		SelectUnitSelectUnitBox.Items.Clear()
		Dim Units As List(Of ArresterUnit) = primaryWorkOrder.CatNum.Units
		Dim unit As ArresterUnit
		For Each unit In Units
			Dim arresterMod As ArresterModule
			For Each arresterMod In unit.Modules
				If unit.UnitMCOV > 0 Then	' added this test so as not unused modules (DLimburg 26Mar2012)
					If (primaryWorkOrder.numRequired > arresterMod.NumberBuilt) Then
						SelectUnitSelectUnitBox.Items.Add(arresterMod.moduleString())
					End If
				End If
			Next

		Next unit
		If SelectUnitSelectUnitBox.Items.Count > 0 Then
			SelectUnitSelectUnitBox.SelectedItem = SelectUnitSelectUnitBox.Items(0)
		End If


	End Sub

	Private Sub FillSNBox()
		'Get unbuilt
		DiskInfoSelectSNBox.Items.Clear()
		Dim serials() As Int64 = primaryWorkOrder.GetSerialNumbersUnBuilt(activeModule.OldModNumber)
		'7May2012 RPost
		Dim BuiltSNs As New List(Of Int64) '7May2012 RPost
		BuiltSNs = primaryWorkOrder.GetAllSerialNumbers() '7May2012 RPost

		Dim serial As Int64
		For Each serial In serials
			DiskInfoSelectSNBox.Items.Add(serial.ToString)
			'Get all, subtract unbuilt = built
			BuiltSNs.Remove(serial)	'7May2012 RPost
		Next serial


		'7May2012 RPost
		lbBuiltSN.Items.Clear()
		For Each serial In BuiltSNs
			lbBuiltSN.Items.Add(serial)
		Next


		If (DiskInfoSelectSNBox.Items.Count > 0) Then
			' DiskInfoSelectSNBox.SelectedItem = DiskInfoSelectSNBox.Items(0)
		Else
			MsgBox("No unbuilt serial numbers for this work order!", MsgBoxStyle.Critical)
		End If

	End Sub

	Private Function GetModuleFromString(ByVal modString As String) As ArresterModule
		Dim retVal As ArresterModule = Nothing
		Dim unit As ArresterUnit
		Dim Units As List(Of ArresterUnit) = primaryWorkOrder.CatNum.Units
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

	Private Sub SetActiveModule()
		activeModule = GetModuleFromString(SelectUnitSelectUnitBox.SelectedItem.ToString)
		'activeModule.setWorkOrder(primaryWorkOrder)
		NumDisksUnit.Text = activeModule.Unit.UnitLetter
		NumDisksModule.Text = CStr(activeModule.ModuleNumber)
	End Sub

	Private Sub DiskInfoUpdateSums()
		Dim VRefSum As Double = 0.0
		Dim V10kASum As Double = 0.0
		Dim categorySum As Integer = 0
		Dim numRows As Integer = 0

		Dim Row As DataGridViewRow
		For Each Row In DiskInfoDiskDetailDataGrid.Rows()
			Dim tmp As Double
			If (Not Row.Cells("VrefGridColumn").Value Is Nothing) Then
				Double.TryParse(CStr(Row.Cells("vrefGridColumn").Value), tmp)
				VRefSum += tmp
			End If
			If (Not Row.Cells("ActV10kAGridColumn").Value Is Nothing) Then
				Double.TryParse(CStr(Row.Cells("ActV10kAGridColumn").Value), tmp)
				V10kASum += tmp
				numRows += 1
			End If
			If (Not Row.Cells("DiskCategoryGridColumn").Value Is Nothing) Then
				Double.TryParse(CStr(Row.Cells("DiskCategoryGridColumn").Value), tmp)
				categorySum += CInt(tmp)
			End If
		Next
		DiskInfoVRefSum.Text = CStr(VRefSum)
		DiskInfoV10kASum.Text = CStr(V10kASum)
		DiskInfoCategorySum.Text = CStr(categorySum)
		If (numRows > 0) Then
			DiskInfoAvgV10kA.Text = (categorySum / (numRows * 1.0)).ToString("0.000")
		Else
			DiskInfoAvgV10kA.Text = "0"
		End If

	End Sub
#End Region

	Private Sub PrintButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintButton.Click
		BuildReportDocPrinter.Print()
	End Sub

	Private Sub miData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miViewDiskInfo.Click
		frmData.Show()
		frmData.tbWO.Text = SelectWOWorkOrderTextBox.Text
	End Sub


	Private Sub ManualWOCatNumTextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ManualWOCatNumTextBox.KeyPress
		If e.KeyChar = vbCrLf Then
			ManualWOFirstSNTextBox.Focus()

		End If
	End Sub

	Private Sub ManualWOFirstSNTextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ManualWOFirstSNTextBox.KeyPress
		If e.KeyChar = vbCrLf Then
			ManualWONumPartsTextBox.Focus()
		ElseIf e.KeyChar = "$"c OrElse e.KeyChar = " "c Then
			e.Handled = True
			ManualWONumPartsTextBox.Focus()
		End If
	End Sub

Private Sub SelectUnitSelectUnitBox_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SelectUnitSelectUnitBox.MouseDoubleClick
		nextButton.PerformClick()
End Sub

	Private Sub ManualWOCatNumTextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ManualWOCatNumTextBox.Leave
		primaryWorkOrder.CatNum.GivenCatNum = ManualWOCatNumTextBox.Text
		primaryWorkOrder.CatNum.CatNumString = ManualWOCatNumTextBox.Text
	End Sub
End Class