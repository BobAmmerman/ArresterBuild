Option Strict On
Public Class PrintMultReport


    Dim units As New List(Of Integer)
    Dim mainForm As frmMain
    Dim reports As New List(Of String)
    Dim tempRTB As New RichTextBoxPrintCtrl.RichTextBoxPrintCtrl.RichTextBoxPrintCtrl
    Public Sub New(ByRef mainFormIn As frmMain)
        InitializeComponent()
        mainForm = mainFormIn
    End Sub

    Private Sub CloseMultUnitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseMultUnitButton.Click
        Me.Close()
    End Sub

    Private Sub AddSerialNumberButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddSerialNumberButton.Click
        Dim serialNumTmp As Int64
        Int64.TryParse(SerialNumTextBox.Text, serialNumTmp)
        Dim qtyTmp As Integer
        If (Integer.TryParse(QtyTextBox.Text, qtyTmp) And qtyTmp > 0 And qtyTmp < 500) Then
            Dim i As Integer
            For i = 0 To (qtyTmp - 1)
                If (SerialNumListBox.Items.Contains(serialNumTmp + i)) Then
                    MsgBox("Duplicate serial number " & serialNumTmp + i & "!")
                    ' 2019-01-22 RVA - Allow 12 digit serial numbers
                ElseIf (SerialNumTextBox.Text.Length < 10 Or SerialNumTextBox.Text.Length > 12 Or Not IsNumeric(SerialNumTextBox.Text)) Then
                    MsgBox("Serial number must be a number of 10, 11 or 12 digits.")
                Else
                    SerialNumListBox.Items.Add(serialNumTmp + i)
                End If
            Next i

            SerialNumTextBox.Text = ""
            QtyTextBox.Text = "1"
        Else
            MsgBox("You have entered an invalid quantity. Must be between 1 and 500.")
        End If
    End Sub

    Private Sub ViewButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewButton.Click

        If (CompileReports()) Then
            Dim row As String
            Dim output As String = ""

            For Each row In reports
                output = output & row
            Next

            mainForm.UnitReportReportRichTextBox.Rtf = output
            mainForm.ChangeState(frmMain.States.unitReport)
            Me.Close()
        End If
    End Sub

    Private Function CompileReports() As Boolean
        Dim success As Boolean = True
        reports.Clear()
        Dim serRow As String
        Dim tmpReport As Report
        For Each serRow In SerialNumListBox.Items
            Dim serialNumTmp As Int64
            Int64.TryParse(serRow, serialNumTmp)
            tmpReport = New Report(serialNumTmp)
            reports.Add(tmpReport.GetReport)
        Next serRow
        Return success
    End Function

    Private Sub ClearSerialNumbersButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearSerialNumbersButton.Click
        SerialNumListBox.Items.Clear()
        SerialNumTextBox.Clear()
        QtyTextBox.Clear()
    End Sub

    Private Sub RemoveSelectedSerialNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveSelectedSerialNumber.Click
        SerialNumListBox.Items.Remove(SerialNumListBox.SelectedItem)
    End Sub

    Private Sub PrintButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintButton.Click
        If (CompileReports()) Then
            Dim Row As String
            For Each Row In reports
                tempRTB.Rtf = Row
                MultReportPrintDoc.Print()
            Next Row
            Me.Close()
        End If
    End Sub

    Private checkPrint As Integer

    Private Sub MultReportPrintDoc_BeginPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles MultReportPrintDoc.BeginPrint
        checkPrint = 0
    End Sub

    Private Sub MultReportPrintDoc_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles MultReportPrintDoc.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        checkPrint = tempRTB.Print(checkPrint, tempRTB.TextLength, e)
        ' Look for more pages
        If checkPrint < tempRTB.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub SerialNumListBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles SerialNumListBox.SelectedIndexChanged

    End Sub
End Class