Option Strict On
Public Class PrintUnitReport

    Private mainForm As frmMain
    Private checkPrint As Integer
    Private tempRTB As New RichTextBoxPrintCtrl.RichTextBoxPrintCtrl.RichTextBoxPrintCtrl

    Public Sub New(ByRef mainFormIn As frmMain)
        InitializeComponent()
        mainForm = mainFormIn
    End Sub

    Private Sub ViewUnitReportButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewUnitReportButton.Click
        Dim serNumTmp As Int64
        If (Int64.TryParse(SerialNumberTextBox.Text, serNumTmp) _
            And SerialNumberTextBox.Text.Length >= 10 And SerialNumberTextBox.Text.Length <= 12) Then   'MMR changed  11 to 12 1/22/19
            Dim report As New Report(serNumTmp)
            mainForm.ChangeState(frmMain.States.unitReport)
            mainForm.UnitReportReportRichTextBox.Rtf = report.GetReport()
            Me.Close()
        Else
            MsgBox("Invalid serial number entered. The serial number must be 10, 11 or 12 digits.")    'MMR changed to from 10 and 11  to 10, 11 and 12 1/22/19
        End If


    End Sub

    Private Sub CancelPrintUnitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelPrintUnitButton.Click
        Me.Close()
    End Sub

    Private Sub PrintUnitReportButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintUnitReportButton.Click
        Dim serNumTmp As Int64
        Int64.TryParse(SerialNumberTextBox.Text, serNumTmp)
        Dim report As New Report(serNumTmp)
        tempRTB.Rtf = report.GetReport
        PrintUnitReportDocument.Print()
        Me.Close()
    End Sub

    Private Sub PrintUnitReportDocument_BeginPrint(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintUnitReportDocument.BeginPrint
        checkPrint = 0
    End Sub

    Private Sub PrintUnitReportDocument_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintUnitReportDocument.PrintPage
        ' Print the content of the RichTextBox. Store the last character printed.
        checkPrint = tempRTB.Print(checkPrint, tempRTB.TextLength, e)
        ' Look for more pages
        If checkPrint < tempRTB.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub PrintUnitReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class