<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintUnitReport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintUnitReport))
        Me.SerialNumberTextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CancelPrintUnitButton = New System.Windows.Forms.Button()
        Me.ViewUnitReportButton = New System.Windows.Forms.Button()
        Me.PrintUnitReportButton = New System.Windows.Forms.Button()
        Me.PrintUnitReportDocument = New System.Drawing.Printing.PrintDocument()
        Me.SuspendLayout()
        '
        'SerialNumberTextBox
        '
        Me.SerialNumberTextBox.Location = New System.Drawing.Point(144, 21)
        Me.SerialNumberTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SerialNumberTextBox.MaxLength = 12
        Me.SerialNumberTextBox.Name = "SerialNumberTextBox"
        Me.SerialNumberTextBox.Size = New System.Drawing.Size(132, 22)
        Me.SerialNumberTextBox.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(35, 25)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Serial Number:"
        '
        'CancelPrintUnitButton
        '
        Me.CancelPrintUnitButton.Location = New System.Drawing.Point(16, 58)
        Me.CancelPrintUnitButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CancelPrintUnitButton.Name = "CancelPrintUnitButton"
        Me.CancelPrintUnitButton.Size = New System.Drawing.Size(100, 28)
        Me.CancelPrintUnitButton.TabIndex = 4
        Me.CancelPrintUnitButton.Text = "Cancel"
        Me.CancelPrintUnitButton.UseVisualStyleBackColor = True
        '
        'ViewUnitReportButton
        '
        Me.ViewUnitReportButton.Location = New System.Drawing.Point(124, 58)
        Me.ViewUnitReportButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ViewUnitReportButton.Name = "ViewUnitReportButton"
        Me.ViewUnitReportButton.Size = New System.Drawing.Size(100, 28)
        Me.ViewUnitReportButton.TabIndex = 5
        Me.ViewUnitReportButton.Text = "View"
        Me.ViewUnitReportButton.UseVisualStyleBackColor = True
        '
        'PrintUnitReportButton
        '
        Me.PrintUnitReportButton.Location = New System.Drawing.Point(232, 58)
        Me.PrintUnitReportButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PrintUnitReportButton.Name = "PrintUnitReportButton"
        Me.PrintUnitReportButton.Size = New System.Drawing.Size(100, 28)
        Me.PrintUnitReportButton.TabIndex = 6
        Me.PrintUnitReportButton.Text = "Print"
        Me.PrintUnitReportButton.UseVisualStyleBackColor = True
        '
        'PrintUnitReportDocument
        '
        '
        'PrintUnitReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(345, 101)
        Me.Controls.Add(Me.PrintUnitReportButton)
        Me.Controls.Add(Me.ViewUnitReportButton)
        Me.Controls.Add(Me.CancelPrintUnitButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SerialNumberTextBox)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "PrintUnitReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Single Report"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SerialNumberTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CancelPrintUnitButton As System.Windows.Forms.Button
    Friend WithEvents PrintUnitReportButton As System.Windows.Forms.Button
    Friend WithEvents ViewUnitReportButton As System.Windows.Forms.Button
    Friend WithEvents PrintUnitReportDocument As System.Drawing.Printing.PrintDocument
End Class
