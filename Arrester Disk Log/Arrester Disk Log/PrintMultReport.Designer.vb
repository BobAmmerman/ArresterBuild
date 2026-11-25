<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintMultReport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintMultReport))
        Me.SerialNumListBox = New System.Windows.Forms.ListBox()
        Me.QtyTextBox = New System.Windows.Forms.TextBox()
        Me.SerialNumTextBox = New System.Windows.Forms.TextBox()
        Me.AddSerialNumberButton = New System.Windows.Forms.Button()
        Me.ClearSerialNumbersButton = New System.Windows.Forms.Button()
        Me.StartSerialLabel = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.PrintButton = New System.Windows.Forms.Button()
        Me.ViewButton = New System.Windows.Forms.Button()
        Me.CloseMultUnitButton = New System.Windows.Forms.Button()
        Me.RemoveSelectedSerialNumber = New System.Windows.Forms.Button()
        Me.MultReportPrintDoc = New System.Drawing.Printing.PrintDocument()
        Me.SuspendLayout()
        '
        'SerialNumListBox
        '
        Me.SerialNumListBox.FormattingEnabled = True
        Me.SerialNumListBox.ItemHeight = 16
        Me.SerialNumListBox.Location = New System.Drawing.Point(16, 64)
        Me.SerialNumListBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SerialNumListBox.Name = "SerialNumListBox"
        Me.SerialNumListBox.Size = New System.Drawing.Size(257, 196)
        Me.SerialNumListBox.TabIndex = 4
        '
        'QtyTextBox
        '
        Me.QtyTextBox.Location = New System.Drawing.Point(155, 32)
        Me.QtyTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.QtyTextBox.MaxLength = 3
        Me.QtyTextBox.Name = "QtyTextBox"
        Me.QtyTextBox.Size = New System.Drawing.Size(39, 22)
        Me.QtyTextBox.TabIndex = 2
        '
        'SerialNumTextBox
        '
        Me.SerialNumTextBox.Location = New System.Drawing.Point(16, 32)
        Me.SerialNumTextBox.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SerialNumTextBox.MaxLength = 12
        Me.SerialNumTextBox.Name = "SerialNumTextBox"
        Me.SerialNumTextBox.Size = New System.Drawing.Size(129, 22)
        Me.SerialNumTextBox.TabIndex = 1
        '
        'AddSerialNumberButton
        '
        Me.AddSerialNumberButton.Location = New System.Drawing.Point(203, 30)
        Me.AddSerialNumberButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AddSerialNumberButton.Name = "AddSerialNumberButton"
        Me.AddSerialNumberButton.Size = New System.Drawing.Size(72, 28)
        Me.AddSerialNumberButton.TabIndex = 3
        Me.AddSerialNumberButton.Text = "Add"
        Me.AddSerialNumberButton.UseVisualStyleBackColor = True
        '
        'ClearSerialNumbersButton
        '
        Me.ClearSerialNumbersButton.Location = New System.Drawing.Point(201, 268)
        Me.ClearSerialNumbersButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ClearSerialNumbersButton.Name = "ClearSerialNumbersButton"
        Me.ClearSerialNumbersButton.Size = New System.Drawing.Size(73, 28)
        Me.ClearSerialNumbersButton.TabIndex = 9
        Me.ClearSerialNumbersButton.Text = "Clear All"
        Me.ClearSerialNumbersButton.UseVisualStyleBackColor = True
        '
        'StartSerialLabel
        '
        Me.StartSerialLabel.AutoSize = True
        Me.StartSerialLabel.Location = New System.Drawing.Point(12, 12)
        Me.StartSerialLabel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.StartSerialLabel.Name = "StartSerialLabel"
        Me.StartSerialLabel.Size = New System.Drawing.Size(78, 17)
        Me.StartSerialLabel.TabIndex = 5
        Me.StartSerialLabel.Text = "Start Serial"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(151, 12)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Qty"
        '
        'PrintButton
        '
        Me.PrintButton.Location = New System.Drawing.Point(203, 306)
        Me.PrintButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PrintButton.Name = "PrintButton"
        Me.PrintButton.Size = New System.Drawing.Size(73, 28)
        Me.PrintButton.TabIndex = 5
        Me.PrintButton.Text = "Print"
        Me.PrintButton.UseVisualStyleBackColor = True
        '
        'ViewButton
        '
        Me.ViewButton.Location = New System.Drawing.Point(108, 306)
        Me.ViewButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ViewButton.Name = "ViewButton"
        Me.ViewButton.Size = New System.Drawing.Size(87, 28)
        Me.ViewButton.TabIndex = 6
        Me.ViewButton.Text = "View"
        Me.ViewButton.UseVisualStyleBackColor = True
        '
        'CloseMultUnitButton
        '
        Me.CloseMultUnitButton.Location = New System.Drawing.Point(15, 306)
        Me.CloseMultUnitButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CloseMultUnitButton.Name = "CloseMultUnitButton"
        Me.CloseMultUnitButton.Size = New System.Drawing.Size(85, 28)
        Me.CloseMultUnitButton.TabIndex = 7
        Me.CloseMultUnitButton.Text = "Cancel"
        Me.CloseMultUnitButton.UseVisualStyleBackColor = True
        '
        'RemoveSelectedSerialNumber
        '
        Me.RemoveSelectedSerialNumber.Location = New System.Drawing.Point(15, 268)
        Me.RemoveSelectedSerialNumber.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.RemoveSelectedSerialNumber.Name = "RemoveSelectedSerialNumber"
        Me.RemoveSelectedSerialNumber.Size = New System.Drawing.Size(75, 31)
        Me.RemoveSelectedSerialNumber.TabIndex = 8
        Me.RemoveSelectedSerialNumber.Text = "Remove"
        Me.RemoveSelectedSerialNumber.UseVisualStyleBackColor = True
        '
        'MultReportPrintDoc
        '
        Me.MultReportPrintDoc.DocumentName = "Module Report"
        '
        'PrintMultReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(291, 342)
        Me.Controls.Add(Me.RemoveSelectedSerialNumber)
        Me.Controls.Add(Me.CloseMultUnitButton)
        Me.Controls.Add(Me.ViewButton)
        Me.Controls.Add(Me.PrintButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.StartSerialLabel)
        Me.Controls.Add(Me.ClearSerialNumbersButton)
        Me.Controls.Add(Me.AddSerialNumberButton)
        Me.Controls.Add(Me.SerialNumTextBox)
        Me.Controls.Add(Me.QtyTextBox)
        Me.Controls.Add(Me.SerialNumListBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "PrintMultReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Multiple Reports"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SerialNumListBox As System.Windows.Forms.ListBox
    Friend WithEvents QtyTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SerialNumTextBox As System.Windows.Forms.TextBox
    Friend WithEvents AddSerialNumberButton As System.Windows.Forms.Button
    Friend WithEvents ClearSerialNumbersButton As System.Windows.Forms.Button
    Friend WithEvents StartSerialLabel As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PrintButton As System.Windows.Forms.Button
    Friend WithEvents ViewButton As System.Windows.Forms.Button
    Friend WithEvents CloseMultUnitButton As System.Windows.Forms.Button
    Friend WithEvents RemoveSelectedSerialNumber As System.Windows.Forms.Button
    Friend WithEvents MultReportPrintDoc As System.Drawing.Printing.PrintDocument
End Class
