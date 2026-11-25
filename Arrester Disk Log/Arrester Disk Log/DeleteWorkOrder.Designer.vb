<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DeleteWorkOrder
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DeleteWorkOrder))
        Me.WorkOrderToDeleteLabel = New System.Windows.Forms.Label()
        Me.WorkOrderToDeleteTextBox = New System.Windows.Forms.TextBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.CancelDelWOButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'WorkOrderToDeleteLabel
        '
        Me.WorkOrderToDeleteLabel.AutoSize = True
        Me.WorkOrderToDeleteLabel.Location = New System.Drawing.Point(12, 13)
        Me.WorkOrderToDeleteLabel.Name = "WorkOrderToDeleteLabel"
        Me.WorkOrderToDeleteLabel.Size = New System.Drawing.Size(105, 13)
        Me.WorkOrderToDeleteLabel.TabIndex = 0
        Me.WorkOrderToDeleteLabel.Text = "Work Order Number:"
        '
        'WorkOrderToDeleteTextBox
        '
        Me.WorkOrderToDeleteTextBox.Location = New System.Drawing.Point(126, 10)
        Me.WorkOrderToDeleteTextBox.MaxLength = 8
        Me.WorkOrderToDeleteTextBox.Name = "WorkOrderToDeleteTextBox"
        Me.WorkOrderToDeleteTextBox.Size = New System.Drawing.Size(100, 20)
        Me.WorkOrderToDeleteTextBox.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(148, 45)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 2
        Me.OKButton.Text = "Okay"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'CancelDelWOButton
        '
        Me.CancelDelWOButton.Location = New System.Drawing.Point(15, 45)
        Me.CancelDelWOButton.Name = "CancelDelWOButton"
        Me.CancelDelWOButton.Size = New System.Drawing.Size(75, 23)
        Me.CancelDelWOButton.TabIndex = 3
        Me.CancelDelWOButton.Text = "Cancel"
        Me.CancelDelWOButton.UseVisualStyleBackColor = True
        '
        'DeleteWorkOrder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(238, 81)
        Me.Controls.Add(Me.CancelDelWOButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.WorkOrderToDeleteTextBox)
        Me.Controls.Add(Me.WorkOrderToDeleteLabel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DeleteWorkOrder"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delete Work Order"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents WorkOrderToDeleteLabel As System.Windows.Forms.Label
    Friend WithEvents WorkOrderToDeleteTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents CancelDelWOButton As System.Windows.Forms.Button
End Class
