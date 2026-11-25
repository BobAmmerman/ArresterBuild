<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmData
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
        Dim Label1 As System.Windows.Forms.Label
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Batch, DiskNo, Category")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Batch, DiskNo, Category")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Batch, DiskNo, Category")
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("A1", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2, TreeNode3})
        Dim TreeNode5 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node18")
        Dim TreeNode6 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Batch1", New System.Windows.Forms.TreeNode() {TreeNode5})
        Dim TreeNode7 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Batch 2")
        Dim TreeNode8 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("B1", New System.Windows.Forms.TreeNode() {TreeNode6, TreeNode7})
        Dim TreeNode9 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("B2")
        Dim TreeNode10 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("C1")
        Dim TreeNode11 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("C2")
        Dim TreeNode12 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SN123456789", New System.Windows.Forms.TreeNode() {TreeNode4, TreeNode8, TreeNode9, TreeNode10, TreeNode11})
        Dim TreeNode13 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node4")
        Dim TreeNode14 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node6")
        Dim TreeNode15 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node7")
        Dim TreeNode16 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Node5", New System.Windows.Forms.TreeNode() {TreeNode14, TreeNode15})
        Dim TreeNode17 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SN123456000", New System.Windows.Forms.TreeNode() {TreeNode13, TreeNode16})
        Dim TreeNode18 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SN123456001")
        Dim TreeNode19 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("SN123456002")
        Me.tvData = New System.Windows.Forms.TreeView()
        Me.tbWO = New System.Windows.Forms.TextBox()
        Me.lblCompleted = New System.Windows.Forms.Label()
        Me.lblCatNum = New System.Windows.Forms.Label()
        Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(14, 15)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(51, 13)
        Label1.TabIndex = 2
        Label1.Text = "WO Filter"
        '
        'tvData
        '
        Me.tvData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvData.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tvData.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvData.Location = New System.Drawing.Point(12, 69)
        Me.tvData.Name = "tvData"
        TreeNode1.Name = "Node15"
        TreeNode1.Text = "Batch, DiskNo, Category"
        TreeNode2.Name = "Node16"
        TreeNode2.Text = "Batch, DiskNo, Category"
        TreeNode3.Name = "Node17"
        TreeNode3.Text = "Batch, DiskNo, Category"
        TreeNode4.Name = "Node8"
        TreeNode4.Text = "A1"
        TreeNode5.Name = "Node18"
        TreeNode5.Text = "Node18"
        TreeNode6.Name = "Node10"
        TreeNode6.Text = "Batch1"
        TreeNode7.Name = "Node11"
        TreeNode7.Text = "Batch 2"
        TreeNode8.Name = "Node9"
        TreeNode8.Text = "B1"
        TreeNode9.Name = "Node12"
        TreeNode9.Text = "B2"
        TreeNode10.Name = "Node13"
        TreeNode10.Text = "C1"
        TreeNode11.Name = "Node14"
        TreeNode11.Text = "C2"
        TreeNode12.Name = "Node0"
        TreeNode12.Text = "SN123456789"
        TreeNode13.Name = "Node4"
        TreeNode13.Text = "Node4"
        TreeNode14.Name = "Node6"
        TreeNode14.Text = "Node6"
        TreeNode15.Name = "Node7"
        TreeNode15.Text = "Node7"
        TreeNode16.Name = "Node5"
        TreeNode16.Text = "Node5"
        TreeNode17.Name = "Node1"
        TreeNode17.Text = "SN123456000"
        TreeNode18.Name = "Node2"
        TreeNode18.Text = "SN123456001"
        TreeNode19.Name = "Node3"
        TreeNode19.Text = "SN123456002"
        Me.tvData.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode12, TreeNode17, TreeNode18, TreeNode19})
        Me.tvData.Size = New System.Drawing.Size(542, 380)
        Me.tvData.TabIndex = 0
        '
        'tbWO
        '
        Me.tbWO.Location = New System.Drawing.Point(79, 12)
        Me.tbWO.Name = "tbWO"
        Me.tbWO.Size = New System.Drawing.Size(100, 20)
        Me.tbWO.TabIndex = 1
        '
        'lblCompleted
        '
        Me.lblCompleted.AutoSize = True
        Me.lblCompleted.Location = New System.Drawing.Point(212, 12)
        Me.lblCompleted.Name = "lblCompleted"
        Me.lblCompleted.Size = New System.Drawing.Size(123, 13)
        Me.lblCompleted.TabIndex = 3
        Me.lblCompleted.Text = "15,15,14 of 15 Complete"
        '
        'lblCatNum
        '
        Me.lblCatNum.AutoSize = True
        Me.lblCatNum.Location = New System.Drawing.Point(76, 41)
        Me.lblCatNum.Name = "lblCatNum"
        Me.lblCatNum.Size = New System.Drawing.Size(83, 13)
        Me.lblCatNum.TabIndex = 4
        Me.lblCatNum.Text = "Catalog Number"
        '
        'frmData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(566, 461)
        Me.Controls.Add(Me.lblCatNum)
        Me.Controls.Add(Me.lblCompleted)
        Me.Controls.Add(Label1)
        Me.Controls.Add(Me.tbWO)
        Me.Controls.Add(Me.tvData)
        Me.Name = "frmData"
        Me.Text = "View Disk Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tvData As System.Windows.Forms.TreeView
    Friend WithEvents tbWO As System.Windows.Forms.TextBox
    Friend WithEvents lblCompleted As System.Windows.Forms.Label
    Friend WithEvents lblCatNum As System.Windows.Forms.Label
End Class
