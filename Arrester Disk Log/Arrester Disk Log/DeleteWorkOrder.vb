Option Strict On
Public Class DeleteWorkOrder

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Dim deletedWO As New WorkOrder
        deletedWO.WorkOrderNum = WorkOrderToDeleteTextBox.Text
        deletedWO.DeleteWorkOrder()
        Me.Close()
    End Sub

    Private Sub CancelDelWOButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelDelWOButton.Click
        Me.Close()
    End Sub

    Private Sub WorkOrderToDeleteTextBox_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WorkOrderToDeleteTextBox.KeyPress
        If (e.KeyChar = vbCr) Then
            OKButton_Click(sender, e)
        End If
    End Sub
End Class