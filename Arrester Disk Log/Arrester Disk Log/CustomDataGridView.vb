Option Strict On
Public Class CustomDataGridView
    Inherits DataGridView

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean


        If keyData = Keys.Enter Then

            MyBase.ProcessTabKey(Keys.Tab)
            'Still do the default processing for enter so it will move down to the next row

            Return False 'say we have *not* handled the key

        Else

            Return MyBase.ProcessCmdKey(msg, keyData)

        End If

    End Function



End Class

