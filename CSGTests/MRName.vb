Imports CSGCore

Public Class MRName

    Private Sub BConfirm_Click(sender As Object, e As EventArgs) Handles BConfirm.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If CBMRName.Text = "" Or CBMRSuffix.Text = "" Then
            MsgBox("Invalid input for Model Range")
        Else
            General.currentunit.ModelRangeName = CBMRName.Text.ToUpper
            General.currentunit.ModelRangeSuffix = CBMRSuffix.Text.ToUpper
            ConSysProps.BCSData.PerformClick()
            Close()
        End If
    End Sub

    Private Sub CBMRName_TextChanged(sender As Object, e As EventArgs) Handles CBMRName.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBMRName.Text)
        CBMRSuffix.Items.Clear()

        If CBMRName.Text = "GACV" Then
            CBMRSuffix.Items.AddRange({"FP", "AX", "RX", "CX", "AP", "CP"})
        Else
            CBMRSuffix.Items.AddRange({"FD", "OD", "AD", "RD", "CD"})
        End If
    End Sub

    Private Sub CBMRSuffix_TextChanged(sender As Object, e As EventArgs) Handles CBMRSuffix.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBMRSuffix.Text)
    End Sub

    Private Sub CBMRName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBMRName.SelectedIndexChanged

    End Sub
End Class