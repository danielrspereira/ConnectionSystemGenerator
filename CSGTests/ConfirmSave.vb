Imports CSGCore

Public Class ConfirmSave
    Public Labeltext As String

    Private Sub BYes_Click(sender As Object, e As EventArgs) Handles BYes.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        UnitProps.saveok = True
        UnitProps.BSave.PerformClick()
        Close()
    End Sub

    Private Sub BNo_Click(sender As Object, e As EventArgs) Handles BNo.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        UnitProps.saveok = False
        Close()
    End Sub
End Class