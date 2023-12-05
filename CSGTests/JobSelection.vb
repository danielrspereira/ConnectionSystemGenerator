Imports CSGCore

Public Class JobSelection
    Private Sub BOK_Click(sender As Object, e As EventArgs) Handles BOK.Click
        Try
            Dim rowno As Integer = CSGTable.SelectedCells(0).RowIndex
            Dim DGRow As DataGridViewRow = CSGTable.Rows(rowno)
            Dim uid As String = DGRow.Cells("Uid").Value

            General.currentjob = Database.GetJobInfo(uid)
            OrderProps.ContinueSaved()
            Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub JobSelection_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Shared Sub FillGrid(joblist As List(Of JobInfo))

        JobSelection.CSGTable.Rows.Clear()

        Try
            Dim tempjobs = From jlist In joblist Order By jlist.Uid Descending

            For i As Integer = 0 To tempjobs.Count - 1
                Dim row() As String = {tempjobs(i).Uid, tempjobs(i).OrderNumber, tempjobs(i).OrderPosition, tempjobs(i).ProjectNumber, tempjobs(i).Prio.ToString, tempjobs(i).Status.ToString, tempjobs(i).Users, tempjobs(i).RequestTime.ToShortDateString + " " + tempjobs(i).RequestTime.ToLongTimeString, tempjobs(i).PDMID}
                Dim rows() As Object = {row}
                Dim rowArray As String()
                For Each rowArray In rows
                    JobSelection.CSGTable.Rows.Add(rowArray)
                Next rowArray
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub
End Class