Imports CSGCore

Public Class DBInfo

    Private Sub DBInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CurrentJobData()
    End Sub

    Private Sub BRefresh_Click(sender As Object, e As EventArgs) Handles BRefresh.Click
        CurrentJobData()
    End Sub

    Shared Sub CurrentJobData()
        Dim names() As String = {"OrderNumber", "OrderPosition", "Prio"}

        Dim values() As String = {General.currentjob.OrderNumber, General.currentjob.OrderPosition, General.currentjob.Prio}
        Dim joblist As List(Of JobInfo) = Database.SearchJobs(names, values)

        DBInfo.CSGTable.Rows.Clear()

        Try
            Dim tempjobs = From jlist In joblist Order By jlist.Uid Descending

            For i As Integer = 0 To tempjobs.Count - 1
                Dim row() As String = {tempjobs(i).Uid, tempjobs(i).OrderNumber, tempjobs(i).OrderPosition, tempjobs(i).ProjectNumber, tempjobs(i).Prio.ToString, tempjobs(i).Status.ToString, tempjobs(i).Users, tempjobs(i).RequestTime.ToShortDateString + " " + tempjobs(i).RequestTime.ToLongTimeString}
                Dim rows() As Object = {row}
                Dim rowArray As String()
                For Each rowArray In rows
                    DBInfo.CSGTable.Rows.Add(rowArray)
                Next rowArray
            Next

            DBInfo.UpdateLabel(joblist.Count, "found for this order")

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Private Sub BShow_Click(sender As Object, e As EventArgs) Handles BShow.Click
        Dim joblist As List(Of JobInfo) = Database.SearchJobs({"Status", "Prio"}, {"0", "2"})
        Dim joblist2 As List(Of JobInfo) = Database.SearchJobs({"Status", "Prio"}, {"0", "3"})

        If joblist2.Count > 0 Then
            joblist.AddRange(joblist2.ToArray)
        End If

        CSGTable.Rows.Clear()

        Try
            Dim tempjobs = From jlist In joblist Order By jlist.Uid Descending

            For i As Integer = 0 To tempjobs.Count - 1
                Dim row() As String = {tempjobs(i).Uid, tempjobs(i).OrderNumber, tempjobs(i).OrderPosition, tempjobs(i).ProjectNumber, tempjobs(i).Prio.ToString, tempjobs(i).Status.ToString, tempjobs(i).Users, tempjobs(i).RequestTime.ToShortDateString + " " + tempjobs(i).RequestTime.ToLongTimeString}
                Dim rows() As Object = {row}
                Dim rowArray As String()
                For Each rowArray In rows
                    CSGTable.Rows.Add(rowArray)
                Next rowArray
            Next

            UpdateLabel(joblist.Count, "waiting in queue")

        Catch ex As Exception
            CSGCore.General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Private Sub UpdateLabel(count As Integer, partialtext As String)
        If count = 1 Then
            Linfo.Text = count.ToString + " entry " + partialtext
        Else
            Linfo.Text = count.ToString + " entries " + partialtext
        End If
    End Sub

    Private Sub DBInfo_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        Dim height As Integer = Size.Height
        Dim newhgrid As Integer = height - 182

        CSGTable.Height = newhgrid
    End Sub

    Private Sub BCancel_Click(sender As Object, e As EventArgs) Handles BCancel.Click
        Try
            BCancel.BackColor = BCancel.BackColor
            Dim rowno As Integer = CSGTable.SelectedCells(0).RowIndex
            Dim DGRow As DataGridViewRow = CSGTable.Rows(rowno)
            Dim uid As String = DGRow.Cells("CUid").Value
            Dim username As String = DGRow.Cells("CUsers").Value
            Dim status As Integer

            If username = General.username Then
                status = Database.GetValue("CSG.batch_csg", "Status", "uid", uid, "int")
                If status = 0 Then
                    If Database.CancelJob("CSG.batch_csg", uid) = 1 Then
                        BCancel.BackColor = Color.Green
                    Else
                        BCancel.BackColor = Color.Red
                    End If
                Else
                    MsgBox("No permission, job has been started already")
                End If
            Else
                MsgBox("No permission, job was requested by a different user")
            End If

        Catch ex As Exception
            MsgBox("Error cancelling the job")
        End Try
    End Sub
End Class