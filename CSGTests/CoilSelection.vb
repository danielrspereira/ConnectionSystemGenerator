Imports CSGCore
Public Class CoilSelection
    Public caller As String = ""

    Private Sub CoilSelection_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        General.CreateActionLogEntry(Me.Name, sender.name, "loading")
        'create TreeView
        CreateTreeView(General.currentjob.OrderNumber, General.currentjob.OrderPosition)
    End Sub

    Shared Sub CreateTreeView(orderno As String, orderpos As String)
        Dim itemlist As List(Of BOMItem)
        Dim unititem As BOMItem

        Try
            itemlist = Order.GetBOMItemByQuery(General.currentunit.BOMList, "parent", "0")
            unititem = itemlist(0)

            CoilSelection.UnitTree.Nodes.Clear()
            CopyProps.AddTreeNode(CoilSelection.UnitTree, unititem.Item, unititem.ItemDescription + " - " + unititem.Item, "")
            AddChildren(unititem, "")

            CoilSelection.UnitTree.CollapseAll()
        Catch ex As Exception

        End Try
    End Sub

    Shared Sub AddChildren(parentitem As BOMItem, fullpath As String)
        Dim clist As List(Of BOMItem) = Order.GetBOMItemByQuery(General.currentunit.BOMList, "parent", parentitem.ItemNumber)
        Dim path As String

        Try
            If fullpath = "" Then
                path = parentitem.Item + "\"
            Else
                path = fullpath
            End If

            For Each child In clist
                CopyProps.AddTreeNode(CoilSelection.UnitTree, child.Item, child.ItemDescription + " - " + child.Item, path)
                AddChildren(child, path + child.Item + "\")
            Next
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub BConfirm_Click(sender As Object, e As EventArgs) Handles BConfirm.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Dim coilnode As TreeNode

        Try
            coilnode = UnitTree.SelectedNode
            If coilnode.Name <> "" Then
                Dim slist As List(Of BOMItem) = Order.GetBOMItemByQuery(General.currentunit.BOMList, "item", coilnode.Name)
                If slist.Count > 0 Then
                    If caller = "copy" Then
                        General.currentunit.Coillist(0).BOMItem = slist(0)
                        CopyProps.FindERPCodes()
                    Else
                        UnitProps.selectedcoil.BOMItem = slist(0)
                    End If
                End If
                General.CreateActionLogEntry(Me.Name, sender.name, "closed")
                Close()
            Else
                MsgBox("Invalid selection for coil item")
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class