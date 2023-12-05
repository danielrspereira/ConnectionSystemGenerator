Imports System.ComponentModel
Imports System.IO
Imports System.Windows
Imports CSGCore

Public Class CopyProps
    Public partlist, asmlist As New List(Of String)
    Public oldunit As New UnitData
    Public agpnew As Boolean = False
    Public Delegate Sub UpdateTree(objtree As TreeView, name As String, text As String, master As String)
    Public Delegate Sub UpdateLog(content As String)
    Public Delegate Sub AbleButton(active As Boolean)
    Public Delegate Sub OleFilter(active As Boolean)

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        BWAnalyze.WorkerReportsProgress = True
        BWAnalyze.WorkerSupportsCancellation = True
        BWRename.WorkerReportsProgress = True
        BWRename.WorkerSupportsCancellation = True

    End Sub

    Private Sub CopyProps_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With OrderProps
            TBOrder1.Text = .TBOrder.Text
            TBPos1.Text = .TBPos.Text
            TBNr1.Text = .TBNr.Text
        End With
        BAnalyze.Enabled = General.isProdavit
        BRename.Enabled = General.isProdavit
        BImport.Visible = Not General.isProdavit
        General.currentunit.CopiedFrom = New JobInfo
    End Sub

    Private Sub BFind_Click(sender As Object, e As EventArgs) Handles BFind.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If TBOrder2.Text <> "" And TBPos2.Text <> "" Then
                'search for existing job in batch table
                Dim joblist As List(Of JobInfo) = Database.SearchJobs({"OrderNumber", "OrderPosition", "Status"}, {TBOrder2.Text, TBPos2.Text, "100"})
                If joblist.Count = 1 Then
                    Dim pdmid As String = Database.GetValue("batch_csg", "PDMID", "uid", joblist(0).Uid)
                    TBMaster.Text = pdmid
                    With General.currentunit.CopiedFrom
                        .OrderNumber = TBOrder2.Text
                        .OrderPosition = TBPos2.Text
                        .ProjectNumber = TBNr2.Text
                        .PDMID = TBMaster.Text
                    End With
                ElseIf joblist.Count = 0 Then
                    MsgBox("No finished job found!")
                Else
                    MsgBox("More than 1 finished job found, result is ambiguous." + vbNewLine + "Please enter PDM number of Masterassembly manually!")
                End If
            Else
                MsgBox("Missing Order information!")
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub BLoad_Click(sender As Object, e As EventArgs) Handles BLoad.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            OleMessageFilter.Register()

            If TBMaster.Text <> "" Then
                If General.currentunit.CopiedFrom.Uid = 0 OrElse General.currentunit.CopiedFrom.PDMID <> TBMaster.Text Then
                    With General.currentunit.CopiedFrom
                        .OrderNumber = TBOrder2.Text
                        .OrderPosition = TBPos2.Text
                        .ProjectNumber = TBNr2.Text
                        .PDMID = TBMaster.Text
                    End With
                End If
                Dim masterid As String = TBMaster.Text.Replace(" ", "")
                WSM.fullpartids.Clear()
                WSM.CheckoutPart(masterid, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile, "asm")
                If General.WaitForFile(General.currentjob.Workspace, masterid, "asm", 300) Then
                    'rename to masterassembly
                    File.Move(General.GetFullFilename(General.currentjob.Workspace, masterid, "asm").Replace(".asm", ".cfg"), Path.Combine(General.currentjob.Workspace, "Masterassembly.cfg"))
                    File.Move(General.GetFullFilename(General.currentjob.Workspace, masterid, "asm"), Path.Combine(General.currentjob.Workspace, "Masterassembly.asm"))
                    RTBLog.AppendText("Done loading reference asssembly" + vbNewLine)
                End If

            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        Finally
            OleMessageFilter.Unregister()
        End Try

    End Sub

    Private Sub BAnalyze_Click(sender As Object, e As EventArgs) Handles BAnalyze.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If Not BWAnalyze.IsBusy Then
            Try
                General.seapp.Documents.Open(Path.Combine(General.currentjob.Workspace, "Masterassembly.asm"))
                General.seapp.DoIdle()
                ResetProgress()
                BWAnalyze.RunWorkerAsync()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub BRename_Click(sender As Object, e As EventArgs) Handles BRename.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If Not BWAnalyze.IsBusy And oldunit.Coillist.Count > 0 And Not BWRename.IsBusy Then
            BWRename.RunWorkerAsync()
        End If
    End Sub

    Private Sub BSave_Click(sender As Object, e As EventArgs) Handles BSave.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Dim coilnames As New List(Of String)
        Try
            If Not BWAnalyze.IsBusy And Not BWRename.IsBusy Then
                Dim masterasm As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(Path.Combine(General.currentjob.Workspace, "Masterassembly.asm"))
                General.seapp.DoIdle()

                'overwrite masterasm file properties
                SEPart.GetSetCustomProp(masterasm, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                SEPart.GetSetCustomProp(masterasm, "Position", General.currentjob.OrderPosition, "write")
                SEPart.GetSetCustomProp(masterasm, "Order_Projekt", General.currentjob.ProjectNumber, "write")

                'save process 
                If WSM.SaveDFT() Then
                    'wait until finished
                    WSM.WaitforWSMDialog()

                    masterasm = General.seapp.ActiveDocument
                    General.seapp.DoIdle()

                    For Each objocc As SolidEdgeAssembly.Occurrence In masterasm.Occurrences
                        If objocc.Name.Contains(".asm") Then
                            coilnames.Add(objocc.OccurrenceFileName)
                        End If
                    Next

                    Dim dirinfo As New DirectoryInfo(General.currentjob.Workspace)

                    'update all drawing files in the workspace
                    For Each f As FileInfo In dirinfo.GetFiles()
                        If f.Extension = ".dft" Then
                            SEDrawing.UpdateDrawingAfterSave(f.FullName)
                        End If
                    Next

                    For Each c In coilnames
                        'open coilasm
                        General.seapp.Documents.Open(c)
                        General.seapp.DoIdle()

                        'save the coil and consys
                        WSM.SaveDFT()

                        WSM.WaitforWSMDialog()

                        RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Saved " + General.GetShortName(c))

                        'close the coilasm
                        General.seapp.Documents.CloseDocument(c, SaveChanges:=True, DoIdle:=True)
                    Next

                    General.currentunit.UnitFile.Fullfilename = masterasm.FullName
                    General.currentjob.PDMID = General.GetShortName(masterasm.Name)

                    Database.UpdateJob(General.currentjob, {"Status", "FinishedTime", "ModelRange", "PDMID", "Saved"}, {"100", Database.ConvertDatetoStr(Date.UtcNow), General.currentunit.ModelRangeName, masterasm.Name.Substring(0, 10), "true"})

                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Copy process successfully finished")
                    General.seapp.Documents.CloseDocument(masterasm.FullName, SaveChanges:=False, DoIdle:=True)
                    BSave.Enabled = False
                Else
                    MsgBox("Error in the saving process")
                End If
            End If
        Catch ex As Exception
            RTBLog.AppendText(ex.ToString)
        End Try
    End Sub

    Private Sub BWAnalyze_DoWork(sender As Object, e As DoWorkEventArgs) Handles BWAnalyze.DoWork
        Invoke(New OleFilter(AddressOf ActivateMsgFilter), True)
        Try
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Scanning Workspace...")
            'scan workspace for all .par and .asm files
            ScanWorkspaceFiles(General.currentjob.Workspace)
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Finished scan...")


            Dim masterdoc As SolidEdgeAssembly.AssemblyDocument = General.seapp.ActiveDocument

            'create a treeview of structure
            TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Master", "Masterassembly - " + masterdoc.Name, "")

            'loop
            For i As Integer = 0 To masterdoc.Occurrences.Count - 1
                Dim occ1 As SolidEdgeAssembly.Occurrence = masterdoc.Occurrences(i)
                If occ1.Name.Contains(".asm") Then
                    TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Coil" + (i + 1).ToString, "Coil - " + occ1.Name, "Master\")

                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Scanning Coil" + (i + 1).ToString + "...")
                    oldunit.Coillist.Add(CreateCoilItem(occ1, i + 1))

                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Finished Coil" + (i + 1).ToString + "...")
                End If
            Next

            'close masterasm
            General.seapp.Documents.Close()
            General.seapp.DoIdle()

            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Finished analyzing")
            Database.UpdateJob(General.currentjob, {"Status"}, {"10"})
        Catch ex As Exception
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), ex.ToString)
            Debug.Print(ex.ToString)
        Finally
            Invoke(New OleFilter(AddressOf ActivateMsgFilter), False)
        End Try
    End Sub

    Private Sub BWAnalyze_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BWAnalyze.RunWorkerCompleted
        Try
            With General.currentunit
                If .Coillist(0).BOMItem.Item = "" Then
                    CoilSelection.Show()
                    CoilSelection.Activate()
                    CoilSelection.caller = "copy"
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BWRename_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BWRename.DoWork
        Dim consysdic, coildic, partdic As New Dictionary(Of String, String)
        Invoke(New OleFilter(AddressOf ActivateMsgFilter), True)
        Try
            'use the occlists to rename files
            For Each coil In oldunit.Coillist
                For Each consys In coil.ConSyss
                    'rename the occurences of the consys assembly
                    For i As Integer = 0 To consys.Occlist.Count - 1
                        If consys.Occlist(i).Configref = "" Then
                            If (agpnew And consys.Occlist(i).BOMAGP = "1") Or consys.Occlist(i).BOMAGP <> "1" Then
                                Dim oldname As String = Path.Combine(General.currentjob.Workspace, consys.Occlist(i).Occname)
                                Dim newname As String = Path.Combine(General.currentjob.Workspace, consys.Occlist(i).Occindex.ToString + "_" + consys.Occlist(i).Occname)
                                If File.Exists(oldname) Then
                                    File.Move(oldname, newname)
                                    partdic.Add(oldname, newname)
                                End If
                            End If
                        End If
                    Next
                    'rename the assembly
                    Dim newconsysname As String = Path.Combine(General.currentjob.Workspace, coil.Number.ToString + "_" + consys.Circnumber.ToString + consys.ConSysFile.Shortname)
                    File.Move(consys.ConSysFile.Fullfilename, newconsysname)
                    File.Move(consys.ConSysFile.Fullfilename.Replace(".asm", ".cfg"), newconsysname.Replace(".asm", ".cfg"))
                    consysdic.Add(consys.ConSysFile.Fullfilename, newconsysname)
                Next

                For i As Integer = 0 To coil.Occlist.Count - 1
                    If coil.Occlist(i).Configref = "" Then
                        Dim oldname As String = Path.Combine(General.currentjob.Workspace, coil.Occlist(i).Occname)
                        Dim newname As String = Path.Combine(General.currentjob.Workspace, coil.Occlist(i).Occindex.ToString + "_" + coil.Occlist(i).Occname)

                        If oldname.Contains(".asm") Then
                            'already renamed, override oldname
                            For Each en In consysdic
                                If en.Key = oldname Then
                                    oldname = en.Value
                                    newname = en.Value
                                End If
                            Next
                        End If

                        If File.Exists(oldname) Then
                            If oldname.Contains(".asm") Then
                                Dim oldconsysdftname As String = General.GetFullFilename(General.currentjob.Workspace, General.GetShortName(oldname).Substring(3, 10), "dft")
                                If oldconsysdftname <> "" Then
                                    File.Move(oldconsysdftname, newname.Replace(".asm", ".dft"))
                                Else
                                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Missing dft for " + General.GetShortName(oldname))
                                End If
                            ElseIf (agpnew And coil.Occlist(i).BOMAGP = "1") Or coil.Occlist(i).BOMAGP <> "1" Then
                                File.Move(oldname, newname)
                                partdic.Add(oldname, newname)
                            End If
                        End If
                    End If
                Next
                'rename the assembly
                Dim newcoilname As String = Path.Combine(General.currentjob.Workspace, coil.Number.ToString + coil.CoilFile.Shortname)
                File.Move(coil.CoilFile.Fullfilename, newcoilname)
                Dim oldcoildftname As String = General.GetFullFilename(General.currentjob.Workspace, coil.CoilFile.Shortname.Substring(0, 10), "dft")
                If oldcoildftname <> "" Then
                    File.Move(oldcoildftname, newcoilname.Replace(".asm", ".dft"))
                    File.Move(coil.CoilFile.Fullfilename.Replace(".asm", ".cfg"), newcoilname.Replace(".asm", ".cfg"))
                    coildic.Add(coil.CoilFile.Fullfilename, newcoilname)
                End If
            Next

            For Each entry In consysdic
                Dim consysasmdoc As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(entry.Value)
                General.seapp.DoIdle()
                'use entry.key to search in all coillist for the correct consys.occlist
                Dim consysocclist As List(Of PartData) = GetConsysOcclist(entry.Key)

                For Each en In partdic
                    Dim partocc As SolidEdgeAssembly.Occurrence = SEAsm.CheckOccExists(consysasmdoc.Occurrences, General.GetShortName(en.Key) + ":1")
                    If partocc IsNot Nothing Then
                        partocc.Replace(en.Value, True)
                        General.seapp.DoIdle()

                        SEPart.GetSetCustomProp(partocc.PartDocument, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                        SEPart.GetSetCustomProp(partocc.PartDocument, "Position", General.currentjob.OrderPosition, "write")
                        SEPart.GetSetCustomProp(partocc.PartDocument, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                    End If
                Next

                'show all components
                General.seapp.StartCommand(33071)

                General.seapp.Documents.CloseDocument(entry.Value, SaveChanges:=True, DoIdle:=True)
            Next

            Dim coilcount As Integer = 0
            For Each entry In coildic
                Dim coilasmdoc As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(entry.Value)
                General.seapp.DoIdle()

                'get the occlist of the coil
                Dim coilocclist As List(Of PartData) = GetCoilOcclist(entry.Key)

                For Each en In partdic
                    Dim partocc As SolidEdgeAssembly.Occurrence = SEAsm.CheckOccExists(coilasmdoc.Occurrences, General.GetShortName(en.Key) + ":1")
                    If partocc IsNot Nothing Then
                        partocc.Replace(en.Value, True)
                        General.seapp.DoIdle()

                        SEPart.GetSetCustomProp(partocc.PartDocument, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                        SEPart.GetSetCustomProp(partocc.PartDocument, "Position", General.currentjob.OrderPosition, "write")
                        SEPart.GetSetCustomProp(partocc.PartDocument, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                    End If
                Next

                Dim consyscount As Integer = 0
                For Each en In consysdic
                    Dim asmocc As SolidEdgeAssembly.Occurrence = SEAsm.CheckOccExists(coilasmdoc.Occurrences, General.GetShortName(en.Key) + ":1")
                    If asmocc IsNot Nothing Then
                        asmocc.Replace(en.Value, True)
                        General.seapp.DoIdle()
                        RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Replaced occurence " + en.Key + " with " + en.Value)

                        SEPart.GetSetCustomProp(asmocc.PartDocument, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                        SEPart.GetSetCustomProp(asmocc.PartDocument, "Position", General.currentjob.OrderPosition, "write")
                        SEPart.GetSetCustomProp(asmocc.PartDocument, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                        SEPart.GetSetCustomProp(asmocc.PartDocument, "GUE_Item", General.currentunit.Coillist(coilcount).ConSyss(consyscount).BOMItem.Item, "write")
                        RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Changes file properties of " + en.Key)
                    End If

                    'get consys drawing and change modellink
                    Dim consysdftname As String = General.GetFullFilename(General.currentjob.Workspace, General.GetShortName(en.Value).Substring(3, 10), "dft")
                    If consysdftname <> "" Then
                        General.currentunit.Coillist(coilcount).ConSyss(consyscount).ConSysFile.Shortname = General.GetShortName(en.Value)
                        Dim consysdftdoc As SolidEdgeDraft.DraftDocument = General.seapp.Documents.Open(consysdftname)
                        General.seapp.DoIdle()
                        SEDrawing.ChangeModellink(consysdftdoc, General.currentjob.Workspace, General.GetShortName(asmocc.PartFileName))
                        SEPart.GetSetCustomProp(consysdftdoc, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                        SEPart.GetSetCustomProp(consysdftdoc, "Position", General.currentjob.OrderPosition, "write")
                        SEPart.GetSetCustomProp(consysdftdoc, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                        SEPart.GetSetCustomProp(consysdftdoc, "GUE_Item", General.currentunit.Coillist(coilcount).ConSyss(consyscount).BOMItem.Item, "write")
                        General.seapp.Documents.CloseDocument(consysdftdoc.FullName, SaveChanges:=True, DoIdle:=True)
                    End If
                    consyscount += 1
                Next

                General.seapp.StartCommand(33071)

                SEPart.GetSetCustomProp(coilasmdoc, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                SEPart.GetSetCustomProp(coilasmdoc, "Position", General.currentjob.OrderPosition, "write")
                SEPart.GetSetCustomProp(coilasmdoc, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                SEPart.GetSetCustomProp(coilasmdoc, "GUE_Item", General.currentunit.Coillist(coilcount).BOMItem.Item, "write")
                General.currentunit.Coillist(coilcount).CoilFile.Shortname = General.GetShortName(entry.Value)
                Dim coildftname As String = General.GetFullFilename(General.currentjob.Workspace, General.GetShortName(entry.Value).Substring(1, 10), "dft")
                If coildftname <> "" Then
                    Dim coildftdoc As SolidEdgeDraft.DraftDocument = General.seapp.Documents.Open(coildftname)
                    General.seapp.DoIdle()
                    SEDrawing.ChangeModellink(coildftdoc, General.currentjob.Workspace, General.GetShortName(entry.Value))
                    SEPart.GetSetCustomProp(coildftdoc, "Auftragsnummer", General.currentjob.OrderNumber, "write")
                    SEPart.GetSetCustomProp(coildftdoc, "Position", General.currentjob.OrderPosition, "write")
                    SEPart.GetSetCustomProp(coildftdoc, "Order_Projekt", General.currentjob.ProjectNumber, "write")
                    SEPart.GetSetCustomProp(coildftdoc, "GUE_Item", General.currentunit.Coillist(coilcount).BOMItem.Item, "write")
                    General.seapp.Documents.CloseDocument(coildftdoc.FullName, SaveChanges:=True, DoIdle:=True)
                End If
                coilcount += 1

                General.seapp.Documents.CloseDocument(entry.Value, SaveChanges:=True, DoIdle:=True)
            Next

            'replace the coils in the master assembly
            Dim masterasm As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(Path.Combine(General.currentjob.Workspace, "Masterassembly.asm"))
            General.seapp.DoIdle()

            coilcount = 0
            For Each occ As SolidEdgeAssembly.Occurrence In masterasm.Occurrences
                If occ.Name.Contains(".asm") Then
                    Debug.Print(occ.Name)
                    Dim newcoilfile As String = General.GetFullFilename(General.currentjob.Workspace, General.currentunit.Coillist(coilcount).CoilFile.Shortname, "asm")
                    Debug.Print(newcoilfile)
                    occ.Replace(newcoilfile, True)
                    General.seapp.DoIdle()
                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Replaced occurence " + occ.Name + " with " + newcoilfile)
                    coilcount += 1
                End If
            Next

            General.seapp.StartCommand(33071)

            General.seapp.Documents.CloseDocument(masterasm.FullName, SaveChanges:=True, DoIdle:=True)

            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Finished renaming files and replacing occurences")
            Database.UpdateJob(General.currentjob, {"Status"}, {"20"})
            BSave.Invoke(New AbleButton(AddressOf EnableSaving), True)
        Catch ex As Exception
            Debug.Print(ex.ToString)
            Database.UpdateJob(General.currentjob, {"Status"}, {"-100"})
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Error renaming files:")
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), ex.ToString)
            BSave.Invoke(New AbleButton(AddressOf EnableSaving), False)
        Finally
            Invoke(New OleFilter(AddressOf ActivateMsgFilter), False)
        End Try
    End Sub

    Private Sub ScanWorkspaceFiles(workspace As String)

        For Each f As String In Directory.GetFiles(workspace)
            If f.Contains(".par") And Not f.Contains(".par.") Then
                partlist.Add(f.Substring(f.LastIndexOf("\") + 1))
            ElseIf f.Contains(".asm") And Not f.Contains(".asm.") Then
                asmlist.Add(f.Substring(f.LastIndexOf("\") + 1))
            End If
        Next

    End Sub

    Shared Sub AddTreeNode(objtree As TreeView, name As String, displaytext As String, parents As String)
        Try
            If parents.Contains("\") Then
                For Each tnode As TreeNode In objtree.Nodes
                    If tnode.Name = parents.Substring(0, parents.LastIndexOf("\")) Then
                        tnode.Nodes.Add(New TreeNode With {.Name = name, .Text = displaytext})
                    Else
                        CheckChildNodes(tnode, name, displaytext, parents.Substring(parents.IndexOf("\") + 1))
                    End If
                Next
            Else
                objtree.Nodes.Add(New TreeNode With {.Name = name, .Text = displaytext})
            End If

            objtree.ExpandAll()
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Shared Sub CheckChildNodes(parentnode As TreeNode, name As String, displaytext As String, parents As String)
        Try
            If parents.Contains("\") Then
                For Each tnode As TreeNode In parentnode.Nodes
                    If tnode.Name = parents.Substring(0, parents.LastIndexOf("\")) Then
                        tnode.Nodes.Add(New TreeNode With {.Name = name, .Text = displaytext})
                    Else
                        CheckChildNodes(tnode, name, displaytext, parents.Substring(parents.IndexOf("\") + 1))
                    End If
                Next
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub UpdateRTB(message As String)
        RTBLog.AppendText(message)
        RTBLog.AppendText(vbNewLine)
    End Sub

    Private Sub ActivateMsgFilter(active As Boolean)
        If active Then
            OleMessageFilter.Register()
        Else
            OleMessageFilter.Unregister()
        End If
    End Sub

    Private Sub EnableSaving(isactive As Boolean)
        BSave.Enabled = isactive
        If Not isactive Then

        End If
    End Sub

    Private Sub RTBLog_TextChanged(sender As Object, e As EventArgs) Handles RTBLog.TextChanged
        RTBLog.ScrollToCaret()
    End Sub

    Shared Sub FindERPCodes()
        With General.currentunit
            If .BOMList.Count > 0 Then
                If .Coillist(0).BOMItem.Item <> "" Then
                    For i As Integer = 0 To .Coillist(0).ConSyss.Count - 1
                        .Coillist(0).ConSyss(i).BOMItem = Order.GetConsysBOMItem(.Coillist(0).ConSyss(i), .Coillist(0), .Coillist(0).Circuits(i).CircuitType)
                    Next
                End If
            End If
        End With
    End Sub

    Private Function CreateCoilItem(coilocc As SolidEdgeAssembly.Occurrence, no As Integer) As CoilData
        Dim oldcoil As New CoilData With {.CoilFile = New FileData With {.Fullfilename = coilocc.PartFileName, .Shortname = General.GetShortName(coilocc.PartFileName)}, .Number = no}
        Dim j As Integer = 1
        Dim occcount As Integer = 0

        Try
            Dim coilasmdoc As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(oldcoil.CoilFile.Fullfilename)
            General.seapp.DoIdle()

            'load related coil drawing
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Loading Coil drawing from PDM")
            WSM.CheckoutCircs(oldcoil.CoilFile.Shortname.Substring(0, 10), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            'find BOMItem of new order
            Order.GetCoilBOMItem(General.currentunit.Coillist(no - 1), General.currentunit.Coillist(no - 1).Circuits.First)
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Coil" + no.ToString + " ERPCode: " + General.currentunit.Coillist(no - 1).BOMItem.Item)

            oldcoil.CoilFile.LNCode = SEPart.GetSetCustomProp(coilasmdoc, "GUE_Item", "", "get")

            'get added components, but only once → use partlist and asmlist
            For i As Integer = 0 To partlist.Count - 1
                Dim partname As String = partlist(i) + ":1"
                CheckPart(partname, coilasmdoc, oldcoil.Occlist, no, "Coil", no, occcount)
            Next

            For i As Integer = 0 To asmlist.Count - 1
                Dim asmname As String = asmlist(i) + ":1"
                Dim asmocc As SolidEdgeAssembly.Occurrence = SEAsm.CheckOccExists(coilasmdoc.Occurrences, asmname)
                If asmocc IsNot Nothing Then
                    Dim oldconsys As New ConSysData With {.ConSysFile = New FileData With {.Fullfilename = asmocc.PartFileName, .Shortname = General.GetShortName(asmocc.PartFileName)}, .Circnumber = j}

                    oldcoil.Occlist.Add(New PartData With {.Occname = General.GetShortName(asmocc.PartFileName), .Configref = "", .Occindex = asmocc.Index})

                    'find BOMItem in new order
                    With General.currentunit
                        .Coillist(no - 1).ConSyss(i).BOMItem = Order.GetConsysBOMItem(.Coillist(no - 1).ConSyss(i), .Coillist(no - 1), .Coillist(no - 1).Circuits(i).CircuitType)
                        RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Consys" + no.ToString + "_" + (i + 1).ToString + " ERPCode: " + .Coillist(no - 1).ConSyss(i).BOMItem.Item)

                        RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "Loading Consys drawing from PDM")
                        WSM.CheckoutCircs(oldconsys.ConSysFile.Shortname.Substring(0, 10), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    End With

                    TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Consys" + no.ToString + "_" + j.ToString, "ConSys - " + General.GetShortName(asmocc.PartFileName), "Master\Coil" + no.ToString + "\")
                    RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), "New Consys item in Coil" + no.ToString + " - " + General.GetShortName(asmocc.PartFileName))

                    Dim consysasmdoc As SolidEdgeAssembly.AssemblyDocument = General.seapp.Documents.Open(asmocc.PartFileName)
                    General.seapp.DoIdle()

                    For k As Integer = 0 To partlist.Count - 1
                        Dim partname As String = partlist(k) + ":1"
                        CheckPart(partname, consysasmdoc, oldconsys.Occlist, j, "Coil" + no.ToString + "\Consys", no, occcount)
                    Next

                    oldcoil.ConSyss.Add(oldconsys)
                    General.seapp.Documents.CloseDocument(oldconsys.ConSysFile.Fullfilename, SaveChanges:=False, DoIdle:=True)
                    j += 1
                End If
            Next

            General.seapp.Documents.CloseDocument(oldcoil.CoilFile.Fullfilename, SaveChanges:=False, DoIdle:=True)

        Catch ex As Exception
            Debug.Print(ex.ToString)
            RTBLog.Invoke(New UpdateLog(AddressOf UpdateRTB), ex.ToString)
        End Try

        Return oldcoil
    End Function

    Private Sub CheckPart(partname As String, asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef occlist As List(Of PartData), consysno As Integer, refname As String, coilno As Integer, ByRef occount As Integer)
        Dim partocc As SolidEdgeAssembly.Occurrence = SEAsm.CheckOccExists(asmdoc.Occurrences, partname)
        Dim treepath As String

        If partocc IsNot Nothing Then
            occount += 1
            Dim erpcode As String = ""
            Dim CDB_de As String = ""
            Dim CSG As String = ""
            Dim partno As String = ""
            SEPart.GetSetCustomProp(partocc.PartDocument, "CDB_ERP_Artnr.", erpcode, "get")
            SEPart.GetSetCustomProp(partocc.PartDocument, "CDB_Benennung_de", CDB_de, "get")
            SEPart.GetSetCustomProp(partocc.PartDocument, "CSG", CSG, "get")
            SEPart.GetSetCustomProp(partocc.PartDocument, "CDB_teilenummer", partno, "get")

            If (CDB_de.Contains("Stutzen") Or CDB_de.Contains("Bogen")) And erpcode = "" Then
                erpcode = Database.GetValue("CSG.DB_CPs", "ERPCode", "Article_Number", partno)
            End If
            occlist.Add(New PartData With {.Occname = General.GetShortName(partocc.PartFileName), .Configref = erpcode, .Occindex = partocc.Index, .BOMAGP = CSG})

            If refname = "Coil" Then
                treepath = "Master\" + refname + coilno.ToString + "\"
            Else
                treepath = "Master\" + refname + coilno.ToString + "_" + consysno.ToString + "\"
            End If

            If erpcode <> "" Then
                TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Part" + coilno.ToString + occount.ToString, "Standard part - " + erpcode, treepath)
            ElseIf CSG = "1" Then
                TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Part" + coilno.ToString + occount.ToString, "CSG/AGP part - " + CDB_de + " - " + partno, treepath)
            Else
                TV1.Invoke(New UpdateTree(AddressOf AddTreeNode), TV1, "Part" + coilno.ToString + occount.ToString, "Custom part - " + General.GetShortName(partocc.PartFileName), treepath)
            End If

        End If

    End Sub

    Private Function GetConsysOcclist(oldname As String) As List(Of PartData)
        For i As Integer = 0 To oldunit.Coillist.Count - 1
            For j As Integer = 0 To oldunit.Coillist(i).Occlist.Count - 1
                If oldname = oldunit.Coillist(i).Occlist(j).Occname Then
                    Return oldunit.Coillist(i).ConSyss(j).Occlist
                End If
            Next
        Next
    End Function

    Private Function GetCoilOcclist(oldname As String) As List(Of PartData)
        For i As Integer = 0 To oldunit.Coillist.Count - 1
            If oldunit.Coillist(i).CoilFile.Fullfilename = oldname Then
                Return oldunit.Coillist(i).Occlist
            End If
        Next
    End Function

    Private Sub ResetProgress()
        Try
            oldunit.Coillist.Clear()
            TV1.Nodes.Clear()
            partlist.Clear()
            asmlist.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CopyProps_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        WriteDatatoFile(General.currentjob.OrderDir + "\Close_oldunit")
    End Sub

    Private Sub BExport_Click(sender As Object, e As EventArgs) Handles BExport.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If Directory.Exists(General.currentjob.OrderDir) Then
                With General.currentunit
                    .DecSeperator = General.decsym
                    .OrderData = General.currentjob
                    .IsProdavit = General.isProdavit
                    .APPVersion = General.apptime.ToString
                    .DLLVersion = General.dlltime.ToString
                End With
                General.WriteDatatoFile(General.currentjob.OrderDir + "\Export_newunit")
                WriteDatatoFile(General.currentjob.OrderDir + "\Export_oldunit")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BImport_Click(sender As Object, e As EventArgs) Handles BImport.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Dim payload As String = OrderProps.ImportJsonFile(TBMaster.Text)
        If payload <> "" Then
            General.currentunit = General.DeserializeUnit(payload)
            'update BOMList (currentunit contains the old one)
            General.currentunit.BOMList = Order.ConstructBOM(Order.LNProdOrder("web_bom", "dxv?0cc5xot7jcz-1gAQ", TBOrder1.Text, TBPos1.Text, General.currentjob.Plant))
            BAnalyze.Enabled = True
            BRename.Enabled = True
        End If
    End Sub

    Private Sub BReconnect_Click(sender As Object, e As EventArgs) Handles BReconnect.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            General.ReleaseObject(General.seapp)
            Do
                General.seapp = SEUtils.ReConnect()
            Loop Until General.seapp IsNot Nothing
            General.seapp.DisplayAlerts = False
            General.GetLanguage()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BCopy_Click(sender As Object, e As EventArgs) Handles BCopy.Click
        TBOrder2.Text = TBOrder1.Text
        TBNr2.Text = TBNr1.Text
    End Sub

    Private Sub CheckAGP_CheckedChanged(sender As Object, e As EventArgs) Handles CheckAGP.CheckedChanged
        agpnew = CheckAGP.Checked
    End Sub

    Private Sub WriteDatatoFile(path As String)
        Dim payload, filename As String
        Dim utime As Integer

        Try
            oldunit.APPVersion = File.GetLastWriteTimeUtc(Application.StartupPath + "\CSG.exe")
            oldunit.DLLVersion = File.GetLastWriteTimeUtc(Application.StartupPath + "\CSGCore.dll")
            utime = (Date.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
            filename = path + "_" + utime.ToString + ".json"
            payload = General.CreatePayload(oldunit)
            My.Computer.FileSystem.WriteAllText(filename, payload, False)
        Catch ex As Exception

        End Try
    End Sub


End Class