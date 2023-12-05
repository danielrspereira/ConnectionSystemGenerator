Imports CSGCore
Imports SolidEdgeGeometry

Public Class UnitProps
    Public selectedcoil As CoilData
    Public selectedcirc As CircuitData
    Public tubesheet As New SheetData
    Public saveok As Boolean = False

    Private Sub ResetUI()
        CBCoil.Items.Clear()
        CBCircuit.Items.Clear()
    End Sub

    Private Sub DisableButtons()
        Dim blist As New List(Of Button) From {BBuild, BDrawing, BReconnect, BEmptyConsys}
        For Each b In blist
            b.Enabled = False
        Next
    End Sub

    Private Sub RefreshData(item)

        Select Case item
            Case "coil"
                selectedcoil = General.currentunit.Coillist(CInt(CBCoil.Text) - 1)
            Case "circuit"
                selectedcirc = selectedcoil.Circuits(CInt(CBCircuit.Text) - 1)
        End Select
    End Sub

    Private Sub UnitProps_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If General.isProdavit Or OrderProps.CBCSGMode.Text.Contains("Continue") Then
            'reset data
            ResetUI()
            If General.currentjob.Prio = 5 Then
                DisableButtons()
                BSave.Visible = True
            End If
            For i As Integer = 1 To General.currentunit.Coillist.Count
                CBCoil.Items.Add(i.ToString)
            Next
            tubesheet = General.currentunit.TubeSheet
            Try
                CBCoil.Text = "1"
                CBCircuit.Text = "1"
            Catch ex As Exception

            End Try
        End If

        If General.currentunit.UnitDescription = "Dual" Then
            LGap.Visible = True
            TBGap.Visible = True
        End If

        General.GetLanguage()
        Text = "Order: " + General.currentjob.OrderNumber + " Pos.: " + General.currentjob.OrderPosition
        If OrderProps.CBCSGMode.Text.Contains("Continue") Then
            BReset.Visible = True
            BSave.Visible = True
        End If
        If General.username <> "mlewin" And General.username <> "csgen" Then
            BEmptyConsys.Visible = False
        End If
    End Sub

    Private Sub CBCoil_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBCoil.SelectedIndexChanged

        If CBCoil.Text <> "" Then
            Dim tempno As String = CBCircuit.Text
            CBCircuit.Items.Clear()
            ConSysProps.CBConSys.Items.Clear()
            RefreshData("coil")
            If selectedcoil.Circuits.Count > 0 Then
                For i As Integer = 1 To selectedcoil.Circuits.Count
                    CBCircuit.Items.Add(i.ToString)
                    CBCircuit.Text = "1"
                    ConSysProps.CBConSys.Items.Add(i.ToString)
                Next
            End If
            If tempno <> "" Then
                CBCircuit.Text = tempno
            End If
            'fill UI with coil data
            With selectedcoil
                CBAlignment.Text = .Alignment
                TBFinnedLength.Text = General.TextToDouble(.FinnedLength)
                TBFinnedHeight.Text = General.TextToDouble(.FinnedHeight)
                TBFinnedDepth.Text = General.TextToDouble(.FinnedDepth)
                TBDPDMID.Text = .EDefrostPDMID
                TBBTCoil.Text = .NoBlindTubes.ToString
                .CoilFile.Orderno = General.currentjob.OrderNumber
                .CoilFile.Orderpos = General.currentjob.OrderPosition
                .CoilFile.Projectno = General.currentjob.ProjectNumber
                If General.currentunit.UnitDescription = "Dual" Then
                    TBGap.Text = General.TextToDouble(.Gap)
                End If
            End With
            With tubesheet
                CBTSMat.Text = .MaterialCodeLetter
                CBTSWT.Text = .Thickness.ToString.Replace(".", ",")
                CBPunchingType.Text = .PunchingType.ToString
            End With

            If selectedcirc IsNot Nothing Then
                Order.GetCoilBOMItem(selectedcoil, selectedcirc)
                If selectedcoil.BOMItem.Item = "" And General.currentunit.BOMList.Count > 2 Then
                    'select manually
                    CoilSelection.Activate()
                    CoilSelection.Show()
                    CoilSelection.caller = "create"
                End If
            End If
        End If

    End Sub

    Private Sub CBCircuit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBCircuit.SelectedIndexChanged

        If CBCircuit.Text <> "" Then
            RefreshData("circuit")
            Dim circtype As String = GetSetCBCircText(selectedcirc.CircuitType, "set")

            With selectedcirc
                TBPDMID1.Text = .PDMID
                TBQuantity1.Text = .Quantity.ToString
                TBDistr.Text = .NoDistributions.ToString
                TBPasses.Text = .NoPasses.ToString
                CBLocation.Text = .ConnectionSide
                CBFinType.Text = .FinType
                CBOP.Text = .Pressure.ToString
                CBCTMaterial.Text = .CoreTube.Material
                TBCTOverhang.Text = .CoreTubeOverhang.ToString
                CBCType.Text = circtype
                CBFin.Checked = .CustomCirc
            End With
            If selectedcirc.CircuitNumber = 0 Then
                selectedcirc.CircuitNumber = CInt(CBCircuit.Text)
            End If
            If selectedcirc.CircuitType.Contains("Defrost") Then
                TBDPDMID.Text = selectedcirc.SupportPDMID
            Else
                TBDPDMID.Text = selectedcoil.EDefrostPDMID
            End If
            ConSysProps.CBConSys.Text = CBCircuit.Text
        End If

    End Sub

    Private Sub CBTSMat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTSMat.SelectedIndexChanged
        tubesheet.MaterialCodeLetter = CBTSMat.Text
    End Sub

    Private Sub CBTSWT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBTSWT.SelectedIndexChanged
        tubesheet.Thickness = General.TextToDouble(CBTSWT.Text)
    End Sub

    Private Sub CBPunchingType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBPunchingType.SelectedIndexChanged
        tubesheet.PunchingType = CInt(CBPunchingType.Text)
    End Sub

    Private Sub BAddCoil_Click(sender As Object, e As EventArgs) Handles BAddCoil.Click
        Dim itemcount As Integer = CBCoil.Items.Count
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            Dim newcoil As New CoilData With {
                .Number = itemcount + 1,
                .Alignment = CBAlignment.Text,
                .FinnedLength = General.TextToDouble(TBFinnedLength.Text),
                .FinnedHeight = General.TextToDouble(TBFinnedHeight.Text),
                .FinnedDepth = General.TextToDouble(TBFinnedDepth.Text),
                .EDefrostPDMID = TBDPDMID.Text
            }
            If TBBTCoil.Text = "" Then
                newcoil.NoBlindTubes = 0
            Else
                newcoil.NoBlindTubes = CInt(TBBTCoil.Text)
            End If

            If Not General.isProdavit AndAlso General.currentunit.UnitDescription = "VShape" AndAlso newcoil.FinnedHeight < 1100 Then
                General.currentunit.UnitSize = "Compact"
            End If

            selectedcoil = newcoil

            CBCoil.Items.Add(newcoil.Number.ToString)
            General.currentunit.Coillist.Add(newcoil)
            CBCoil.Text = newcoil.Number.ToString
            General.currentunit.TubeSheet = tubesheet
            General.coillist.Add(selectedcoil)
        Catch ex As Exception
            MsgBox("Error adding coil")
        End Try

    End Sub

    Private Sub BEditCoil_Click(sender As Object, e As EventArgs) Handles BEditCoil.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        'update the current coil data
        Try
            With selectedcoil
                .Number = CInt(CBCoil.Text)
                .Alignment = CBAlignment.Text
                .FinnedLength = General.TextToDouble(TBFinnedLength.Text)
                .FinnedHeight = General.TextToDouble(TBFinnedHeight.Text)
                .FinnedDepth = General.TextToDouble(TBFinnedDepth.Text)
                If TBBTCoil.Text = "" Then
                    .NoBlindTubes = 0
                Else
                    .NoBlindTubes = CInt(TBBTCoil.Text)
                End If
                .EDefrostPDMID = TBDPDMID.Text
                If selectedcirc IsNot Nothing Then
                    If selectedcirc.CircuitType = "Defrost" Then
                        .EDefrostPDMID = TBDPDMID.Text
                    End If
                End If
                If General.currentunit.UnitDescription = "Dual" Then
                    .Gap = General.TextToDouble(TBGap.Text)
                End If
            End With

            If Not General.isProdavit AndAlso General.currentunit.UnitDescription = "VShape" AndAlso selectedcoil.FinnedHeight < 1100 Then
                General.currentunit.UnitSize = "Compact"
            Else
                General.currentunit.UnitSize = ""
            End If

            General.currentunit.TubeSheet = tubesheet

        Catch ex As Exception
            MsgBox("Error reading coil data")
        End Try
    End Sub

    Private Sub BAddCircuit_Click(sender As Object, e As EventArgs) Handles BAddCircuit.Click
        Dim itemcount As Integer = CBCircuit.Items.Count + 1
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If CBCoil.Text <> "" Then
                Dim newcircuit As New CircuitData With {
            .Coilnumber = CBCoil.Text,
            .CircuitType = GetSetCBCircText(CBCType.Text, "get"),
            .IsOnebranchEvap = False,
            .PDMID = TBPDMID1.Text,
            .Quantity = TBQuantity1.Text,
            .NoDistributions = TBDistr.Text,
            .NoPasses = TBPasses.Text,
            .ConnectionSide = CBLocation.Text,
            .FinType = CBFinType.Text,
            .Pressure = CBOP.Text,
            .CustomCirc = CBFin.Checked,
            .Orbitalwelding = CheckOrbital.Checked,
            .CoreTubeOverhang = General.TextToDouble(TBCTOverhang.Text),
            .CircuitNumber = itemcount
            }
                'create the core tube
                Dim ctdiameter As Double = GNData.GetTubeDiameter(newcircuit.FinType)
                Dim materialcode As String = GNData.GetMaterialcode(CBCTMaterial.Text, "coretube")
                Dim ctwt As Double = Database.GetTubeThickness("CoreTube", ctdiameter, materialcode, newcircuit.Pressure)
                Dim newct As New TubeData With {
            .Diameter = ctdiameter,
            .Material = CBCTMaterial.Text,
            .Materialcodeletter = materialcode,
            .WallThickness = ctwt}
                newcircuit.CoreTube = newct

                'get pitch
                Dim finpitch() As Double = GNData.GetFinPitch(selectedcoil.Alignment, newcircuit.FinType)
                newcircuit.PitchX = finpitch(0)
                newcircuit.PitchY = finpitch(1)

                If General.currentunit.ApplicationType = "Evaporator" And (General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or newcircuit.Pressure <= 16) And newcircuit.NoDistributions = 1 And newcircuit.Quantity = 1 Then
                    newcircuit.IsOnebranchEvap = True
                End If

                selectedcirc = newcircuit
                selectedcoil.Circuits.Add(selectedcirc)

                'add connection system
                selectedcoil.ConSyss.Add(New ConSysData With {.Circnumber = itemcount,
                                                .OutletHeaders = New List(Of HeaderData) From {New HeaderData},
                                                .InletHeaders = New List(Of HeaderData) From {New HeaderData},
                                                .OutletNipples = New List(Of TubeData) From {New TubeData},
                                                .InletNipples = New List(Of TubeData) From {New TubeData}})

                CBCircuit.Items.Add(itemcount.ToString)
                ConSysProps.CBConSys.Items.Add(itemcount.ToString)
                CBCircuit.Text = itemcount.ToString

                'calculate RR and RL
                With selectedcoil
                    Dim pitch1, pitch2 As Double
                    If .Alignment = "vertical" Then
                        pitch1 = selectedcirc.PitchX
                        pitch2 = selectedcirc.PitchY
                    Else
                        pitch1 = selectedcirc.PitchY
                        pitch2 = selectedcirc.PitchX
                    End If

                    .NoRows = Math.Round(.FinnedDepth / pitch1)
                    .NoLayers = Math.Round(.FinnedHeight / pitch2)
                End With
                If TBDPDMID.Text = "" And selectedcirc.CircuitType.Contains("Defrost") Then
                    MsgBox("Enter PDMID for support tube position in the defrost textbox.")
                Else
                    selectedcirc.SupportPDMID = TBDPDMID.Text
                End If
                General.circuitlist.Add(selectedcirc)

                If selectedcirc.NoDistributions > 1 Then
                    With selectedcoil.ConSyss(itemcount - 1)
                        .OutletHeaders.First.Tube.Quantity = 1
                        .OutletHeaders.First.Tube.HeaderType = "outlet"
                        .OutletHeaders.First.Nippletubes = 1
                        .OutletNipples.First.HeaderType = "outlet"

                        If General.currentunit.ApplicationType = "Condenser" Then
                            .InletHeaders.First.Tube.Quantity = 1
                            .InletHeaders.First.Tube.HeaderType = "inlet"
                            .InletHeaders.First.Nippletubes = 1
                            .InletNipples.First.HeaderType = "inlet"
                        End If
                    End With
                End If

                If Not General.isProdavit And selectedcoil.BOMItem.Item = "" Then
                    Order.GetCoilBOMItem(selectedcoil, selectedcirc)
                    If selectedcoil.BOMItem.Item = "" And General.currentunit.BOMList.Count > 2 Then
                        'select manually
                        CoilSelection.Activate()
                        CoilSelection.Show()
                        CoilSelection.caller = "create"
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox("Error adding circuit")
        End Try
    End Sub

    Private Sub BEditCircuit_Click(sender As Object, e As EventArgs) Handles BEditCircuit.Click
        'update the current circuit data
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            With selectedcirc
                .Coilnumber = CInt(CBCoil.Text)
                .CircuitType = GetSetCBCircText(CBCType.Text, "get")
                .PDMID = TBPDMID1.Text
                .Quantity = CInt(TBQuantity1.Text)
                .NoDistributions = CInt(TBDistr.Text)
                .NoPasses = CInt(TBPasses.Text)
                .ConnectionSide = CBLocation.Text
                .FinType = CBFinType.Text
                .Pressure = CInt(CBOP.Text)
                .CoreTubeOverhang = General.TextToDouble(TBCTOverhang.Text)
                .CircuitNumber = CInt(CBCircuit.Text)
                .CustomCirc = CBFin.Checked
                .Orbitalwelding = CheckOrbital.Checked
            End With

            'create the core tube
            Dim ctdiameter As Double = GNData.GetTubeDiameter(selectedcirc.FinType)
            Dim materialcode As String = GNData.GetMaterialcode(CBCTMaterial.Text, "coretube")
            Dim ctwt As Double = Database.GetTubeThickness("CoreTube", ctdiameter, materialcode, selectedcirc.Pressure)
            Dim newct As New TubeData With {
                .Diameter = ctdiameter,
                .Material = CBCTMaterial.Text,
                .Materialcodeletter = materialcode,
                .WallThickness = ctwt}
            selectedcirc.CoreTube = newct

            'get pitch
            Dim finpitch() As Double = GNData.GetFinPitch(selectedcoil.Alignment, selectedcirc.FinType)
            If selectedcirc.CircuitType = "Defrost" Then
                finpitch(0) *= 2
                finpitch(1) *= 2
            End If
            selectedcirc.PitchX = finpitch(0)
            selectedcirc.PitchY = finpitch(1)

            'calculate RR and RL
            With selectedcoil
                Dim pitch1, pitch2 As Double
                If .Alignment = "vertical" Then
                    pitch1 = selectedcirc.PitchX
                    pitch2 = selectedcirc.PitchY
                Else
                    pitch1 = selectedcirc.PitchY
                    pitch2 = selectedcirc.PitchX
                End If

                .NoRows = Math.Round(.FinnedDepth / pitch1)
                .NoLayers = Math.Round(.FinnedHeight / pitch2)
            End With
            selectedcirc.SupportPDMID = TBDPDMID.Text
            If selectedcirc.CircuitType.Contains("Defrost") And TBPDMID1.Text = "" Then
                MsgBox("Missing information about PDMID of support tube drawing.")
            End If
        Catch ex As Exception
            MsgBox("Error saving circuit")
        End Try
    End Sub

    Private Function GetSetCBCircText(circtype As String, mode As String) As String
        Dim cbtext As String

        If mode = "set" Then
            If circtype = "Defrost" Then
                cbtext = "Brine Defrost"
            ElseIf circtype = "Subcooler" Then
                cbtext = "Integrated Subcooler"
            Else
                cbtext = circtype
            End If
        Else
            If circtype = "Brine Defrost" Then
                cbtext = "Defrost"
            ElseIf circtype = "Integrated Subcooler" Then
                cbtext = "Subcooler"
            Else
                cbtext = circtype
            End If
        End If
        Return cbtext
    End Function

    Private Sub BBuild_Click(sender As Object, e As EventArgs) Handles BBuild.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If CBCoil.Text <> "" Then
                If selectedcoil.CoilFile.Fullfilename = "" And Not selectedcirc.CircuitType.Contains("Defrost") Then
                    Unit.CreateCoil(selectedcoil, selectedcirc)
                End If

                If selectedcirc.CircuitType.Contains("Defrost") And General.currentunit.ApplicationType = "Evaporator" Then
                    If Not IO.File.Exists(General.currentjob.Workspace + "\CoilSupporttube" + selectedcoil.Number.ToString + ".par") Then
                        'adjust coil
                        Unit.AdjustCoilDefrost(selectedcoil, selectedcirc)
                    End If
                    Unit.CreateBrineCircuit(selectedcoil, selectedcirc)
                Else
                    Unit.CreateCircuit(selectedcoil, selectedcirc)
                End If

                'assemble coils in same assembly
                General.currentunit.UnitFile.Fullfilename = General.currentjob.Workspace + "\Masterassembly.asm"
                If Not IO.File.Exists(General.currentunit.UnitFile.Fullfilename) Then
                    'create the assembly
                    Unit.CreateMasterAssembly()
                End If

                'add coil
                Unit.AddAssembly(General.currentunit.UnitFile.Fullfilename, selectedcoil.CoilFile.Fullfilename, General.currentunit.Occlist)

                If CBCoil.Text <> "1" Then
                    'reposition
                    SEAsm.RepositionAssembly(General.seapp.ActiveDocument, selectedcoil)
                End If

            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

    End Sub

    Private Sub BDrawing_Click(sender As Object, e As EventArgs) Handles BDrawing.Click
        Dim create2d As Boolean = True
        Dim missings As New List(Of String)
        Dim missingtext As String = ""

        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If selectedcoil.CoilFile.Fullfilename <> "" Then
            General.seapp.DisplayAlerts = False
            General.seapp.Documents.Close()
            General.seapp.DoIdle()
            'if dft exists already, dont proceed
            If IO.File.Exists(selectedcoil.CoilFile.Fullfilename.Replace(".asm", ".dft")) Then
                MsgBox("Coil has a referenced drawing already")
            Else
                'check if all consys assemblies have been created
                For i As Integer = 1 To CBCircuit.Items.Count
                    If Not IO.File.Exists(General.currentjob.Workspace + "\Consys" + i.ToString + "_" + selectedcoil.Number.ToString + ".asm") Then
                        missings.Add("Consys " + i.ToString)
                        create2d = False
                    End If
                Next
                If create2d Then
                    'create view configs
                    If SEAsm.SingleBowViews(selectedcoil) Then
                        If SEDrawing.CreateDrawingCoil(selectedcoil) Then
                            BSave.Visible = True
                        Else
                            MsgBox("Error creating the coil drawing." + vbNewLine +
                                   "1) Restart Solid Edge" + vbNewLine +
                                   "2) Reconnect CSG to Solid Edge" + vbNewLine +
                                   "3) Click the Reset button for the coil 2D" + vbNewLine +
                                   "4) Click Create 2D again")
                        End If
                    Else
                        MsgBox("Error creating view configurations")
                    End If
                    BReset.Enabled = True
                    BReset.Visible = True
                Else
                    For Each m In missings
                        missingtext += vbNewLine + m
                    Next
                    MsgBox("Missing assembly for:" + missingtext)
                End If
            End If
        End If
    End Sub

    Private Sub BConSys_Click(sender As Object, e As EventArgs) Handles BConSys.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If CBCircuit.Text <> "" And CBCoil.Text <> "" Then
            'delete all consys items and add them again
            ConSysProps.CBConSys.Items.Clear()
            For i As Integer = 1 To CBCircuit.Items.Count
                For Each cons In selectedcoil.ConSyss
                    If cons.Circnumber = i Then
                        ConSysProps.CBConSys.Items.Add(i.ToString)
                    End If
                Next
            Next
            ConSysProps.Show()
            ConSysProps.Activate()
        Else
            MsgBox("Add a Circuit first")
        End If
    End Sub

    Private Sub BSave_Click(sender As Object, e As EventArgs) Handles BSave.Click
        Dim missingtext, masterassembly As String
        Dim missings As New List(Of String)

        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If selectedcoil.CoilFile.Fullfilename <> "" Then
                'check, if consys exists - if not, ask for confirmation to save
                If Not saveok Then
                    For i As Integer = 1 To CBCircuit.Items.Count
                        If Not IO.File.Exists(General.currentjob.Workspace + "\Consys" + i.ToString + "_" + selectedcoil.Number.ToString + ".dft") Then
                            missings.Add("Coil " + selectedcoil.Number.ToString + " - Consys " + i.ToString)
                            saveok = False
                        End If
                    Next

                    If missings.Count > 0 Then
                        missingtext = "No connection system drawing for:" + vbNewLine
                        For i As Integer = 0 To missings.Count - 1
                            missingtext += missings(i) + vbNewLine
                        Next
                        ConfirmSave.Labeltext = missingtext
                    Else
                        saveok = True
                    End If
                End If

                If IO.File.Exists(selectedcoil.CoilFile.Fullfilename.Replace("asm", "dft")) Then
                    If saveok Then
                        masterassembly = General.GetFullFilename(General.currentjob.Workspace, "Masterassembly", "asm")
                        'open document
                        If IO.File.Exists(masterassembly) Then
                            'create a backup
                            Unit.PrepareForBackup()
                            BExport.PerformClick()

                            General.seapp.Documents.Open(masterassembly)
                            General.seapp.DoIdle()

                            'save process 
                            If WSM.SaveDFT() Then
                                'wait until finished
                                WSM.WaitforWSMDialog()

                                'change the coilfile.fullfilename and all consysfile.fullfilename
                                General.RenameFiles()

                                BConSys.Enabled = False
                            Else
                                MsgBox("Error in the saving process")
                            End If
                        Else
                            MsgBox("Missing Masterassembly")
                        End If
                    Else
                        ConfirmSave.Show()
                        ConfirmSave.Activate()
                    End If
                Else
                    MsgBox("Missing Coil assembly" + vbNewLine + "File: " + selectedcoil.CoilFile.Fullfilename)
                End If
            ElseIf General.currentjob.Prio = 5 Then
                General.currentjob.Uid = Database.AddJob(General.currentjob)
                General.SendDataToWebservice("http://deffbswap14.europe.guentner-corp.com:800/api/values/save-input", "POST")
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
            MsgBox("Error saving")
        End Try
    End Sub

    Private Sub BReset_Click(sender As Object, e As EventArgs) Handles BReset.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        'delete coil.dft files and reset the information in the coildata class
        Try
            If IO.File.Exists(selectedcoil.CoilFile.Fullfilename.Replace(".asm", ".dft")) Then
                IO.File.Delete(selectedcoil.CoilFile.Fullfilename.Replace(".asm", ".dft"))
            End If
            Threading.Thread.Sleep(400)
            If IO.File.Exists(selectedcoil.CoilFile.Fullfilename) Then
                BReset.Enabled = False
            End If
            selectedcoil.Frontbowids.Clear()
            selectedcoil.Backbowids.Clear()
        Catch ex As Exception

        End Try

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

    Private Sub CBCoil_SelectedValueChanged(sender As Object, e As EventArgs) Handles CBCoil.SelectedValueChanged
        BConSys.Enabled = True
    End Sub

    Private Sub BExport_Click(sender As Object, e As EventArgs) Handles BExport.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        Try
            If IO.Directory.Exists(General.currentjob.OrderDir) Then
                With General.currentunit
                    .DecSeperator = General.decsym
                    .OrderData = General.currentjob
                    .IsProdavit = General.isProdavit
                    .APPVersion = General.apptime.ToString
                    .DLLVersion = General.dlltime.ToString
                End With
                General.WriteDatatoFile(General.currentjob.OrderDir + "\LogExport")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CBOP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBOP.SelectedIndexChanged
        If CInt(CBOP.Text) <= 16 Then
            ConSysProps.CheckFlange.Visible = True
            ConSysProps.CheckValve.Visible = True
        Else
            ConSysProps.CheckFlange.Visible = False
            ConSysProps.CheckValve.Visible = False
        End If
    End Sub

    Private Sub CBCoil_TextChanged(sender As Object, e As EventArgs) Handles CBCoil.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBCoil.Text)
    End Sub

    Private Sub CBCircuit_TextChanged(sender As Object, e As EventArgs) Handles CBCircuit.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBCircuit.Text)
    End Sub

    Private Sub CBAlignment_TextChanged(sender As Object, e As EventArgs) Handles CBAlignment.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBAlignment.Text)
    End Sub

    Private Sub CBTSMat_TextChanged(sender As Object, e As EventArgs) Handles CBTSMat.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBTSMat.Text)
    End Sub

    Private Sub CBTSWT_TextChanged(sender As Object, e As EventArgs) Handles CBTSWT.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBTSWT.Text)
    End Sub

    Private Sub CBPunchingType_TextChanged(sender As Object, e As EventArgs) Handles CBPunchingType.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBPunchingType.Text)
    End Sub

    Private Sub CBCType_TextChanged(sender As Object, e As EventArgs) Handles CBCType.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBCType.Text)
    End Sub

    Private Sub CBOP_TextChanged(sender As Object, e As EventArgs) Handles CBOP.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBOP.Text)
    End Sub

    Private Sub CBCTMaterial_TextChanged(sender As Object, e As EventArgs) Handles CBCTMaterial.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBCTMaterial.Text)
    End Sub

    Private Sub CBFinType_TextChanged(sender As Object, e As EventArgs) Handles CBFinType.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBFinType.Text)
    End Sub

    Private Sub CBLocation_TextChanged(sender As Object, e As EventArgs) Handles CBLocation.TextChanged
        General.CreateActionLogEntry(Me.Name, sender.name, "changed", CBLocation.Text)
    End Sub

    Private Sub CBFin_CheckedChanged(sender As Object, e As EventArgs) Handles CBFin.CheckedChanged
        General.CreateActionLogEntry("UnitProps", "CustomCirc", "changed", CBFin.Checked.ToString)
        If selectedcirc IsNot Nothing Then
            selectedcirc.CustomCirc = CBFin.Checked
        End If
    End Sub

    Private Sub CheckOrbital_CheckedChanged(sender As Object, e As EventArgs) Handles CheckOrbital.CheckedChanged
        General.CreateActionLogEntry("UnitProps", "OrbitalWelding", "changed", CheckOrbital.Checked.ToString)
        If CheckOrbital.Checked Then
            TBCTOverhang.Text = "55"
            BEditCircuit.PerformClick()
        End If
    End Sub

    Private Sub CBFinType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBFinType.SelectedIndexChanged
        If CBFinType.Text <> "" And Not General.isProdavit Then
            If CBFinType.Text = "M" Or CBFinType.Text = "N" Then
                CBFin.Checked = False
                CBFin.Visible = False
            Else
                CBFin.Visible = True
            End If
        End If
        If General.username = "mlewin" Then
            If CBFinType.Text = "N" And CBCTMaterial.Text.Contains("V") Then
                CheckOrbital.Visible = True
            Else
                CheckOrbital.Visible = False
                CheckOrbital.Checked = False
            End If
        End If
    End Sub

    Private Sub CBCTMaterial_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBCTMaterial.SelectedIndexChanged
        If General.username = "mlewin" Then
            If CBFinType.Text = "N" And CBCTMaterial.Text.Contains("V") Then
                CheckOrbital.Visible = True
            Else
                CheckOrbital.Visible = False
                CheckOrbital.Checked = False
            End If
        End If
    End Sub

    Private Sub BImport_Click(sender As Object, e As EventArgs) Handles BImport.Click
        Dim payload As String

        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")

        Try
            payload = OrderProps.ImportJsonFile(General.currentjob.PDMID)
            If payload <> "" Then
                BSave.Visible = True
                BImport.Enabled = False
            Else
                General.OverrideData(payload)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub BEmptyConsys_Click(sender As Object, e As EventArgs) Handles BEmptyConsys.Click
        General.CreateActionLogEntry(Me.Name, sender.name, "pressed")
        If CBCircuit.Text <> "" And CBCoil.Text <> "" Then
            'delete all consys items and add them again
            ConSysProps.CBConSys.Items.Clear()
            For i As Integer = 1 To CBCircuit.Items.Count
                For Each cons In selectedcoil.ConSyss
                    If cons.Circnumber = i Then
                        ConSysProps.CBConSys.Items.Add(i.ToString)
                    End If
                Next
            Next
            'check if coil assembly exists and consys assembly not
            ConSysProps.selectedconsys = selectedcoil.ConSyss(CBCircuit.Text - 1)
            If ConSysProps.selectedconsys.ConSysFile.Fullfilename = "" Then
                ConSysProps.selectedconsys.BOMItem = Order.GetConsysBOMItem(ConSysProps.selectedconsys, selectedcoil, selectedcirc.CircuitType)
                Unit.CreateConsys(selectedcoil, selectedcirc, ConSysProps.selectedconsys, True)

                Unit.AddAssembly(selectedcoil.CoilFile.Fullfilename, ConSysProps.selectedconsys.ConSysFile.Fullfilename, selectedcoil.Occlist)
                'checkout template drawing
                WSM.CheckoutCircs("767687", General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

                If General.WaitForFile(General.currentjob.Workspace, "767687", ".dft", 100) Then
                    IO.File.Copy(General.GetFullFilename(General.currentjob.Workspace, "767687", ".dft"), ConSysProps.selectedconsys.ConSysFile.Fullfilename.Replace("asm", "dft"))
                    Threading.Thread.Sleep(1000)

                    General.seapp.Documents.Open(ConSysProps.selectedconsys.ConSysFile.Fullfilename.Replace("asm", "dft"))
                    Threading.Thread.Sleep(1000)

                    General.seapp.DoIdle()
                    CondenserDrawings.EmptyConsys(General.seapp.ActiveDocument, selectedcoil, ConSysProps.selectedconsys)
                End If
            End If
        Else
            MsgBox("Add a Circuit first")
        End If
    End Sub


End Class