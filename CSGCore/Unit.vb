Imports SED = SolidEdgeDraft

Public Class Unit
    Public Shared supportcoords(), heatingcoords(), heatingbows(), coolingfrontbowprops(), coolingbackbowprops(), defrostfrontbowprops(), defrostbackbowprops() As List(Of Double)
    Public Shared pairlistfront, pairlistback As New List(Of List(Of Integer))
    Public Shared coolingfrontids, coolingbackids, defrostfrontids, defrostbackids, bowkeys, templateids As New List(Of String)
    Public Shared bpairlistfront, bpairlistback, cpairlistfront, cpairlistback As New List(Of List(Of Integer))
    Public Shared brinefrontbows, brinebackbows As New List(Of BowData)
    Public Shared repairfront, repairback As New List(Of Double())
    Public Shared rotatefin As Boolean
    Public Shared stutzendict As New Dictionary(Of String, String)

    Shared Sub CreateMasterAssembly()
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument

        Try
            asmdoc = General.seapp.Documents.Add(ProgID:="SolidEdge.AssemblyDocument")
            General.seapp.DoIdle()
            asmdoc.SaveAs(General.currentjob.Workspace + "\Masterassembly.asm")

            'fill a few file properties
            SEPart.GetSetCustomProp(asmdoc, "CSG", "1", "write")
            SEPart.GetSetCustomProp(asmdoc, "Auftragsnummer", General.currentjob.OrderNumber, "write")
            SEPart.GetSetCustomProp(asmdoc, "Position", General.currentjob.OrderPosition, "write")
            SEPart.GetSetCustomProp(asmdoc, "Order_Projekt", General.currentjob.ProjectNumber, "write")
            SEPart.GetSetCustomProp(asmdoc, "CDB_Benennung_de", "Masterassembly", "write")
            SEPart.GetSetCustomProp(asmdoc, "CDB_Benennung_en", "Masterassembly", "write")
            General.seapp.DoIdle()

            General.currentunit.UnitFile.Fullfilename = asmdoc.FullName

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub AddAssembly(mastername As String, subasmname As String, ByRef occlist As List(Of PartData))
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument

        Try
            asmdoc = General.seapp.Documents.Open(mastername)
            General.seapp.DoIdle()
            'check occurances → avoid a 2nd add
            For Each occ As SolidEdgeAssembly.Occurrence In asmdoc.Occurrences
                Dim occname As String = ""
                Select Case occ.Type.ToString
                    Case "igPart"
                        Dim occdoc As SolidEdgePart.PartDocument = occ.PartDocument
                        occname = occdoc.FullName
                    Case "igSubAssembly"
                        Dim occdoc As SolidEdgeAssembly.AssemblyDocument = occ.PartDocument
                        occname = occdoc.FullName
                End Select
                If occname = subasmname Then
                    General.seapp.Documents.CloseDocument(mastername, SaveChanges:=True, DoIdle:=True)
                    Exit Sub
                End If
            Next

            SEAsm.AddSubAssemblytoMaster(asmdoc, subasmname, occlist)

            If Not subasmname.Contains("Coil2.asm") Then
                General.seapp.Documents.CloseDocument(mastername, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub CreateCoil(ByRef coil As CoilData, ByRef circuit As CircuitData)

        Try
            OleMessageFilter.Register()
            'update jobtable
            Dim sno As Integer = coil.Number * 10 + circuit.CircuitNumber
            Database.UpdateJob(General.currentjob, {"Status"}, {sno.ToString})

            'create coretube
            SEPart.CreateCoreTube(circuit.CoreTube, coil.Number, "CoilCoretube", circuit.Orbitalwelding)

            'create fin
            SEPart.CreateFin(coil, circuit, General.currentunit.TubeSheet, "CoilFin", circuit.CircuitNumber)

            'create the assembly file
            With coil.CoilFile
                .Fullfilename = General.currentjob.Workspace + "\Coil" + coil.Number.ToString + ".asm"
                .CDB_Zusatzbenennung = Order.CreateAddDesignationCoil(circuit, coil)
                .CDB_z_Bemerkung = Order.CreateOrderCommentCoil(circuit, coil)
                .LNCode = coil.BOMItem.Item
                .CDB_de = "Block"
                .CDB_en = "Coil"
                .AGPno = "101"
                .Plant = General.currentjob.Plant
                .Orderno = General.currentjob.OrderNumber
                .Orderpos = General.currentjob.OrderPosition
                .Projectno = General.currentjob.ProjectNumber
            End With
            coil.RotationDirection = Calculation.GetRotationDir(coil, coil.Circuits.First.ConnectionSide)
            SEAsm.CreateSubAssembly(coil, General.currentunit.TubeSheet.Thickness, circuit.CoreTubeOverhang, coil.CoilFile.Fullfilename, circuit.CircuitNumber)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AdjustCoilDefrost(coil As CoilData, ByRef circuit As CircuitData)

        Try
            'create coretube
            SEPart.CreateCoreTube(circuit.CoreTube, coil.Number, "CoilSupporttube", False)

            'create cutouts 
            SEPart.AdjustFinDefrost(coil, circuit, General.currentjob.Workspace + "\CoilFin" + coil.Number.ToString + "1.par")

            'add supporttubes to assembly
            SEAsm.AdjustCoilAssemblyDefrost(coil, circuit, General.currentunit.TubeSheet.Thickness)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AdjustConsysDefrost(coil As CoilData, ByRef circuit As CircuitData)

        Try
            'create coretube
            SEPart.CreateCoreTube(circuit.CoreTube, coil.Number, "ConsysSupporttube", False)

            'create cutouts 
            SEPart.AdjustFinDefrost(coil, circuit, General.currentjob.Workspace + "\ConsysFin" + coil.Number.ToString + circuit.CircuitNumber.ToString + ".par")

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateBrineCircuit(coil As CoilData, ByRef circuit As CircuitData)
        Dim coolingbowprops()() As List(Of Double)
        Dim bowids(), frontbowids, backbowids As List(Of String)

        Try

            WSM.CheckoutCircs(circuit.PDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            'get the bow properties
            coolingbowprops = GetTubeArrangement(coil, circuit, General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, ".dft"), "defrost")
            defrostfrontbowprops = coolingbowprops(0)
            defrostbackbowprops = coolingbowprops(1)

            'get the bow IDs
            If circuit.NoPasses > 1 Then
                bowids = GNData.GetBowIDs(defrostfrontbowprops, defrostbackbowprops, circuit.CoreTube.Material, circuit.FinType, circuit.Pressure, circuit.Orbitalwelding)
                frontbowids = bowids(0)
                backbowids = bowids(1)
                For Each l1value In bowids(2)
                    circuit.L1Levels.Add(l1value)
                Next

                'extend bowprops by L1
                CircProps.ExtendBowProps(defrostfrontbowprops, frontbowids, circuit.L1Levels, circuit.CircuitType.ToLower)
                CircProps.ExtendBowProps(defrostbackbowprops, backbowids, circuit.L1Levels, circuit.CircuitType.ToLower)

                'recalc L1
                defrostfrontbowprops = CircProps.RecalcL1B(cpairlistfront, defrostfrontbowprops, coolingfrontbowprops, frontbowids, circuit.CoreTube.Diameter)
                defrostbackbowprops = CircProps.RecalcL1B(cpairlistback, defrostbackbowprops, coolingbackbowprops, backbowids, circuit.CoreTube.Diameter)

                'recalc L1 again, this time level 2+ of brine must be checked
                defrostfrontbowprops = CircProps.RecalcL1B2(bpairlistfront, defrostfrontbowprops, frontbowids, circuit.L1Levels.Count)
                defrostbackbowprops = CircProps.RecalcL1B2(bpairlistback, defrostbackbowprops, backbowids, circuit.L1Levels.Count)

                'close the drawing
                General.seapp.Documents.CloseDocument(General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, ".dft"), SaveChanges:=False, DoIdle:=True)

                'Create unique bowid list and check out parts
                CheckoutList(General.GetUniqueStrings(frontbowids, backbowids), General.currentjob.Workspace)

                'check for new bow items
                NewBrineBowKeys(defrostfrontbowprops, defrostbackbowprops, frontbowids, backbowids, circuit)

                'assemble bows 
                BuildCoilAsm(circuit, defrostfrontbowprops, defrostbackbowprops, frontbowids, backbowids, coil)

                defrostfrontids = frontbowids
                defrostbackids = backbowids
            End If

        Catch ex As Exception

        End Try

    End Sub

    Shared Sub CreateCircuit(ByRef coil As CoilData, ByRef circuit As CircuitData)
        Dim dftdoc As SED.DraftDocument
        Dim coolingbowprops()(), frontbowprops(), backbowprops() As List(Of Double)
        Dim bowids() As List(Of String)
        Dim frontbowids, backbowids As New List(Of String)
        Dim hairpinlist As New List(Of BOMItem)

        Try
            If circuit.CoreTube.FileName = "" Then
                circuit.CoreTube.FileName = General.currentjob.Workspace + "\CoilCoretube" + coil.Number.ToString + ".par"
            End If
            'get circuit drawing
            If coil.EDefrostPDMID <> "" Then
                WSM.CheckoutCircs(coil.EDefrostPDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                dftdoc = SEDraft.OpenDFT(General.GetFullFilename(General.currentjob.Workspace, coil.EDefrostPDMID, ".dft"))
                GetDefrostArrangement(coil, circuit.ConnectionSide, dftdoc)
            End If

            WSM.CheckoutCircs(circuit.PDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            'wait and check if file exists
            If General.WaitForFile(General.currentjob.Workspace, circuit.PDMID, ".dft", maxloops:=50) = False Then
                If General.username = "csgen" Then
                    Database.UpdateJob(General.currentjob, {"Status", "FinishedTime", "Saved"}, {"-2", Database.ConvertDatetoStr(Date.UtcNow), "false"})
                    Environment.Exit(0)
                Else
                    MsgBox("Error! Could not find the circuiting drawing.")
                End If
            End If

            'get the bow properties
            If circuit.NoPasses > 1 Then
                coolingbowprops = GetTubeArrangement(coil, circuit, General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, ".dft"), "cooling")
                frontbowprops = coolingbowprops(0)
                backbowprops = coolingbowprops(1)

                'get the bow IDs
                bowids = GNData.GetBowIDs(frontbowprops, backbowprops, circuit.CoreTube.Material, circuit.FinType, circuit.Pressure, circuit.Orbitalwelding)
                frontbowids = bowids(0)
                backbowids = bowids(1)
                For Each l1value In bowids(2)
                    circuit.L1Levels.Add(l1value)
                Next

                'extend bowprops by L1
                If frontbowids.Count > 0 Then
                    CircProps.ExtendBowProps(coolingfrontbowprops, frontbowids, circuit.L1Levels, circuit.CircuitType.ToLower)
                End If

                If backbowids.Count > 0 Then
                    CircProps.ExtendBowProps(coolingbackbowprops, backbowids, circuit.L1Levels, circuit.CircuitType.ToLower)
                End If

                'recalc L1
                With circuit
                    If frontbowids.Count > 0 Then
                        coolingfrontbowprops = CircProps.RecalcL1(pairlistfront, coolingfrontbowprops, frontbowids, .L1Levels.Count, .CoreTube.Diameter, .Pressure, .CoreTube.Material, .FinType, .Orbitalwelding)
                    End If

                    If backbowids.Count > 0 Then
                        coolingbackbowprops = CircProps.RecalcL1(pairlistback, coolingbackbowprops, backbowids, .L1Levels.Count, .CoreTube.Diameter, .Pressure, .CoreTube.Material, .FinType, .Orbitalwelding)
                    End If
                End With

                'close the drawing
                General.seapp.Documents.CloseDocument(General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, ".dft"), SaveChanges:=False, DoIdle:=True)

                'check BOMList for Hairpins
                Dim hp_de As List(Of BOMItem) = Order.GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "description"}, {coil.BOMItem.ItemNumber.ToString, "Haarnadel"})

                If hp_de.Count > 0 Then
                    hairpinlist.AddRange(hp_de.ToArray)
                End If

                Dim hp_en As List(Of BOMItem) = Order.GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "description"}, {coil.BOMItem.ItemNumber.ToString, "Hairpin"})

                If hp_en.Count > 0 Then
                    hairpinlist.AddRange(hp_en.ToArray)
                End If

                Dim hppitches As New Dictionary(Of String, Double)
                For Each hairpin In hairpinlist
                    hairpin.Item = hairpin.Item.Replace("_500", "")

                    'get pdmid and check it out
                    Dim pdmid As String = Database.GetValue("CSG.Hairpins", "Article_Number", "ERPCode", hairpin.Item.Trim)
                    Dim pitch As Double = Database.GetValue("CSG.Hairpins", "Pitch", "ERPCode", hairpin.Item.Trim, "double")
                    Dim finnedlength As Double = Database.GetValue("CSG.Hairpins", "FinnedLength", "ERPCode", hairpin.Item.Trim, "double")

                    If Math.Abs(finnedlength - coil.FinnedLength) < 10 Then
                        circuit.Hairpins.Add(New HairpinData With {.ERPCode = hairpin.Item.Trim, .PDMID = pdmid, .Pitch = pitch})
                        If pdmid <> "NULL" And pdmid <> "" Then
                            WSM.CheckoutPart(pdmid, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        End If
                    End If
                Next

                CircProps.ReplaceBows(backbowids, backbowprops, circuit.Hairpins)

                'Create unique bowid list and check out parts
                CheckoutList(General.GetUniqueStrings(frontbowids, backbowids), General.currentjob.Workspace)

                'check for new bow items
                NewBowKeys(coolingfrontbowprops, coolingbackbowprops, frontbowids, backbowids, circuit)

                If circuit.Quantity > 1 And General.currentunit.UnitDescription <> "Dual" Then
                    'extend lists
                    If frontbowids.Count > 0 Then
                        CircProps.ExtendElementaryCirc(frontbowids, coolingfrontbowprops, circuit.Quantity, circuit.CircuitSize, coil.Alignment)
                    End If
                    If backbowids.Count > 0 Then
                        CircProps.ExtendElementaryCirc(backbowids, coolingbackbowprops, circuit.Quantity, circuit.CircuitSize, coil.Alignment)
                    End If
                End If

                coolingfrontids = frontbowids
                coolingbackids = backbowids
            End If

            'assemble bows 
            BuildCoilAsm(circuit, coolingfrontbowprops, coolingbackbowprops, frontbowids, backbowids, coil)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateConsys(ByRef coil As CoilData, ByRef circuit As CircuitData, ByRef consys As ConSysData, Optional isempty As Boolean = False)
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim stutzencoords() As List(Of Double)
        Dim co2split As Boolean = False
        Dim dxsplit As Boolean = False
        Dim rdsplit As Boolean = False
        Dim slist As New List(Of String)

        Try
            'update jobtable
            Dim sno As Integer = coil.Number * 10 + circuit.CircuitNumber + 5 + consys.Circnumber
            Database.UpdateJob(General.currentjob, {"Status", "ModelRange"}, {sno.ToString, General.currentunit.ModelRangeName})

            If circuit.CircuitType = "Defrost" Then
                SEPart.CreateCoreTube(coil.Circuits.First.CoreTube, 2, "ConsysCoretube", False)

                SEPart.CreateFin(coil, coil.Circuits.First, General.currentunit.TubeSheet, "ConsysFin", 2)

                AdjustConsysDefrost(coil, circuit)

            Else
                'create coretube
                SEPart.CreateCoreTube(circuit.CoreTube, circuit.CircuitNumber, "ConsysCoretube", False)

                'create fin
                SEPart.CreateFin(coil, circuit, General.currentunit.TubeSheet, "ConsysFin", circuit.CircuitNumber)
            End If

            'create the assembly
            With consys.ConSysFile
                .Fullfilename = General.currentjob.Workspace + "\Consys" + consys.Circnumber.ToString + "_" + coil.Number.ToString + ".asm"
                .CDB_Zusatzbenennung = Order.CreateAddDesignationConsys(consys, circuit)
                .CDB_z_Bemerkung = Order.CreateOrderCommentConsys(consys, coil.Number)
                .LNCode = consys.BOMItem.Item
                .CDB_de = "Anschlusssystem"
                .CDB_en = "Connection System"
                .AGPno = "101"
                .Plant = General.currentjob.Plant
                .Orderno = General.currentjob.OrderNumber
                .Orderpos = General.currentjob.OrderPosition
                .Projectno = General.currentjob.ProjectNumber
            End With

            SEAsm.CreateSubAssembly(coil, General.currentunit.TubeSheet.Thickness, circuit.CoreTubeOverhang, consys.ConSysFile.Fullfilename, circuit.CircuitNumber)

            If circuit.CoreTube.FileName = "" Then
                circuit.CoreTube.FileName = General.currentjob.Workspace + "\ConsysCoretube" + circuit.CircuitNumber.ToString + ".par"
            End If

            If Not isempty Then
                'create header tubes
                If consys.OutletHeaders.First.Tube.Quantity > 0 Then
                    For i As Integer = 1 To consys.OutletHeaders.Count
                        SEPart.CreateTube(consys.OutletHeaders(i - 1).Tube, General.currentjob.Workspace + "\OutletHeader" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "header", circuit.Pressure)
                    Next
                End If

                If consys.InletHeaders.First.Tube.Quantity > 0 Then
                    For i As Integer = 1 To consys.InletHeaders.Count
                        SEPart.CreateTube(consys.InletHeaders(i - 1).Tube, General.currentjob.Workspace + "\InletHeader" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "header", circuit.Pressure)
                    Next
                End If

                'create nipple tubes
                If consys.OutletNipples.First.Quantity > 0 Then
                    For i As Integer = 1 To consys.OutletNipples.Count
                        SEPart.CreateTube(consys.OutletNipples(i - 1), General.currentjob.Workspace + "\OutletNipple" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "nipple", circuit.Pressure)
                    Next
                End If

                If consys.InletNipples.First.Quantity > 0 Then
                    For i As Integer = 1 To consys.InletNipples.Count
                        SEPart.CreateTube(consys.InletNipples(i - 1), General.currentjob.Workspace + "\InletNipple" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "nipple", circuit.Pressure)
                    Next
                End If

                If General.currentunit.UnitDescription = "Dual" Then
                    'create conjunction tubes
                    If consys.OutletConjunctions.First.Quantity > 0 Then
                        For i As Integer = 1 To consys.OutletConjunctions.Count
                            SEPart.CreateTube(consys.OutletConjunctions(i - 1), General.currentjob.Workspace + "\OutletConjunction" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "nipple", circuit.Pressure)
                        Next
                    End If

                    If consys.InletConjunctions.First.Quantity > 0 Then
                        For i As Integer = 1 To consys.InletConjunctions.Count
                            SEPart.CreateTube(consys.InletConjunctions(i - 1), General.currentjob.Workspace + "\InletConjunction" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "nipple", circuit.Pressure)
                        Next
                    End If

                    'create sifon tube
                    If consys.OilSifons.First.Tube.Quantity > 0 Then
                        For i As Integer = 1 To consys.OilSifons.Count
                            SEPart.CreateTube(consys.OilSifons(i - 1).Tube, General.currentjob.Workspace + "\OutletPot" + i.ToString + "_" + circuit.CircuitNumber.ToString + "_" + coil.Number.ToString + ".par", "nipple", circuit.Pressure)
                        Next
                    End If
                    If circuit.NoDistributions > 1 Then
                        Calculation.AddGADCHeader(consys.OutletHeaders, circuit.Pressure)
                        If circuit.Pressure < 17 Then
                            Calculation.AddGADCHeader(consys.InletHeaders, circuit.Pressure)
                        End If
                    End If
                End If

                If consys.ConType = 2 Then
                    Calculation.GetFlangeLength(consys)
                End If

                'get in/out position
                stutzencoords = CSData.GetCoords(coil, circuit, General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, "dft"))
                If General.currentunit.UnitDescription = "Dual" Then
                    CSData.AssignCoords(stutzencoords, consys, coil.FinnedDepth)
                Else
                    consys.InletHeaders.First.Xlist = stutzencoords(0)
                    consys.InletHeaders.First.Ylist = stutzencoords(1)
                    consys.OutletHeaders.First.Xlist = stutzencoords(2)
                    consys.OutletHeaders.First.Ylist = stutzencoords(3)
                End If

                If circuit.CircuitType.Contains("Defrost") Then
                    If consys.HeaderAlignment = "horizontal" Then
                        'overwrite overhang
                        If circuit.ConnectionSide = "left" Then
                            consys.InletHeaders.First.Overhangbottom = consys.InletHeaders.First.Xlist.Min + General.currentunit.TubeSheet.Dim_d + 90
                            consys.OutletHeaders.First.Overhangbottom = consys.OutletHeaders.First.Xlist.Min + General.currentunit.TubeSheet.Dim_d + 62
                        Else
                            consys.InletHeaders.First.Overhangtop = coil.FinnedDepth - consys.InletHeaders.First.Xlist.Max + General.currentunit.TubeSheet.Dim_d + 90
                            consys.OutletHeaders.First.Overhangtop = coil.FinnedDepth - consys.OutletHeaders.First.Xlist.Max + General.currentunit.TubeSheet.Dim_d + 62
                        End If
                    End If

                    CSData.GetBrineStutzen(consys, coil, circuit)

                    For Each sID In consys.InletHeaders.First.StutzenDatalist
                        slist.Add(sID.ID)
                    Next
                    For Each sID In consys.OutletHeaders.First.StutzenDatalist
                        slist.Add(sID.ID)
                    Next

                    'checkout Stutzen
                    CheckoutStutzen(slist)

                    'get SV position
                    CSData.SVPosition(consys, circuit)

                    'adjust header(s)
                    If circuit.NoDistributions > 1 Then
                        For i As Integer = 0 To consys.InletHeaders.Count - 1
                            SEPart.CreateHeaderTube(consys.InletHeaders(i), consys, circuit, coil)
                        Next
                        For i As Integer = 0 To consys.OutletHeaders.Count - 1
                            SEPart.CreateHeaderTube(consys.OutletHeaders(i), consys, circuit, coil)
                        Next

                        If consys.HeaderAlignment = "vertical" Then
                            'create nippletubes
                            For Each n In consys.InletNipples
                                SEPart.CreateNippleTube(consys.InletHeaders.First.Tube.Diameter, n, 10, consys, circuit.ConnectionSide, coil.FinnedDepth, circuit.NoPasses)
                            Next
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.InletNipples.First.Diameter,
                                                         .YPos = Math.Round(consys.InletHeaders.First.Dim_a + consys.InletHeaders.First.Tube.Diameter / 2, 2),
                                                         .ZPos = consys.InletHeaders.First.Origin(2) + consys.InletHeaders.First.Nipplepositions.First,
                                                         .Alignment = "vertical",
                                                         .Parentfile = consys.InletNipples.First.FileName})
                            For Each n In consys.OutletNipples
                                SEPart.CreateNippleTube(consys.OutletHeaders.First.Tube.Diameter, n, 10, consys, circuit.ConnectionSide, coil.FinnedDepth, circuit.NoPasses)
                            Next
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.OutletNipples.First.Diameter,
                                                         .YPos = Math.Round(consys.OutletHeaders.First.Dim_a + consys.OutletHeaders.First.Tube.Diameter / 2, 2),
                                                         .ZPos = consys.OutletHeaders.First.Origin(2) + consys.OutletHeaders.First.Nipplepositions.First,
                                                         .Alignment = "vertical",
                                                         .Parentfile = consys.OutletNipples.First.FileName})
                        Else
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.InletHeaders.First.Tube.Diameter,
                                                         .YPos = Math.Round(consys.InletHeaders.First.Dim_a + consys.InletHeaders.First.Tube.Diameter / 2, 2),
                                                         .ZPos = consys.InletHeaders.First.Origin(2),
                                                         .Alignment = "horizontal",
                                                         .Parentfile = consys.InletHeaders.First.Tube.FileName})
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.OutletHeaders.First.Tube.Diameter,
                                                        .YPos = Math.Round(consys.OutletHeaders.First.Dim_a + consys.OutletHeaders.First.Tube.Diameter / 2, 2),
                                                        .ZPos = consys.OutletHeaders.First.Origin(2),
                                                        .Alignment = "horizontal",
                                                        .Parentfile = consys.OutletHeaders.First.Tube.FileName})
                        End If
                    Else
                        consys.CoverSheetCutouts.Add(New NippleCutoutData With {.YPos = consys.InletHeaders.First.Dim_a, .ZPos = consys.InletHeaders.First.StutzenDatalist.First.ZPos})
                        consys.CoverSheetCutouts.Add(New NippleCutoutData With {.YPos = consys.OutletHeaders.First.Dim_a, .ZPos = consys.OutletHeaders.First.StutzenDatalist.First.ZPos})
                    End If

                    'add header and Stutzen to assembly
                    BuildConsysAsm(consys, coil, circuit)

                    'CS for GACV
                    If General.currentunit.ModelRangeName = "GACV" AndAlso General.isProdavit Then
                        Dim coversheet As String = GetCoversheet(circuit.ConnectionSide)
                        Dim position() As Double = GNData.CoverSheetOrigin(circuit.ConnectionSide)
                        General.currentunit.CoverSheet.MaterialCodeLetter = PCFData.GetValue("ConnectionCoveringTypeA", "MaterialCodeLetter")
                        If circuit.NoDistributions = 1 Then
                            For Each ncutout In consys.CoverSheetCutouts
                                ncutout.Filename = coversheet
                                ncutout.CutsizeY = 50
                                ncutout.CutsizeZ = 50
                                ncutout.YPos += position(0)
                                ncutout.ZPos = coil.FinnedHeight - ncutout.ZPos + position(1)
                            Next
                        Else
                            For Each ncutout In consys.CoverSheetCutouts
                                Dim mrsuffix As String = "FP"
                                If ncutout.Alignment = "vertical" And ncutout.Parentfile.Contains("Inlet") Then
                                    mrsuffix = "FX"
                                End If
                                Dim cutsize() As Double = GNData.CutoutSize(ncutout.Diameter, mrsuffix, False, 0)
                                ncutout.CutsizeY = cutsize(0)
                                ncutout.CutsizeZ = cutsize(1)
                                ncutout.Filename = coversheet
                                ncutout.YPos += position(0)
                                ncutout.ZPos = coil.FinnedHeight - ncutout.ZPos + position(1)
                            Next
                        End If
                    End If
                Else
                    'get stutzen IDs
                    If circuit.IsOnebranchEvap Then
                        If General.currentunit.UnitDescription = "Dual" Then
                            With consys.InletHeaders.First
                                Dim instutzen As New StutzenData With {.ID = "667013"}
                                instutzen.XPos = .Xlist.First
                                instutzen.YPos = -circuit.CoreTubeOverhang + 5
                                instutzen.ZPos = .Ylist.First
                                .StutzenDatalist.Add(instutzen)
                                slist.Add(instutzen.ID)
                            End With
                            With consys.OutletHeaders.First
                                Dim outstutzen As New StutzenData With {.ID = "660619"}
                                outstutzen.XPos = .Xlist.First
                                outstutzen.YPos = -circuit.CoreTubeOverhang + 5
                                outstutzen.ZPos = .Ylist.First
                                .StutzenDatalist.Add(outstutzen)
                                slist.Add(outstutzen.ID)
                            End With
                        Else
                            Dim singleID As String
                            singleID = GNData.GetSingleStutzen(circuit, consys, "inlet", coil, consys.InletHeaders.First.Dim_a, consys.InletHeaders.First.Xlist.First)
                            consys.InletHeaders.First.StutzenDatalist.Add(New StutzenData With {.ID = singleID, .XPos = consys.InletHeaders.First.Xlist.First, .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.InletHeaders.First.Ylist.First})
                            With consys.InletHeaders.First.Tube
                                .TubeType = "stutzen"
                                .HeaderType = "inlet"
                                .Materialcodeletter = consys.HeaderMaterial
                            End With

                            slist.Add(singleID)

                            singleID = GNData.GetSingleStutzen(circuit, consys, "outlet", coil, consys.OutletHeaders.First.Dim_a, consys.OutletHeaders.First.Xlist.First)
                            consys.OutletHeaders.First.StutzenDatalist.Add(New StutzenData With {.ID = singleID, .XPos = consys.OutletHeaders.First.Xlist.First, .YPos = -circuit.CoreTubeOverhang + 5, .ZPos = consys.OutletHeaders.First.Ylist.First})
                            consys.OutletHeaders.First.Tube.TubeType = "stutzen"
                            slist.Add(singleID)
                        End If

                        'checkout Stutzen
                        CheckoutStutzen(slist)

                        BuildConsysAsm(consys, coil, circuit)

                        If General.currentunit.ModelRangeName = "GACV" Then
                            Dim coversheet As String = GetCoversheet(circuit.ConnectionSide)
                            Dim position() As Double = GNData.CoverSheetOrigin(circuit.ConnectionSide)
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.YPos = consys.InletHeaders.First.Dim_a + position(0), .ZPos = coil.FinnedHeight - consys.InletHeaders.First.StutzenDatalist.First.ZPos + position(1),
                        .Alignment = "inlet", .CutsizeY = 50, .CutsizeZ = 50, .Filename = coversheet})
                            consys.CoverSheetCutouts.Add(New NippleCutoutData With {.YPos = consys.OutletHeaders.First.Dim_a + position(0), .ZPos = coil.FinnedHeight - consys.OutletHeaders.First.StutzenDatalist.First.ZPos + position(1),
                        .Alignment = "outlet", .CutsizeY = 50, .CutsizeZ = 50, .Filename = coversheet})

                            If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "P" Then
                                If circuit.FinType = "N" Or circuit.FinType = "M" Then
                                    consys.CoverSheetCutouts(0).Displacement = 110
                                ElseIf circuit.FinType Then
                                    If consys.HeaderMaterial = "C" Then
                                        consys.CoverSheetCutouts(0).Displacement = 58
                                        consys.CoverSheetCutouts(1).Displacement = -58
                                    Else
                                        consys.CoverSheetCutouts(0).Displacement = 53
                                        consys.CoverSheetCutouts(1).Displacement = -53
                                    End If
                                End If
                            Else
                                consys.CoverSheetCutouts.Remove(consys.CoverSheetCutouts(0))
                            End If
                        End If
                    ElseIf General.currentunit.UnitDescription = "Dual" Then
                        'own entire method for GADC / DHN



                    Else

                        'if header has to be splitted, add new header to the list and assign stutzen accordingly Or GNData.CheckDXSplit(consys.VType, circuit.NoPasses) > 1
                        If circuit.Pressure > 100 Then
                            If GNData.CheckCO2Split(consys.InletHeaders.First.Tube.Diameter, Calculation.DefaultHeaderLength(consys.InletHeaders.First, consys.HeaderAlignment)) Or GNData.CheckCO2Split(consys.OutletHeaders.First.Tube.Diameter, Calculation.DefaultHeaderLength(consys.OutletHeaders.First, consys.HeaderAlignment)) Then
                                co2split = True
                            End If
                        ElseIf GNData.CheckDXSplit(consys.VType, circuit.NoDistributions) > 1 Then
                            dxsplit = True
                        ElseIf ((General.currentunit.UnitDescription = "VShape" AndAlso consys.InletHeaders.First.Tube.Quantity = 2) Or (General.currentunit.ApplicationType = "Condenser" And GNData.CondenserLength(consys))) And circuit.Pressure > 16 Then
                            rdsplit = True
                        End If

                        If General.currentunit.ModelRangeName = "GACV" Then
                            'change overhang / displacement for certain headers
                            If dxsplit And consys.HeaderMaterial <> "C" And circuit.Pressure > 50 Then
                                If CSData.ControlDXGACV(consys.OutletHeaders.First, circuit) Then
                                    consys.SpecialCX = True
                                End If
                            ElseIf dxsplit And consys.HeaderMaterial = "V" And circuit.Pressure = 32 And circuit.FinType = "F" And consys.OutletHeaders.First.Tube.Diameter > 60 Then
                                CSData.ControlDXGACV(consys.OutletHeaders.First, circuit)
                                consys.SpecialRX = True
                            End If
                        End If

                        'standard stutzen → each row has the same for all positions
                        CSData.GetStutzen(consys, coil, circuit, 1)
                        If General.currentunit.UnitDescription = "Dual" Then
                            CSData.GetStutzen(consys, coil, circuit, 2)
                        End If

                        'check for special stutzen and replace the standard ones, limit to GACV and GxDV
                        If General.currentunit.ModelRangeName = "GACV" And consys.HeaderAlignment = "vertical" Then
                            'special stutzen for this one case
                            If (circuit.FinType = "N" Or circuit.FinType = "M") And circuit.Pressure = 80 And circuit.CoreTube.Materialcodeletter <> "V" Then 'materialcodeletter = "D"
                                CSData.GACVSpecialCXBottom(consys.OutletHeaders.First, circuit, consys)
                            End If

                            'only special stutzen when displacever < 0 and inlet F-FP lowest row for Ø60+
                            If consys.InletHeaders.First.Tube.Quantity > 0 Then
                                If consys.InletHeaders.First.Displacever < 0 Then
                                    If consys.HeaderMaterial <> "C" And circuit.Pressure > 16 Then
                                        'replace top inlet rows
                                        CSData.GACVSpecialInletN(consys.InletHeaders.First, circuit, consys)
                                    Else
                                        'replace top inlet row(s) and lowest inlet row
                                        CSData.GACVSpecialInlet(consys.InletHeaders.First, circuit, consys)
                                    End If
                                End If
                            End If
                            If consys.OutletHeaders.First.Tube.Quantity > 0 Then
                                If consys.OutletHeaders.First.Displacever < 0 Then
                                    If (circuit.FinType = "N" Or circuit.FinType = "M") And circuit.Pressure > 16 And consys.HeaderMaterial <> "C" Then
                                        'replace top outlet row1
                                        CSData.GACVSpecialOutletN(consys.OutletHeaders.First, circuit, consys)
                                    ElseIf circuit.FinType = "E" And consys.SpecialCX Then
                                        CSData.GACVSpecialCXE(consys.OutletHeaders.First, circuit, consys)
                                    ElseIf consys.SpecialRX Then
                                        CSData.GACVSpecialRXF(consys.OutletHeaders.First, circuit, consys)
                                    Else
                                        'replace top outlet row1
                                        CSData.GACVSpecialOutlet(consys.OutletHeaders.First, circuit, consys)
                                    End If
                                ElseIf General.currentunit.ModelRangeSuffix = "CP" And consys.OutletHeaders.First.Tube.Diameter = 26.9 Then
                                    'change 2nd row of outlet stutzen
                                    CSData.GACVSpecialCP269(consys.OutletHeaders.First, circuit, consys)
                                End If
                            End If
                        ElseIf General.currentunit.ModelRangeName.Substring(2, 2) = "DV" And General.currentunit.ApplicationType = "Condenser" Then
                            If circuit.Pressure <= 16 Then
                                'GFDV
                                If consys.OutletHeaders.First.Displacever < 0 Then
                                    CSData.GFDVSpecialOutlet(consys.OutletHeaders.First, circuit, consys)
                                End If
                            ElseIf circuit.Pressure > 100 Then
                                'GGDV
                                If consys.InletHeaders.First.Displacever <> 0 Then
                                    'bottom rows
                                    CSData.GGDVSpecialInlet(consys.InletHeaders.First, circuit, consys)
                                End If
                                If consys.OutletHeaders.First.Displacever <> 0 Then
                                    'top rows
                                    CSData.GGDVSpecialOutlet(consys.OutletHeaders.First, circuit, consys)
                                End If
                            ElseIf circuit.Pressure = 32 And circuit.CoreTube.Materialcodeletter <> "C" Then
                                If consys.OutletHeaders.First.Displacever < 0 Then
                                    CSData.GCDVNH3SpecialOutlet(consys.OutletHeaders.First, circuit, consys)
                                End If
                            End If
                        End If

                        If co2split Or dxsplit Or rdsplit Then
                            'not an entire new circuit! but the consys contains more than 1 header now
                            Calculation.SplitHeader(consys, co2split, rdsplit, circuit, coil)
                        End If

                        If rdsplit Then 'Or (General.currentunit.UnitDescription = "VShape" And consys.Circnumber > 1)
                            'all inlet headers have to be changed!
                            'change special stutzen for 2nd inlet header
                            'change origin and displacever
                            If consys.HeaderMaterial = "V" Or consys.HeaderMaterial = "W" Then
                                'GCDV AD
                                If circuit.NoPasses = 2 Or (circuit.NoPasses = 3 And coil.NoRows = 6) Then
                                    consys.InletHeaders.Last.Displacever = CSData.GCDVDisplacever(circuit, coil, consys.InletHeaders.Last.Tube.Diameter)
                                    If consys.InletHeaders.Last.Displacever > 0 Then
                                        consys.InletHeaders.Last.Origin(2) += consys.InletHeaders.Last.Displacever
                                        CSData.GCDVNH3SpecialInlet(consys.InletHeaders.Last, circuit, consys)
                                    End If
                                End If
                            Else
                                'GCDV RD 
                                If (circuit.NoPasses <= 3 And coil.NoRows = 4) Or (circuit.NoPasses <= 4 And coil.NoRows = 6) Then
                                    consys.InletHeaders.Last.Displacever = CSData.GCDVDisplacever(circuit, coil, consys.InletHeaders.Last.Tube.Diameter)
                                    If consys.InletHeaders.Last.Displacever > 0 Then
                                        consys.InletHeaders.Last.Origin(2) += consys.InletHeaders.Last.Displacever
                                        CSData.GCDVSpecialInlet(consys.InletHeaders.Last, circuit, consys)
                                    End If
                                End If
                            End If

                            For Each sID In consys.InletHeaders.Last.StutzenDatalist
                                slist.Add(sID.ID)
                            Next
                        End If

                        'second check, if circuiting is too small for entire coil - Blindtubes
                        If General.currentunit.UnitDescription = "VShape" And coil.FinnedHeight - circuit.CircuitSize(1) > coil.NoBlindTubeLayers * 2 * circuit.PitchY And circuit.Pressure < 17 Then
                            'currently only for GFDV working
                            If consys.Circnumber = 1 Then
                                'adjust inlet bottom row(s)
                                consys.InletHeaders.First.Overhangbottom = 15
                                consys.OutletHeaders.First.Overhangbottom = 15
                                consys.InletHeaders.First.Displacever = CSData.GFDVDisplacever(circuit, coil, consys.InletHeaders.First.Tube.Diameter, "inlet")
                                consys.InletHeaders.First.Origin(2) += 2.5 + consys.InletHeaders.First.Displacever
                                consys.OutletHeaders.First.Origin(2) += 2.5
                                If consys.InletHeaders.First.Displacever <> 0 Then
                                    If coil.NoRows = 6 Then
                                        CSData.GxDVMCInletRR6(consys.InletHeaders.First, circuit, consys)
                                    Else
                                        CSData.GxDVMCInletRR4(consys.InletHeaders.First, circuit, consys)
                                    End If
                                End If
                            Else
                                'adjust outlet top row(s) 
                                consys.InletHeaders.First.Overhangtop = 15
                                consys.OutletHeaders.First.Overhangtop = 15
                                consys.OutletHeaders.First.Displacever = CSData.GFDVDisplacever(circuit, coil, consys.OutletHeaders.First.Tube.Diameter, "outlet")
                                If consys.OutletHeaders.First.Displacever <> 0 Then
                                    If coil.NoRows = 6 Then
                                        CSData.GxDVMCOutletRR6(consys.OutletHeaders.First, circuit, consys)
                                    Else
                                        CSData.GxDVMCOutletRR4(consys.OutletHeaders.First, circuit, consys)
                                    End If
                                End If
                            End If
                        End If

                        For i As Integer = 0 To consys.InletHeaders.Count - 1
                            For Each sID In consys.InletHeaders(i).StutzenDatalist
                                slist.Add(sID.ID)
                            Next
                        Next
                        For i As Integer = 0 To consys.OutletHeaders.Count - 1
                            For Each sID In consys.OutletHeaders(i).StutzenDatalist
                                slist.Add(sID.ID)
                            Next
                        Next

                        'checkout Stutzen
                        CheckoutStutzen(slist)

                        CSData.SVPosition(consys, circuit)

                        'create header(s)
                        If consys.InletHeaders.First.Tube.Quantity > 0 Then
                            For i As Integer = 0 To consys.InletHeaders.Count - 1
                                SEPart.CreateHeaderTube(consys.InletHeaders(i), consys, circuit, coil)
                            Next
                            If consys.InletHeaders.First.Nipplepositions.Count > 0 Then
                                For Each n In consys.InletNipples
                                    If consys.InletHeaders.Count > 1 Then
                                        n.SVPosition = consys.InletNipples.First.SVPosition
                                        n.HeaderType = "inlet"
                                        n.TubeType = "nipple"
                                    End If
                                    SEPart.CreateNippleTube(consys.InletHeaders.First.Tube.Diameter, n, circuit.Pressure, consys, circuit.ConnectionSide, coil.FinnedDepth, circuit.NoPasses)
                                Next
                            End If
                        End If
                        For i As Integer = 0 To consys.OutletHeaders.Count - 1
                            SEPart.CreateHeaderTube(consys.OutletHeaders(i), consys, circuit, coil)
                        Next
                        If consys.OutletHeaders.First.Nipplepositions.Count > 0 Then
                            For Each n In consys.OutletNipples
                                If consys.OutletHeaders.Count > 1 Then
                                    n.SVPosition = consys.OutletNipples.First.SVPosition
                                    n.HeaderType = "outlet"
                                    n.TubeType = "nipple"
                                End If
                                SEPart.CreateNippleTube(consys.OutletHeaders.First.Tube.Diameter, n, circuit.Pressure, consys, circuit.ConnectionSide, coil.FinnedDepth, circuit.NoPasses)
                            Next
                        End If

                        'add header and Stutzen to assembly
                        BuildConsysAsm(consys, coil, circuit)

                        If General.currentunit.ModelRangeName = "GACV" AndAlso General.isProdavit Then
                            Dim coversheet As String = GetCoversheet(circuit.ConnectionSide)
                            Dim position() As Double = GNData.CoverSheetOrigin(circuit.ConnectionSide)
                            If consys.HeaderAlignment = "horizontal" Then
                                consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.InletHeaders.First.Tube.Diameter,
                                             .YPos = Math.Round(consys.InletHeaders.First.Dim_a + consys.InletHeaders.First.Tube.Diameter / 2, 2),
                                             .ZPos = consys.InletHeaders.First.Origin(2),
                                             .Alignment = "horizontal",
                                             .Parentfile = consys.InletHeaders.First.Tube.FileName,
                                             .Filename = coversheet})
                                Dim cutsize() As Double = GNData.CutoutSize(consys.CoverSheetCutouts.Last.Diameter, General.currentunit.ModelRangeSuffix, False, consys.ConType)
                                consys.CoverSheetCutouts.Last.CutsizeY = cutsize(0)
                                consys.CoverSheetCutouts.Last.CutsizeZ = cutsize(1)
                                consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.OutletHeaders.First.Tube.Diameter,
                                            .YPos = Math.Round(consys.OutletHeaders.First.Dim_a + consys.OutletHeaders.First.Tube.Diameter / 2, 2),
                                            .ZPos = consys.OutletHeaders.First.Origin(2),
                                            .Alignment = "horizontal",
                                            .Parentfile = consys.OutletHeaders.First.Tube.FileName,
                                            .Filename = coversheet})
                                cutsize = GNData.CutoutSize(consys.CoverSheetCutouts.Last.Diameter, General.currentunit.ModelRangeSuffix, False, consys.ConType)
                                consys.CoverSheetCutouts.Last.CutsizeY = cutsize(0)
                                consys.CoverSheetCutouts.Last.CutsizeZ = cutsize(1)
                            Else
                                For j As Integer = 0 To consys.InletHeaders.Count - 1
                                    For i As Integer = 0 To consys.InletHeaders(j).Nipplepositions.Count - 1
                                        consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.InletNipples.First.Diameter,
                                            .YPos = Math.Round(consys.InletHeaders(j).Dim_a + consys.InletHeaders(j).Tube.Diameter / 2, 2),
                                            .ZPos = consys.InletHeaders(j).Origin(2) + consys.InletHeaders(j).Nipplepositions(i),
                                            .Alignment = "vertical",
                                            .Parentfile = consys.InletHeaders(j).Tube.FileName,
                                            .HasFlange = consys.HasFTCon,
                                            .Filename = coversheet})
                                        Dim hasSV As Boolean = False
                                        If consys.ConType = 1 Then
                                            consys.CoverSheetCutouts.Last.HasFlange = False
                                        End If
                                        If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or (circuit.Pressure < 17 And Not consys.HasFTCon) Then
                                            hasSV = True
                                        End If
                                        Dim cutsize() As Double = GNData.CutoutSize(consys.CoverSheetCutouts.Last.Diameter, General.currentunit.ModelRangeSuffix, hasSV, consys.ConType)
                                        consys.CoverSheetCutouts.Last.CutsizeY = cutsize(0)
                                        consys.CoverSheetCutouts.Last.CutsizeZ = cutsize(1)
                                        If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or (consys.CoverSheetCutouts.Last.Diameter > 54 And consys.InletNipples.First.Materialcodeletter = "C" And General.currentunit.ModelRangeSuffix = "CP") Then
                                            consys.CoverSheetCutouts.Last.Displacement = -12.5
                                        End If
                                    Next
                                Next

                                For j As Integer = 0 To consys.OutletHeaders.Count - 1
                                    For i As Integer = 0 To consys.OutletHeaders.First.Nipplepositions.Count - 1
                                        consys.CoverSheetCutouts.Add(New NippleCutoutData With {.Diameter = consys.OutletNipples.First.Diameter,
                                            .YPos = Math.Round(consys.OutletHeaders(j).Dim_a + consys.OutletHeaders(j).Tube.Diameter / 2, 2),
                                            .ZPos = consys.OutletHeaders(j).Origin(2) + consys.OutletHeaders(j).Nipplepositions(i),
                                            .Alignment = "vertical",
                                            .Parentfile = consys.OutletHeaders(j).Tube.FileName,
                                            .HasFlange = consys.HasFTCon,
                                            .Filename = coversheet})
                                        Dim hasSV As Boolean = False
                                        If consys.ConType = 1 Then
                                            consys.CoverSheetCutouts.Last.HasFlange = False
                                        End If
                                        If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or (circuit.Pressure < 17 And Not consys.HasFTCon) Then
                                            hasSV = True
                                        End If
                                        Dim cutsize() As Double = GNData.CutoutSize(consys.CoverSheetCutouts.Last.Diameter, General.currentunit.ModelRangeSuffix, hasSV, consys.ConType)
                                        consys.CoverSheetCutouts.Last.CutsizeY = cutsize(0)
                                        consys.CoverSheetCutouts.Last.CutsizeZ = cutsize(1)
                                        If General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or (consys.CoverSheetCutouts.Last.Diameter > 54 And consys.InletNipples.First.Materialcodeletter = "C" And General.currentunit.ModelRangeSuffix = "CP") Then
                                            consys.CoverSheetCutouts.Last.Displacement = -12.5
                                        End If
                                    Next
                                Next
                            End If
                            For Each ncutout In consys.CoverSheetCutouts
                                ncutout.YPos += position(0)
                                ncutout.ZPos = coil.FinnedHeight - ncutout.ZPos + position(1)
                            Next
                        End If
                    End If
                End If

                'create the coversheet for GACV
                If General.currentunit.ModelRangeName = "GACV" AndAlso General.isProdavit Then
                    General.seapp.Documents.Open(consys.CoverSheetCutouts.First.Filename)
                    General.seapp.DoIdle()
                    For i As Integer = 0 To consys.CoverSheetCutouts.Count - 1
                        SEPart.CreateCutout(consys.CoverSheetCutouts(i), General.seapp.ActiveDocument, circuit.ConnectionSide)
                    Next

                    'refresh the flatten model
                    General.seapp.StartCommand(45066)
                    General.seapp.DoIdle()
                    Threading.Thread.Sleep(1000)
                    General.seapp.StartCommand(10768)
                    'change the file properties
                    SEPart.CoversheetProps("part")
                    General.seapp.Documents.CloseDocument(consys.CoverSheetCutouts.First.Filename, SaveChanges:=True, DoIdle:=True)
                    If circuit.Pressure <= 16 Then
                        Dim asmpsm As String = General.GetFullFilename(General.currentjob.Workspace, "\000", "asm")
                        If asmpsm <> "" Then
                            IO.File.Delete(asmpsm)
                        End If
                    End If
                End If
            Else

                asmdoc = General.seapp.Documents.Open(consys.ConSysFile.Fullfilename)
                General.seapp.DoIdle()

                SEAsm.WriteCustomProps(asmdoc, consys.ConSysFile)
                General.seapp.Documents.CloseDocument(consys.ConSysFile.Fullfilename, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub GetDefrostArrangement(ByRef coil As CoilData, conside As String, dftdoc As SED.DraftDocument)
        Dim objsheet As SED.Sheet

        Try
            'check if mirrored
            If conside = "left" Then
                'move all objects by frame width in x direction
                SEDraft.CreateMirroredDFT(dftdoc, False, conside, coil.Alignment, False)
            End If

            'Find the sheet of the main circuiting view
            objsheet = SEDraft.FindSheet(dftdoc)

            If objsheet IsNot Nothing Then
                heatingcoords = SEDraft.HeatingPos(objsheet)
                heatingbows = SEDraft.HeatingBows(objsheet, heatingcoords)
                General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=False, DoIdle:=True)
                coil.Defrostbows = heatingbows
                'adjust fin
                SEPart.AdjustFinEDefrost(heatingcoords, General.currentjob.Workspace + "\CoilFin11.par")
            End If

        Catch ex As Exception

        End Try

    End Sub

    Shared Function GetTubeArrangement(ByRef coil As CoilData, ByRef circuit As CircuitData, circfile As String, circtype As String) As List(Of Double)()()
        Dim dftdoc As SED.DraftDocument
        Dim bowlines(), frontlines, backlines As List(Of SolidEdgeFrameworkSupport.Line2d)
        Dim backbowprops(), frontbowprops(), frontbowprops2(), backbowprops2(), inoutlist(), inoutlist2() As List(Of Double)
        Dim objsheet, mainsheet, objsheet2 As SED.Sheet
        Dim switchcoords As Boolean = False
        Dim changefin As Boolean = False
        Dim circsize, circframe(1), circfinpitch() As Double
        Dim circfin, ADpos1 As String
        Dim p1list, p2list, checklist As New List(Of PointData)
        Dim uniquepoint1, uniquepoint2 As PointData
        Dim xlist, ylist As New List(Of Double)

        Try

            'check if mirrored
            dftdoc = SEDraft.OpenDFT(circfile)
            If Not circuit.CustomCirc Then
                If Not General.currentunit.UnitDescription = "VShape" And Not General.currentunit.UnitDescription = "Dual" And
                    ((circuit.ConnectionSide = "left" And Not circtype.Contains("defrost")) Or (circuit.ConnectionSide = "right" And circtype.Contains("defrost"))) Then
                    'move all objects by frame width in x direction
                    SEDraft.CreateMirroredDFT(dftdoc, False, circuit.ConnectionSide, coil.Alignment, circuit.CustomCirc)
                End If

                If General.currentunit.UnitDescription = "VShape" And circuit.Quantity = 1 And Not circuit.CircuitType.Contains("Subcooler") Then
                    'search correct drawing view - B = right side // A = left side
                    objsheet = If(SEDraft.FindSheetX(circuit.ConnectionSide, circuit.PitchX, circuit.PitchY), SEDraft.FindSheet(dftdoc))
                Else
                    'Find the sheet of the main circuiting view
                    objsheet = SEDraft.FindSheet(dftdoc)
                    If General.currentunit.UnitDescription = "VShape" And circuit.Quantity > 1 Then
                        'check position of first coretube (must be top right)
                        SEDraft.VShapeMirrorDV(objsheet, circuit.ConnectionSide)
                    End If
                End If

                'check for orientation of coil in circuiting → if =/= coil position, then switch coords
                'check also for sandwich design and a split ciruiting

                If General.currentunit.ApplicationType = "Condenser" And circuit.FinType <> "N" And circuit.FinType <> "M" Then
                    circframe = SEDraft.GetCoilFrame(objsheet, coil.Alignment, General.currentunit.MultiCircuitDesign, coil.Number)

                    'check, if fin matches the ct arrangement from the circuiting
                    circfin = CircProps.CheckFin(coil.Alignment, circuit.FinType, objsheet)
                    circfinpitch = GNData.GetFinPitch(coil.Alignment, circfin)

                    If circfinpitch(0) <> circuit.PitchX Or circfinpitch(1) <> circuit.PitchY Then
                        circuit.PitchX = circfinpitch(0)
                        circuit.PitchY = circfinpitch(1)
                        changefin = True

                        'Get new frame
                        circframe = CircProps.ChangeFrame(circframe, circfinpitch, GNData.GetFinPitch(coil.Alignment, circuit.FinType))
                    End If

                    switchcoords = Calculation.GetCircsize(circuit, coil.Alignment, circframe, objsheet)
                    If switchcoords Then
                        If coil.Alignment = "horizontal" Then
                            circsize = circframe(0)
                        Else
                            circsize = circframe(1)
                        End If
                    End If
                End If
            Else
                objsheet = SEDraft.FindSheet(dftdoc)
            End If

            If objsheet IsNot Nothing Then
                'if passno = 1 → no bows
                If circuit.NoPasses > 1 Then
                    bowlines = SEDraft.GetBowLines(objsheet, coil.Alignment, General.currentunit.MultiCircuitDesign, coil.Number)
                    If circtype.Contains("defrost") Then
                        backlines = bowlines(0)
                        frontlines = bowlines(1)
                    Else
                        frontlines = bowlines(0)
                        backlines = bowlines(1)
                    End If

                    'Get properties for front bows (position, length, type, default level)
                    If frontlines.Count > 0 Then
                        frontbowprops = SEDraft.GetBowProps(objsheet, frontlines, circuit.PitchX, circuit.PitchY, coil.Alignment, "front", circuit.CircuitType.ToLower,
                                                            circuit.NoPasses, coil.EDefrostPDMID, General.currentunit.MultiCircuitDesign, coil.Number)
                        'Get the level 
                        frontbowprops = CircProps.GetBowLevels(frontbowprops, circtype, "front", circuit.FinType)
                        If circtype = "defrost" Then
                            frontbowprops = TransformCoords(frontbowprops, coil.FinnedDepth)
                            frontbowprops = CircProps.GetBrineLevels(frontbowprops, coolingfrontbowprops, "front", coil.Circuits(0).FinType)
                        Else
                            If switchcoords Then
                                frontbowprops = SwitchCoordLists(frontbowprops, circsize, circuit.ConnectionSide)
                            End If
                            coolingfrontbowprops = frontbowprops
                        End If
                    End If

                    If backlines.Count > 0 Then
                        'Get properties for back bows
                        backbowprops = SEDraft.GetBowProps(objsheet, backlines, circuit.PitchX, circuit.PitchY, coil.Alignment, "back", circuit.CircuitType.ToLower,
                                                           circuit.NoPasses, coil.EDefrostPDMID, General.currentunit.MultiCircuitDesign, coil.Number)
                        'Get the level
                        backbowprops = CircProps.GetBowLevels(backbowprops, circtype, "back", circuit.FinType)
                        If circtype.Contains("defrost") Then
                            backbowprops = TransformCoords(backbowprops, coil.FinnedDepth)
                            backbowprops = CircProps.GetBrineLevels(backbowprops, coolingbackbowprops, "back", coil.Circuits(0).FinType)
                        Else
                            If switchcoords Then
                                backbowprops = SwitchCoordLists(backbowprops, circsize, circuit.ConnectionSide)
                            End If
                            coolingbackbowprops = backbowprops
                        End If
                    End If
                    If changefin Then
                        If frontlines.Count > 0 Then
                            coolingfrontbowprops = CircProps.ChangeBowprops(coolingfrontbowprops, circfinpitch, GNData.GetFinPitch(coil.Alignment, circuit.FinType))
                        End If
                        coolingbackbowprops = CircProps.ChangeBowprops(coolingbackbowprops, circfinpitch, GNData.GetFinPitch(coil.Alignment, circuit.FinType))
                        circfinpitch = GNData.GetFinPitch(coil.Alignment, circuit.FinType)
                        circuit.PitchX = circfinpitch(0)
                        circuit.PitchY = circfinpitch(1)
                    End If

                    If General.currentunit.UnitDescription = "Dual" Then
                        mainsheet = dftdoc.Sheets.Item(1)
                        mainsheet.Activate()
                        General.seapp.DoIdle()

                        ADpos1 = SEDraft.GetADPosition(objsheet)
                        If ADpos1 = "left" Then
                            If frontlines.Count > 0 Then
                                For i As Integer = 0 To frontbowprops(0).Count - 1
                                    frontbowprops(0)(i) = Math.Round(frontbowprops(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                                    frontbowprops(2)(i) = Math.Round(frontbowprops(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                                Next
                            End If
                            If backlines.Count > 0 Then
                                For i As Integer = 0 To backbowprops(0).Count - 1
                                    backbowprops(0)(i) = Math.Round(backbowprops(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                                    backbowprops(2)(i) = Math.Round(backbowprops(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                                Next
                            End If
                        End If

                        'find second objsheet for other coil // if only 1 DV, then create mirror of first
                        objsheet2 = SEDraft.FindSheet2(circfile, objsheet.Name)
                        If objsheet2 IsNot Nothing Then
                            Dim ADpos2 As String = SEDraft.GetADPosition(objsheet2)
                            If ADpos1 = ADpos2 Then
                                'add same bowprops to existing front and back, but this time in mirrored
                                If frontlines.Count > 0 Then
                                    frontbowprops2 = Calculation.MirrorCoordsFromList(frontbowprops, coil.FinnedDepth, New List(Of Integer) From {0, 2})
                                End If
                                If backlines.Count > 0 Then
                                    backbowprops2 = Calculation.MirrorCoordsFromList(backbowprops, coil.FinnedDepth, New List(Of Integer) From {0, 2})
                                End If
                            Else
                                bowlines = SEDraft.GetBowLines(objsheet2, coil.Alignment, General.currentunit.MultiCircuitDesign, coil.Number)
                                If circtype.Contains("defrost") Then
                                    backlines = bowlines(0)
                                    frontlines = bowlines(1)
                                Else
                                    frontlines = bowlines(0)
                                    backlines = bowlines(1)
                                End If
                                If frontlines.Count > 0 Then
                                    frontbowprops2 = SEDraft.GetBowProps(objsheet2, frontlines, circuit.PitchX, circuit.PitchY, coil.Alignment, "front", circuit.CircuitType.ToLower,
                                                                    circuit.NoPasses, coil.EDefrostPDMID, General.currentunit.MultiCircuitDesign, coil.Number)
                                    'Get the level 
                                    frontbowprops2 = CircProps.GetBowLevels(frontbowprops2, circtype, "front", circuit.FinType)
                                End If
                                If backlines.Count > 0 Then
                                    'Get properties for back bows
                                    backbowprops2 = SEDraft.GetBowProps(objsheet2, backlines, circuit.PitchX, circuit.PitchY, coil.Alignment, "back", circuit.CircuitType.ToLower,
                                                                   circuit.NoPasses, coil.EDefrostPDMID, General.currentunit.MultiCircuitDesign, coil.Number)
                                    'Get the level
                                    backbowprops2 = CircProps.GetBowLevels(backbowprops2, circtype, "back", circuit.FinType)
                                End If
                            End If

                            'check AD location // if right, then add gap + fd value to all front and back x values (0,2)
                            If ADpos1 = "right" Then
                                If frontlines.Count > 0 Then
                                    For i As Integer = 0 To frontbowprops2(0).Count - 1
                                        frontbowprops2(0)(i) = Math.Round(frontbowprops2(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                                        frontbowprops2(2)(i) = Math.Round(frontbowprops2(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                                    Next
                                End If
                                If backlines.Count > 0 Then
                                    For i As Integer = 0 To backbowprops2(0).Count - 1
                                        backbowprops2(0)(i) = Math.Round(backbowprops2(0)(i) + coil.Gap + coil.FinnedDepth, 3)
                                        backbowprops2(2)(i) = Math.Round(backbowprops2(2)(i) + coil.Gap + coil.FinnedDepth, 3)
                                    Next
                                End If
                            End If

                            If circuit.NoDistributions = 1 Then
                                'check if 1 for entire unit → circuit DV would contain only 1 group 
                                If SEDraft.CountStrandGroups(objsheet2) = 1 Then
                                    circuit.IsOnebranchEvap = True
                                    'add a bow backbow (right), frontbow (left) - identify start and end?
                                    'find in / out in each sheet
                                    inoutlist = SEDraft.GetInOutCoords(objsheet, circuit.PitchX, circuit.PitchY, coil.Alignment, "cooling", circuit.NoPasses, coil.Number)
                                    If inoutlist(0).Count = 0 Then
                                        inoutlist(0).Add(inoutlist(2)(0))
                                        inoutlist(1).Add(inoutlist(3)(0))
                                    ElseIf inoutlist(2).Count = 0 Then
                                        inoutlist(2).Add(inoutlist(0)(0))
                                        inoutlist(3).Add(inoutlist(1)(0))
                                    End If
                                    If ADpos1 = "left" Then
                                        inoutlist(0)(0) = Math.Round(inoutlist(0)(0) + coil.Gap + coil.FinnedDepth, 3)
                                        inoutlist(2)(0) = Math.Round(inoutlist(2)(0) + coil.Gap + coil.FinnedDepth, 3)
                                    End If

                                    inoutlist2 = SEDraft.GetInOutCoords(objsheet2, circuit.PitchX, circuit.PitchY, coil.Alignment, "cooling", circuit.NoPasses, coil.Number)
                                    If inoutlist2(0).Count = 0 Then
                                        inoutlist2(0).Add(inoutlist2(2)(0))
                                        inoutlist2(1).Add(inoutlist2(3)(0))
                                    ElseIf inoutlist2(2).Count = 0 Then
                                        inoutlist2(2).Add(inoutlist2(0)(0))
                                        inoutlist2(3).Add(inoutlist2(1)(0))
                                    End If
                                    If ADpos1 = "right" Then
                                        inoutlist2(0)(0) = Math.Round(inoutlist2(0)(0) + coil.Gap + coil.FinnedDepth, 3)
                                        inoutlist2(2)(0) = Math.Round(inoutlist2(2)(0) + coil.Gap + coil.FinnedDepth, 3)
                                    End If

                                    For i As Integer = 0 To coil.NoRows - 1
                                        xlist.Add(Math.Round(circuit.PitchX / 2 + i * circuit.PitchX, 3))
                                        xlist.Add(Math.Round(circuit.PitchX / 2 + i * circuit.PitchX + coil.Gap + coil.FinnedDepth, 3))
                                    Next
                                    For i As Integer = 0 To coil.NoLayers - 1
                                        ylist.Add(Math.Round(circuit.PitchY / 2 + i * circuit.PitchY, 3))
                                    Next
                                    For i As Integer = 0 To xlist.Count - 1
                                        For Each y In ylist
                                            checklist.Add(New PointData With {.X = xlist(i), .Y = y})
                                        Next
                                    Next

                                    'convert lists into a list of points
                                    p1list = Calculation.ConvertToPointlist(New List(Of List(Of Double)()) From {backbowprops, frontbowprops, inoutlist})
                                    p2list = Calculation.ConvertToPointlist(New List(Of List(Of Double)()) From {backbowprops2, frontbowprops2, inoutlist2})

                                    For Each p As PointData In checklist
                                        uniquepoint1 = Calculation.GetUniquePoint(checklist, p1list)
                                        uniquepoint2 = Calculation.GetUniquePoint(checklist, p2list)
                                    Next

                                    Dim bowlength As Double = Math.Round(Math.Sqrt((uniquepoint2.Y - uniquepoint1.Y) ^ 2 + (uniquepoint2.X - uniquepoint1.X) ^ 2))

                                    'check, if front or back → uniquepoint1 in frontbow or backbow lists
                                    If Calculation.ConvertToPointlist(New List(Of List(Of Double)()) From {frontbowprops}).FindAll(Function(c) c.X.Equals(uniquepoint1.X) And c.Y.Equals(uniquepoint1.Y)).Count = 1 Then
                                        'in frontlist → must be backbow
                                        backbowprops(0).Add(uniquepoint1.X)
                                        backbowprops(1).Add(uniquepoint1.Y)
                                        backbowprops(2).Add(uniquepoint2.X)
                                        backbowprops(3).Add(uniquepoint2.Y)
                                        backbowprops(4).Add(bowlength)
                                        backbowprops(5).Add(1)
                                        backbowprops(6).Add(1)
                                    ElseIf Calculation.ConvertToPointlist(New List(Of List(Of Double)()) From {backbowprops}).FindAll(Function(c) c.X.Equals(uniquepoint1.X) And c.Y.Equals(uniquepoint1.Y)).Count = 1 Then
                                        'in backlist → must be frontbow
                                        frontbowprops(0).Add(uniquepoint1.X)
                                        frontbowprops(1).Add(uniquepoint1.Y)
                                        frontbowprops(2).Add(uniquepoint2.X)
                                        frontbowprops(3).Add(uniquepoint2.Y)
                                        frontbowprops(4).Add(bowlength)
                                        frontbowprops(5).Add(1)
                                        frontbowprops(6).Add(1)
                                    End If
                                End If
                            End If
                        Else
                            'add same bowprops to existing front and back, but this time in mirrored
                            If frontlines.Count > 0 Then
                                frontbowprops2 = Calculation.MirrorCoordsFromList(frontbowprops, coil.FinnedDepth, New List(Of Integer) From {0, 2})
                            End If
                            If backlines.Count > 0 Then
                                backbowprops2 = Calculation.MirrorCoordsFromList(backbowprops, coil.FinnedDepth, New List(Of Integer) From {0, 2})
                            End If
                        End If

                        For i As Integer = 0 To frontbowprops2.Count - 1
                            frontbowprops(i).AddRange(frontbowprops2(i).ToArray)
                        Next
                        For i As Integer = 0 To backbowprops2.Count - 1
                            backbowprops(i).AddRange(backbowprops2(i).ToArray)
                        Next
                        coolingfrontbowprops = frontbowprops
                        coolingbackbowprops = backbowprops
                    End If
                End If

                If circuit.FinType = "G" Then
                    coil.SupportTubesPosition = SEDraft.GFinSupport(objsheet, coil)
                End If
            End If

            rotatefin = switchcoords

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return {frontbowprops, backbowprops}
    End Function

    Shared Sub CheckoutList(itemlist As List(Of String), workspace As String)
        For Each singleid In itemlist
            WSM.CheckoutPart(singleid, General.currentjob.OrderDir, workspace, General.batfile)
            If singleid.Contains(Library.TemplateParts.BOW1) Then
                WSM.CheckoutCircs(Library.TemplateParts.BOW1, General.currentjob.OrderDir, workspace, General.batfile)
                General.WaitForFile(General.currentjob.Workspace, Library.TemplateParts.BOW1, "par", 30)
            ElseIf singleid.Contains(Library.TemplateParts.BOW9) Then
                WSM.CheckoutCircs(Library.TemplateParts.BOW9, General.currentjob.OrderDir, workspace, General.batfile)
                General.WaitForFile(General.currentjob.Workspace, Library.TemplateParts.BOW9, "par", 30)
            End If
        Next
    End Sub

    Shared Sub BuildCoilAsm(ByRef circuit As CircuitData, ByRef frontbowprops() As List(Of Double), ByRef backbowprops() As List(Of Double), frontbowids As List(Of String), backbowids As List(Of String), ByRef coil As CoilData)
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim offset() As Double

        Try
            asmdoc = General.seapp.Documents.Open(coil.CoilFile.Fullfilename)
            General.seapp.DoIdle()

            circuit.CoreTubes.Add(New CoreTubeLocations)
            SEAsm.FillCoretubePositions(asmdoc, circuit.CoreTube.FileName, circuit.CoreTubes.Last)

            'consider blind tubes
            If coil.NoBlindTubes > 0 Then
                'no of blind tube layers at the beginning of the circuit
                coil.NoBlindTubeLayers = CircProps.CheckBTLayer(coil.NoBlindTubes, coil.FinnedDepth, {circuit.PitchX, circuit.PitchY}, coil.Alignment)
            End If

            offset = CircProps.GetCircOffset(circuit, coil)

            If rotatefin And coil.Alignment = "vertical" Then
                'only offset(1) relevant
                'new coord = finnedh - old coord - offset(1)
                If frontbowids.Count > 0 Then
                    frontbowprops(0) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, frontbowprops(0), 0)
                    frontbowprops(1) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, frontbowprops(1), offset(1))
                    frontbowprops(2) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, frontbowprops(2), 0)
                    frontbowprops(3) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, frontbowprops(3), offset(1))
                End If
                If backbowids.Count > 0 Then
                    backbowprops(0) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, backbowprops(0), 0)
                    backbowprops(1) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, backbowprops(1), offset(1))
                    backbowprops(2) = CircProps.MoveCoordsbyRotation(coil.FinnedDepth, backbowprops(2), 0)
                    backbowprops(3) = CircProps.MoveCoordsbyRotation(coil.FinnedHeight, backbowprops(3), offset(1))
                End If
                'has to consider blindtubes
                offset(1) = coil.NoBlindTubeLayers * 2 * circuit.PitchY
            End If

            If circuit.NoPasses > 1 Then
                If frontbowids.Count > 0 Then
                    'place bows front
                    SEAsm.PlaceBows(asmdoc, frontbowprops, frontbowids, "front", General.currentjob.Workspace, circuit.FinType, circuit.CoreTube.FileName, circuit.CoreTubeOverhang, offset, circuit, coil)
                End If
                If backbowids.Count > 0 Then
                    SEAsm.PlaceBows(asmdoc, backbowprops, backbowids, "back", General.currentjob.Workspace, circuit.FinType, circuit.CoreTube.FileName, circuit.CoreTubeOverhang, offset, circuit, coil)
                End If
            End If

            Dim blankconfig As String = "BlankC"
            If circuit.CircuitType.Contains("Defrost") Then
                blankconfig = "BlankS"
            End If

            If frontbowids.Count > 0 And backbowids.Count > 0 Then
                'SEAsm.CreateBowLevelViews(asmdoc, frontbowids, backbowids, frontbowprops(frontbowprops.Count - 2), backbowprops(backbowprops.Count - 2), coil.Occlist, blankconfig)
                circuit.Frontbowids.AddRange(frontbowids.ToArray)
                circuit.Backbowids.AddRange(backbowids.ToArray)
            ElseIf frontbowids.Count = 0 Then
                'SEAsm.CreateBowLevelViews(asmdoc, New List(Of String), backbowids, New List(Of Double) From {0}, backbowprops(backbowprops.Count - 2), coil.Occlist, blankconfig)
                circuit.Backbowids.AddRange(backbowids.ToArray)
            Else
                'SEAsm.CreateBowLevelViews(asmdoc, frontbowids, New List(Of String), frontbowprops(frontbowprops.Count - 2), New List(Of Double) From {0}, coil.Occlist, blankconfig)
                circuit.Frontbowids.AddRange(frontbowids.ToArray)
            End If

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function SwitchCoordLists(ByRef bowprops() As List(Of Double), framesize As Double, conside As String) As List(Of Double)()
        Dim tempy1list, tempy2list As List(Of Double)
        Dim tempx1list, tempx2list As New List(Of Double)

        tempy1list = bowprops(0)
        tempy2list = bowprops(2)

        For i As Integer = 0 To bowprops(0).Count - 1
            tempx1list.Add(framesize - bowprops(1)(i))
            tempx2list.Add(framesize - bowprops(3)(i))
        Next

        bowprops(0) = tempx1list
        bowprops(1) = tempy1list
        bowprops(2) = tempx2list
        bowprops(3) = tempy2list

        Return bowprops
    End Function

    Shared Sub NewBowKeys(ByRef frontbowprops() As List(Of Double), ByRef backbowprops() As List(Of Double), ByRef frontbowids As List(Of String), ByRef backbowids As List(Of String), circuit As CircuitData)
        Dim bowkey, bowid As String
        Dim index As Integer

        If frontbowids.Count > 0 Then
            For i As Integer = 0 To frontbowids.Count - 1
                If frontbowids(i).Contains(Library.TemplateParts.BOW1) Or frontbowids(i).Contains(Library.TemplateParts.BOW9) Or frontbowids(i).Contains(Library.TemplateParts.ORIBITALBOW1) Then
                    'create key
                    bowkey = frontbowprops(7)(i).ToString + "\" + frontbowprops(4)(i).ToString + "\" + frontbowprops(5)(i).ToString
                    If bowkeys.Count > 0 Then
                        index = bowkeys.IndexOf(bowkey)
                        If index >= 0 Then
                            'Get bowid
                            bowid = templateids(index)
                            frontbowids(i) = bowid
                        Else
                            bowkeys.Add(bowkey)
                            'Create new bow.par and return bowid + put into list
                            bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, frontbowids(i), circuit.FinType, circuit.CoreTube.Material, circuit.Pressure, False)
                            frontbowids(i) = bowid
                            templateids.Add(bowid)
                        End If
                    Else
                        bowkeys.Add(bowkey)
                        'Create new bow.par and return bowid + put into list
                        bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, frontbowids(i), circuit.FinType, circuit.CoreTube.Material, circuit.Pressure, False)
                        frontbowids(i) = bowid
                        templateids.Add(bowid)
                    End If
                End If
            Next
        End If

        If backbowids.Count > 0 Then
            For i As Integer = 0 To backbowids.Count - 1
                If backbowids(i).Contains(Library.TemplateParts.BOW1) Or backbowids(i).Contains(Library.TemplateParts.BOW9) Or backbowids(i).Contains(Library.TemplateParts.ORIBITALBOW1) Then
                    'Create key
                    bowkey = backbowprops(7)(i).ToString + "\" + backbowprops(4)(i).ToString + "\" + backbowprops(5)(i).ToString
                    'Check if key already in list
                    If bowkeys.Count > 0 Then
                        index = bowkeys.IndexOf(bowkey)
                        If index >= 0 Then
                            'Get bowid
                            bowid = templateids(index)
                            backbowids(i) = bowid
                        Else
                            bowkeys.Add(bowkey)
                            'Create new bow.par and return bowid + put into list
                            bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, backbowids(i), circuit.FinType, circuit.CoreTube.Material, circuit.Pressure, False)
                            backbowids(i) = bowid
                            templateids.Add(bowid)
                        End If
                    Else
                        bowkeys.Add(bowkey)
                        'Create new bow.par and return bowid + put into list
                        bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, backbowids(i), circuit.FinType, circuit.CoreTube.Material, circuit.Pressure, False)
                        backbowids(i) = bowid
                        templateids.Add(bowid)
                    End If
                End If
            Next
        End If

    End Sub

    Shared Sub NewBrineBowKeys(ByRef frontbowprops() As List(Of Double), ByRef backbowprops() As List(Of Double), ByRef frontbowids As List(Of String), ByRef backbowids As List(Of String), circuit As CircuitData)
        Dim bowkey, bowid As String
        Dim bowkeys, templateids As New List(Of String)
        Dim index As Integer

        For i As Integer = 0 To frontbowids.Count - 1
            If frontbowids(i).Contains(Library.TemplateParts.BOW1) Or frontbowids(i).Contains(Library.TemplateParts.BOW9) Then

                'identify the correct level
                If frontbowprops(7)(i) > 1 And frontbowprops(6)(i) > 1 And circuit.CoreTube.Materialcodeletter = "C" Then
                    frontbowprops(8)(i) += 5
                End If

                bowkey = frontbowprops(8)(i).ToString + "\" + frontbowprops(4)(i).ToString + "\" + frontbowprops(5)(i).ToString + "\c"

                'Check if key already in list
                If bowkeys.Count > 0 Then
                    index = bowkeys.IndexOf(bowkey)
                    If index >= 0 Then
                        'Get bowid
                        bowid = templateids(index)
                        frontbowids(i) = bowid
                    Else
                        bowkeys.Add(bowkey)
                        'Create new bow.par and return bowid + put into list
                        bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, frontbowids(i), "F", circuit.CoreTube.Material, 10, True)
                        frontbowids(i) = bowid
                        templateids.Add(bowid)
                    End If
                Else
                    bowkeys.Add(bowkey)
                    'Create new bow.par and return bowid + put into list
                    bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, frontbowids(i), "F", circuit.CoreTube.Material, 10, True)
                    frontbowids(i) = bowid
                    templateids.Add(bowid)
                End If
            End If
        Next

        If backbowids.Count > 0 Then
            For i As Integer = 0 To backbowids.Count - 1
                If backbowids(i).Contains(Library.TemplateParts.BOW1) Or backbowids(i).Contains(Library.TemplateParts.BOW9) Then

                    If backbowprops(7)(i) > 1 And backbowprops(6)(i) > 1 Then
                        backbowprops(8)(i) += 5
                    End If

                    bowkey = backbowprops(8)(i).ToString + "\" + backbowprops(4)(i).ToString + "\" + backbowprops(5)(i).ToString + "\c"

                    'Check if key already in list
                    If bowkeys.Count > 0 Then
                        index = bowkeys.IndexOf(bowkey)
                        If index >= 0 Then
                            'Get bowid
                            bowid = templateids(index)
                            backbowids(i) = bowid
                        Else
                            bowkeys.Add(bowkey)
                            'Create new bow.par and return bowid + put into list
                            bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, backbowids(i), "F", circuit.CoreTube.Material, 10, True)
                            backbowids(i) = bowid
                            templateids.Add(bowid)
                        End If
                    Else
                        bowkeys.Add(bowkey)
                        'Create new bow.par and return bowid + put into list
                        bowid = SEPart.CreateNewBow(General.currentjob.Workspace, bowkey, backbowids(i), "F", circuit.CoreTube.Material, 10, True)
                        backbowids(i) = bowid
                        templateids.Add(bowid)
                    End If
                End If
            Next
        End If

    End Sub

    Shared Function TransformCoords(ByRef props() As List(Of Double), finneddepth As Double) As List(Of Double)()
        Dim tempx1list, tempx2list As New List(Of Double)

        For i As Integer = 0 To props(0).Count - 1
            tempx1list.Add(finneddepth - props(0)(i))
            tempx2list.Add(finneddepth - props(2)(i))
        Next
        props(0) = tempx1list
        props(2) = tempx2list
        Return props
    End Function

    Shared Sub CheckoutStutzen(totalstutzen As List(Of String))

        For Each singlestutzen In General.GetUniqueStrings(totalstutzen)
            If Not GNData.CheckTemplate(singlestutzen) Then
                WSM.CheckoutPart(singlestutzen, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
            End If
        Next

    End Sub

    Shared Sub BuildConsysAsm(ByRef consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument

        Try
            asmdoc = General.seapp.Documents.Open(consys.ConSysFile.Fullfilename)
            General.seapp.DoIdle()

            SEAsm.WriteCustomProps(asmdoc, consys.ConSysFile)

            If circuit.CircuitType.Contains("Defrost") And Not SEAsm.GroupExists(asmdoc, "Supporttubes", True) Then
                'add supporttubes to assembly
                SEAsm.AdjustConsysAssemblyDefrost(asmdoc, consys, circuit, General.currentunit.TubeSheet.Thickness)
            End If

            'gather information about CT position
            consys.CoreTubes.Add(New CoreTubeLocations)
            SEAsm.FillCoretubePositions(asmdoc, General.GetShortName(circuit.CoreTube.FileName).Substring(6, 4), consys.CoreTubes.First)

            If circuit.IsOnebranchEvap Then
                SEAsm.PlaceSingleBranch(asmdoc, consys.InletHeaders.First, consys, circuit)

                SEAsm.PlaceSingleBranch(asmdoc, consys.OutletHeaders.First, consys, circuit)

                SEAsm.PlaceSingleParts(asmdoc, consys.InletHeaders.First, consys, circuit, coil)

                SEAsm.PlaceSingleParts(asmdoc, consys.OutletHeaders.First, consys, circuit, coil)

                If circuit.Pressure <= 16 Then
                    With consys.InletHeaders
                        .First.Ventsize = GNData.GetGACVVentsize(.First, circuit, coil.NoRows)
                        .First.Tube.Materialcodeletter = consys.HeaderMaterial
                        .First.VentIDs = GNData.GetVentIDs(.First, .First.Ventsize, circuit.NoDistributions, "")

                        For Each s In .First.VentIDs
                            WSM.CheckoutPart(s, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        Next
                    End With

                    With consys.OutletHeaders
                        .First.Ventsize = GNData.GetGACVVentsize(.First, circuit, coil.NoRows)
                        .First.Tube.Materialcodeletter = consys.HeaderMaterial
                        .First.VentIDs = GNData.GetVentIDs(.First, .First.Ventsize, circuit.NoDistributions, "")

                        For Each s In .First.VentIDs
                            WSM.CheckoutPart(s, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                        Next
                    End With

                    If consys.HeaderMaterial = "C" Then
                        With consys.InletHeaders
                            If .First.Ventsize <> "G1/8" Then
                                'assemble inlet + stutzen
                                SEAsm.PlaceSingleVentsCU(asmdoc, .First, consys, circuit)
                            Else
                                SEAsm.AssembleVent5(asmdoc, .First, consys, circuit.ConnectionSide, circuit.FinType)
                            End If
                        End With

                        With consys.OutletHeaders
                            If .First.Ventsize <> "G1/8" Then
                                'assemble outlet + stutzen
                                SEAsm.PlaceSingleVentsCU(asmdoc, .First, consys, circuit)
                            Else
                                SEAsm.AssembleVent5(asmdoc, .First, consys, circuit.ConnectionSide, circuit.FinType)
                            End If
                        End With
                    Else
                        'assemble inlet
                        SEAsm.PlaceSingleVentsVA(asmdoc, consys.InletHeaders.First, consys, circuit)
                        'assemble outlet
                        SEAsm.PlaceSingleVentsVA(asmdoc, consys.OutletHeaders.First, consys, circuit)
                    End If
                End If
            ElseIf circuit.NoDistributions = 1 And circuit.CircuitType = "Defrost" Then

                With consys.InletHeaders
                    .First.Ventsize = GNData.GetGACVVentsize(.First, circuit, 0)
                    .First.VentIDs = GNData.GetVentIDs(.First, .First.Ventsize, 1, "")
                    For Each ID In .First.VentIDs
                        WSM.CheckoutPart(ID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    Next
                End With

                With consys.OutletHeaders
                    .First.Ventsize = GNData.GetGACVVentsize(.First, circuit, 0)
                    .First.VentIDs = GNData.GetVentIDs(.First, .First.Ventsize, 1, "")
                End With

                SEAsm.PlaceSingleBrineStutzen(asmdoc, consys.InletHeaders.First, consys, circuit)
                SEAsm.PlaceSingleBrineStutzen(asmdoc, consys.OutletHeaders.First, consys, circuit)

                SEAsm.PlaceSingleBrineParts(asmdoc, consys.InletHeaders.First, consys)
                SEAsm.PlaceSingleBrineParts(asmdoc, consys.OutletHeaders.First, consys)

                'assemble vent with stutzen
                SEAsm.PlaceSingleBrineVents(asmdoc, consys.InletHeaders.First, consys)
                SEAsm.PlaceSingleBrineVents(asmdoc, consys.OutletHeaders.First, consys)

            Else
                'place headers
                SEAsm.PlaceHeaders(asmdoc, consys)

                'place stutzen
                Dim n As Integer = 1
                If consys.InletHeaders.First.Tube.Quantity > 0 Then
                    For i As Integer = 0 To consys.InletHeaders.Count - 1
                        SEAsm.PlaceStutzen(asmdoc, consys.InletHeaders(i), consys, circuit)
                        'check for tube caps
                        If consys.InletHeaders(i).Tube.TopCapNeeded Then
                            Dim capwhole As Boolean = False
                            If consys.InletHeaders(i).Tube.SVPosition(0) = "cap" And circuit.ConnectionSide = "left" Then  'left
                                capwhole = True
                            End If
                            consys.InletHeaders(i).Tube.TopCapID = Database.GetCapID(consys.InletHeaders(i).Tube.Diameter, consys.InletHeaders(i).Tube.Materialcodeletter, circuit.Pressure, capwhole)
                            'assemble
                            SEAsm.AddTubeCap(asmdoc, consys.InletHeaders(i).Tube, consys.InletHeaders(i).Tube.TopCapID, consys, circuit.Pressure, "bottom", 0)
                            If capwhole Then
                                SEAsm.AddSV2Cap(asmdoc, GNData.GetSVID(circuit.Pressure), consys, "inlet")
                            End If
                        End If

                        If consys.InletHeaders(i).Tube.BottomCapNeeded Then
                            Dim capwhole As Boolean = False
                            If consys.InletHeaders(i).Tube.SVPosition(0) = "cap" And circuit.ConnectionSide = "right" Then   'right
                                capwhole = True
                            End If
                            consys.InletHeaders(i).Tube.BottomCapID = Database.GetCapID(consys.InletHeaders(i).Tube.Diameter, consys.InletHeaders(i).Tube.Materialcodeletter, circuit.Pressure, capwhole)
                            'assemble
                            SEAsm.AddTubeCap(asmdoc, consys.InletHeaders(i).Tube, consys.InletHeaders(i).Tube.BottomCapID, consys, circuit.Pressure, "top", 0)
                            If capwhole Then
                                SEAsm.AddSV2Cap(asmdoc, GNData.GetSVID(circuit.Pressure), consys, "inlet")
                            End If
                        End If

                        If consys.InletHeaders(i).Tube.SVPosition(0) = "header" Then
                            SEAsm.AddSV(asmdoc, consys.InletHeaders(i).Tube, GNData.GetSVID(circuit.Pressure), consys, circuit.ConnectionSide)
                        End If

                        'add nipple tubes
                        Dim pcount As Integer = 0
                        For Each p In consys.InletHeaders(i).Nipplepositions
                            pcount += 1
                            If consys.InletNipples.First.Materialcodeletter = "D" Then
                                'add adapter
                                Dim adapterID As String = GNData.GetAdapterID(consys.InletNipples.First.Diameter)
                                WSM.CheckoutPart(adapterID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                                SEAsm.AddNippleTube(asmdoc,
                                                New TubeData With {.Angle = consys.InletNipples.First.Angle, .IsBrine = False, .FileName = General.GetShortName(General.GetFullFilename(General.currentjob.Workspace, adapterID, "par")),
                                                .HeaderType = "inlet", .TubeType = "adapter", .BottomCapID = adapterID}, p, consys.Occlist, consys.InletHeaders(i), circuit, consys.HeaderAlignment)
                                SEAsm.AssembleAdapterNipple(asmdoc, consys.InletNipples(i), consys.Occlist)
                            Else
                                SEAsm.AddNippleTube(asmdoc, consys.InletNipples.First, p, consys.Occlist, consys.InletHeaders(i), circuit, consys.HeaderAlignment)
                            End If

                            If consys.HasFTCon AndAlso consys.FlangeID <> "" Then
                                If consys.ConType = 1 Then
                                    'thread
                                    SEAsm.PlaceThreads(asmdoc, consys, "inlet")
                                Else
                                    SEAsm.PlaceFlanges(asmdoc, consys, consys.InletNipples.First)
                                End If
                            End If

                            'check if cap needs hole for SV
                            Dim capwhole As Boolean = False
                            If consys.InletNipples.First.SVPosition(0) = "cap" And pcount = 1 Then
                                capwhole = True
                            End If

                            If consys.InletNipples(i).TopCapNeeded Then
                                consys.InletNipples(i).TopCapID = Database.GetCapID(consys.InletNipples(i).Diameter, consys.InletNipples(i).Materialcodeletter, circuit.Pressure, capwhole)
                                If capwhole AndAlso consys.InletNipples(i).TopCapID = "" Then
                                    consys.InletNipples(i).TopCapID = Database.GetCapID(consys.InletNipples(i).Diameter, consys.InletNipples(i).Materialcodeletter, circuit.Pressure, False)
                                End If
                                'assemble
                                SEAsm.AddTubeCap(asmdoc, consys.InletNipples.First, consys.InletNipples.First.TopCapID, consys, circuit.Pressure, "bottom", n)
                                n += 1
                            End If

                            If consys.InletNipples.First.SVPosition(0) <> "" And pcount = 1 Then
                                SEAsm.AddSV(asmdoc, consys.InletNipples(0), GNData.GetSVID(circuit.Pressure), consys, circuit.ConnectionSide, i + 1)
                            End If
                        Next
                    Next
                End If

                n = 1
                For i As Integer = 0 To consys.OutletHeaders.Count - 1
                    SEAsm.PlaceStutzen(asmdoc, consys.OutletHeaders(i), consys, circuit)
                    'check for tube caps
                    If consys.OutletHeaders(i).Tube.TopCapNeeded Then
                        Dim capwhole As Boolean = False 'left / right
                        If consys.OutletHeaders(i).Tube.SVPosition(0) = "cap" And
                            ((circuit.ConnectionSide = "left" And Not consys.OutletHeaders.First.Tube.IsBrine) Or (circuit.ConnectionSide = "right" And consys.OutletHeaders.First.Tube.IsBrine)) Then
                            capwhole = True
                        End If
                        consys.OutletHeaders(i).Tube.TopCapID = Database.GetCapID(consys.OutletHeaders(i).Tube.Diameter, consys.OutletHeaders(i).Tube.Materialcodeletter, circuit.Pressure, capwhole)
                        'assemble
                        SEAsm.AddTubeCap(asmdoc, consys.OutletHeaders(i).Tube, consys.OutletHeaders(i).Tube.TopCapID, consys, circuit.Pressure, "bottom", 0)
                        If capwhole Then
                            SEAsm.AddSV2Cap(asmdoc, GNData.GetSVID(circuit.Pressure), consys, "outlet")
                        End If
                    End If

                    If consys.OutletHeaders(i).Tube.BottomCapNeeded Then
                        Dim capwhole As Boolean = False 'right / left
                        If consys.OutletHeaders(i).Tube.SVPosition(0) = "cap" And
                            ((circuit.ConnectionSide = "right" And Not consys.OutletHeaders.First.Tube.IsBrine) Or (circuit.ConnectionSide = "left" And consys.OutletHeaders.First.Tube.IsBrine)) Then
                            capwhole = True
                        End If
                        consys.OutletHeaders(i).Tube.BottomCapID = Database.GetCapID(consys.OutletHeaders(i).Tube.Diameter, consys.OutletHeaders(i).Tube.Materialcodeletter, circuit.Pressure, capwhole)

                        'assemble
                        SEAsm.AddTubeCap(asmdoc, consys.OutletHeaders(i).Tube, consys.OutletHeaders(i).Tube.BottomCapID, consys, circuit.Pressure, "top", 0)
                        If capwhole Then
                            SEAsm.AddSV2Cap(asmdoc, GNData.GetSVID(circuit.Pressure), consys, "outlet")
                        End If
                    End If

                    If consys.OutletHeaders(i).Tube.SVPosition(0) = "header" Then
                        SEAsm.AddSV(asmdoc, consys.OutletHeaders(i).Tube, GNData.GetSVID(circuit.Pressure), consys, circuit.ConnectionSide)
                    End If

                    'add nipple tubes
                    Dim pcount As Integer = 0
                    For Each p In consys.OutletHeaders(i).Nipplepositions
                        pcount += 1
                        If consys.OutletNipples.First.Materialcodeletter = "D" Then
                            'add adapter
                            Dim adapterID As String = GNData.GetAdapterID(consys.OutletNipples.First.Diameter)
                            WSM.CheckoutPart(adapterID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                            SEAsm.AddNippleTube(asmdoc,
                                            New TubeData With {.Angle = consys.OutletNipples.First.Angle, .IsBrine = False, .FileName = General.GetShortName(General.GetFullFilename(General.currentjob.Workspace, adapterID, "par")),
                                            .HeaderType = "outlet", .TubeType = "adapter", .BottomCapID = adapterID}, p, consys.Occlist, consys.OutletHeaders(i), circuit, consys.HeaderAlignment)
                            SEAsm.AssembleAdapterNipple(asmdoc, consys.OutletNipples(i), consys.Occlist)
                        Else
                            SEAsm.AddNippleTube(asmdoc, consys.OutletNipples.First, p, consys.Occlist, consys.OutletHeaders(i), circuit, consys.HeaderAlignment)
                        End If

                        If consys.HasFTCon AndAlso consys.FlangeID <> "" Then
                            If consys.ConType = 1 Then
                                'thread
                                SEAsm.PlaceThreads(asmdoc, consys, "outlet")
                            Else
                                SEAsm.PlaceFlanges(asmdoc, consys, consys.OutletNipples.First)
                            End If
                        End If

                        If consys.VType = "X" Then
                            'get the pressure pipe
                            SEAsm.AddCapillaryTube(asmdoc, consys, consys.OutletHeaders(i).Tube, circuit.CoreTube.Material.Substring(0, 1))
                        End If

                        Dim capwhole As Boolean = False
                        If consys.OutletNipples(i).SVPosition(0) = "cap" And pcount = 1 Then
                            capwhole = True
                        End If

                        If consys.OutletNipples(i).TopCapNeeded Then
                            consys.OutletNipples(i).TopCapID = Database.GetCapID(consys.OutletNipples(i).Diameter, consys.OutletNipples(i).Materialcodeletter, circuit.Pressure, capwhole)
                            If capwhole AndAlso consys.OutletNipples(i).TopCapID = "" Then
                                consys.OutletNipples(i).TopCapID = Database.GetCapID(consys.OutletNipples(i).Diameter, consys.OutletNipples(i).Materialcodeletter, circuit.Pressure, False)
                            End If
                            'assemble
                            SEAsm.AddTubeCap(asmdoc, consys.OutletNipples.First, consys.OutletNipples.First.TopCapID, consys, circuit.Pressure, "bottom", n)
                            n += 1
                        End If

                        If consys.OutletNipples.First.SVPosition(0) <> "" And pcount = 1 Then
                            If i = 1 AndAlso consys.OutletNipples.Count > 1 AndAlso General.currentunit.ModelRangeName = "GACV" Then
                                SEAsm.AddSV(asmdoc, consys.OutletNipples(0), GNData.GetSVID(circuit.Pressure), consys, circuit.ConnectionSide, 2)
                            Else
                                SEAsm.AddSV(asmdoc, consys.OutletNipples(i), GNData.GetSVID(circuit.Pressure), consys, circuit.ConnectionSide, 1)
                            End If
                        End If
                    Next
                Next

                'check vents + valves and load them from PDM
                For Each h In consys.InletHeaders
                    If h.Ventsize <> "" Then
                        h.VentIDs = GNData.GetVentIDs(h, h.Ventsize, circuit.NoPasses, circuit.FinType + coil.NoRows.ToString)
                        For Each id In h.VentIDs
                            WSM.CheckoutPart(id, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                            General.WaitForFile(General.currentjob.Workspace, id, "par", 50)
                        Next
                        SEAsm.PlaceVentsConsys(asmdoc, h, consys, circuit)
                    End If
                Next

                For Each h In consys.OutletHeaders
                    If h.Ventsize <> "" Then
                        h.VentIDs = GNData.GetVentIDs(h, h.Ventsize, circuit.NoPasses, circuit.FinType + coil.NoRows.ToString)
                        For Each id In h.VentIDs
                            WSM.CheckoutPart(id, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                            General.WaitForFile(General.currentjob.Workspace, id, "par", 50)
                        Next
                        SEAsm.PlaceVentsConsys(asmdoc, h, consys, circuit)
                    End If
                Next

                'check for orific ok plates for vert AP and CP
                If General.currentunit.ModelRangeName = "GACV" And consys.HeaderAlignment = "vertical" And (General.currentunit.ModelRangeSuffix = "CP" Or General.currentunit.ModelRangeSuffix = "AP") And Not circuit.CircuitType.Contains("Defrost") Then
                    If PCFData.GetValue("OrificePlateMounted", "Quantity") <> "" Then
                        CSData.GetAddPlates(asmdoc, consys, circuit.CoreTubeOverhang)
                    End If
                End If
            End If

            'view configs
            Dim blankconfig As String = "BlankC"
            If circuit.CircuitType.Contains("Defrost") Then
                blankconfig = "BlankS"
            End If

            SEAsm.CreateConsysViews(asmdoc, consys, circuit, blankconfig)

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetCoversheet(conside As String) As String
        Dim teilenummer, side, filetype, docname, psmdocname As String
        Dim hasflange As Integer

        Try
            psmdocname = ""
            'get data from XML
            If conside = "right" Then
                'ANS
                side = "ConnectionCoveringTypeA"
            Else
                'BGS
                side = "ConnectionCoveringTypeB"
            End If

            teilenummer = GetTeilenummer(side)
            hasflange = PCFData.GetValue(side, "IsForFlangeConnection")

            If hasflange = 1 Then
                filetype = "asm"
            Else
                filetype = "par"
            End If

            WSM.CheckoutPart(teilenummer, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile, filetype)

            If hasflange = 1 Then
                General.WaitForFile(General.currentjob.Workspace, teilenummer, ".asm", 100)
                'identify the correct psm
                docname = General.GetFullFilename(General.currentjob.Workspace, teilenummer, ".asm")
                psmdocname = SEAsm.GetPSMfromASM(docname)
            Else
                General.WaitForFile(General.currentjob.Workspace, teilenummer, ".psm", 100)
                psmdocname = General.GetFullFilename(General.currentjob.Workspace, teilenummer, ".psm")
            End If

            If psmdocname <> "" Then
                WSM.CheckoutCircs(psmdocname.Substring(psmdocname.LastIndexOf("\") + 1, 10), General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return psmdocname
    End Function

    Shared Function GetTeilenummer(side As String) As String
        Dim pcfERP, pdmERP, pdmid As String

        pdmid = ""

        Try
            'get erpcode
            pcfERP = PCFData.GetValue(side, "ERPCode")

            'convert erpcode to the one connected to 3D (must be G)
            pdmERP = ConvertERPCode(pcfERP)

            'get teilenummer
            pdmid = Database.GetTNfromERP(pdmERP)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return pdmid
    End Function

    Shared Function ConvertERPCode(originalerp As String) As String
        Dim newerp As String = "AP7." + originalerp.Substring(4, 6) + "G"
        Return newerp
    End Function

    Shared Sub PrepareForBackup()
        Dim wsname As String = General.currentjob.Workspace.Substring(General.currentjob.Workspace.LastIndexOf("\") + 1)
        General.CreateWinDir("C:\Import\BackupWS")

        Try
            If IO.Directory.Exists("C:\Import\BackupWS\" + wsname) Then
                General.DeleteFolder("C:\Import\BackupWS\" + wsname)
            End If

            General.CopyDir(General.currentjob.Workspace, "C:\Import\BackupWS\" + wsname)
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

End Class
