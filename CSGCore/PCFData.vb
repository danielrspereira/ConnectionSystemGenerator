Imports System.Xml

Public Class PCFData
    Public Shared groupnames, parentgroups As New List(Of String)
    Public Shared namelists, valuelists As New List(Of List(Of String))
    Public Shared unitdescr As String

    Shared Sub ResetData()
        groupnames.Clear()
        namelists.Clear()
        valuelists.Clear()
        parentgroups.Clear()
        unitdescr = ""
    End Sub

    Shared Function GatherDatafromPCF(pcffile As String) As String
        Dim objxml As New XmlDocument
        Dim objnodelist, basenode As XmlNodeList
        Dim nodelist As New List(Of XmlNode)
        Dim objnode As XmlNode
        Dim nodeprops() As List(Of String)
        Dim innerText, mtype As String

        mtype = ""
        Try
            ResetData()

            objxml.Load(pcffile)

            basenode = objxml.GetElementsByTagName("base_component")
            CreateParentlist(basenode.Item(0), "Unit")
            parentgroups.RemoveAt(0)

            objnodelist = objxml.GetElementsByTagName("component")
            mtype = objxml.SelectSingleNode("configuration/meta/manufacturing_type").InnerText

            For Each objnode In objnodelist
                Dim objattribute As XmlAttribute = GetNameAttribute(objnode)

                If objattribute IsNot Nothing Then
                    innerText = objattribute.InnerText
                    For Each childnode As XmlNode In objnode.ChildNodes
                        If childnode.Name = "attributes" Then
                            nodeprops = General.StringToLists(childnode.InnerXml, {"<attribute ", "/attributes"})
                            If nodeprops(0).Count > 0 And nodeprops(1).Count > 0 Then
                                If objattribute.InnerText.Contains("USU_") Then
                                    unitdescr = objattribute.InnerText.Split({"_"}, 0)(1)
                                    groupnames.Add("USU")
                                Else
                                    groupnames.Add(objattribute.InnerText)
                                End If
                                namelists.Add(nodeprops(0))
                                valuelists.Add(nodeprops(1))
                            End If
                        End If
                    Next

                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return mtype
    End Function

    Shared Sub CreateParentlist(objnode As XmlNode, parent As String)
        Try
            Dim nameattribute = GetNameAttribute(objnode)

            If nameattribute IsNot Nothing Then
                parentgroups.Add(parent + "," + nameattribute.InnerText)

                For Each childnode As XmlNode In objnode.ChildNodes
                    If childnode.Name = "components" Then
                        For i As Integer = 0 To childnode.ChildNodes.Count - 1
                            CreateParentlist(childnode.ChildNodes(i), parent + "," + nameattribute.InnerText)
                        Next
                    End If
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetNameAttribute(objnode As XmlNode) As XmlAttribute
        For Each objattribute As XmlAttribute In objnode.Attributes
            If objattribute.Name = "name" Then
                Return objattribute
            End If
        Next
    End Function

    Shared Sub FillUnitDataFromXML(ByRef currentunit As UnitData)
        Dim sheetgroup As String = "TubeSheetTypeA"
        With currentunit
            .ApplicationType = GetApplicationType()
            .HasBrineDefrost = GetBrineDefrost()
            .HasIntegratedSubcooler = GetSubcooler()
            .ModelRangeName = GetValue("USU", "ModelRangeName")
            If .ModelRangeName.Substring(3, 1) = "V" Then
                .UnitSize = "Vario"
            Else
                .UnitSize = "Compact"
            End If
            .ModelRangeSuffix = GetValue("USU", "ModelRangeSuffix")
            .MultiCircuitDesign = GetValue("USU", "MultipleCircuitDesign")
            .UnitDescription = unitdescr
            .OrderData = General.currentjob
            If groupnames.IndexOf(sheetgroup) = -1 Then
                sheetgroup = "TubeSheetTypeB"
                If groupnames.IndexOf(sheetgroup) = -1 Then
                    sheetgroup = "TubeSheetTypeC"
                End If
            End If
            .TubeSheet = New SheetData With {
            .Thickness = GetValue(sheetgroup, "SheetThickness", "double"),
            .IsPowderCoated = False,
            .MaterialCodeLetter = GetValue(sheetgroup, "MaterialCodeLetter"),
            .PunchingType = GetValue(sheetgroup, "TubeSheetPunchingType"),
            .Dim_d = GetValue(sheetgroup, "ExcessToHeatExchangerInAD", "double")}

            If .TubeSheet.MaterialCodeLetter = "S" Then
                .TubeSheet.IsPowderCoated = True
            End If
            If GetValue("USU", "PEDClassification").Contains("inf") Then
                .PEDCategory = 0
            Else
                .PEDCategory = GetValue("USU", "PEDClassification", "double")
            End If
            .SELanguageID = General.userlangID
        End With

    End Sub

    Shared Function FillCoilDataFromXML(coilnumber As Integer) As CoilData
        Dim newcoil As New CoilData With {.Number = coilnumber}

        With newcoil
            .Alignment = GetCoilAlignment()
            .FinnedDepth = GetValue("HeatExchangerModuleFinoox", "FinnedDepth", "double")
            .FinnedHeight = GetValue("HeatExchangerModuleFinoox", "FinnedHeight", "double")
            .FinnedLength = GetValue("HeatExchangerModuleFinoox", "FinnedLength", "double")
            .EDefrostPDMID = GetValue("ElHeatingPositionScheme", "PDMID")
            .NoBlindTubes = GetValue("FluidCircuitPartMain", "NoBlindTubes")
            .NoRows = GetValue("FluidCircuitPartMain", "NoCoreTubeRows")
            .NoLayers = GetValue("FluidCircuitPartMain", "NoCoreTubeLayers")
            If General.currentunit.UnitDescription = "Dual" Then
                .Gap = GNData.GetGADCGap(.NoRows)
            ElseIf General.currentunit.ModelRangeName.Substring(2, 2) = "DC" Then
                .FinnedHeight += 50
            End If
        End With
        Return newcoil
    End Function

    Shared Function FillCircuitDataFromXML(circtype As String, circnumber As Integer, coil As CoilData, circcounter As Integer) As CircuitData
        Dim newcircuit As New CircuitData With {.Coilnumber = coil.Number}
        Dim fgroup As String = "FluidCircuit" + circnumber.ToString
        Dim fpartgroup As String
        Dim coretube As TubeData = CreateTubeData("HeatExchangerModuleFinoox,CoreTube")

        If circtype.Contains("Subcooler") Then
            fpartgroup = "FluidCircuitPartSubcooler"
        Else
            fpartgroup = "FluidCircuitPartMain"
        End If

        With newcircuit
            .CircuitNumber = circcounter
            .CircuitType = circtype
            .ConnectionSide = GetConSide(GetValue("USU", "ConnectionSystemInletPosition"))
            .CoreTube = coretube
            .FinType = GetValue("FinooxFin", "FinTypeLetter")
            .NoDistributions = GetValue(fgroup + "," + fpartgroup, "NoDistributions")
            .NoPasses = GetValue(fgroup + "," + fpartgroup, "NoPasses")
            .PEDCategory = GetPEDCat(False)
            .PDMID = GetValue(fgroup + "," + fpartgroup + ",CoreTubeCircuitingScheme", "PDMID")
            .Pressure = GetValue("USU", "MaxOperatingPressure")
            .Quantity = GetValue(fgroup + "," + fpartgroup + ",CoreTubeCircuitingScheme", "Quantity", "double")
            If General.currentunit.UnitDescription = "VShape" Then
                .Quantity /= 2
                .NoDistributions /= 2
            End If
        End With

        If coretube.Materialcodeletter = "W" Or coretube.Materialcodeletter = "V" Then
            newcircuit.CoreTubeOverhang = 35
        Else
            newcircuit.CoreTubeOverhang = 25
            If newcircuit.FinType = "N" Or newcircuit.FinType = "M" Or (General.currentunit.ApplicationType = "Evaporator" And General.currentunit.TubeSheet.IsPowderCoated) Then
                newcircuit.CoreTubeOverhang = 30
            End If
        End If
        If General.currentunit.ModelRangeName.Contains("HV") Or General.currentunit.ModelRangeName.Contains("VV") And newcircuit.FinType <> "G" Then
            newcircuit.CoreTubeOverhang = 35
        ElseIf General.currentunit.UnitDescription = "VShape" And coretube.Materialcodeletter = "C" Then
            newcircuit.CoreTubeOverhang = 30
        ElseIf General.currentunit.ModelRangeName = "GGDV" Then
            newcircuit.CoreTubeOverhang = 35
        End If
        If General.currentunit.ApplicationType = "Evaporator" And (General.currentunit.ModelRangeSuffix.Substring(1, 1) = "X" Or newcircuit.Pressure <= 16) And
            newcircuit.NoDistributions = 1 And newcircuit.Quantity = 1 And Not General.currentunit.UnitDescription = "Dual" Then
            newcircuit.IsOnebranchEvap = True
        Else
            newcircuit.IsOnebranchEvap = False
        End If

        Dim finpitch() As Double = GNData.GetFinPitch(coil.Alignment, newcircuit.FinType)
        newcircuit.PitchX = finpitch(0)
        newcircuit.PitchY = finpitch(1)

        Return newcircuit
    End Function

    Shared Function FillConSysDataFromXML(circuit As CircuitData, consysnumber As Integer, consystype As String) As ConSysData
        Dim newconsys As New ConSysData With {.Circnumber = circuit.CircuitNumber}
        Dim inheader, outheader, innipple, outnipple, inconjunction, outconjunction, sifon As New TubeData
        Dim outconjunctions, inconjunctions As New List(Of TubeData)
        Dim cgroup, hmaterial, nmaterial, vtype As String
        Dim hasFTCon, hasHotgas, hasSensor, hasBallValve As Boolean
        Dim contype As Integer

        If consystype = "brine" Then
            cgroup = "BrineDefrostCoil"
            hmaterial = GetValue("BrineTube", "MaterialCodeLetter")
            nmaterial = hmaterial
            vtype = "P"
            outheader = CreateTubeData(cgroup + ",BrineDefrostConnSystemOutlet,HeaderTube")
            outheader.IsBrine = True
            inheader = CreateTubeData(cgroup + ",BrineDefrostConnSystemInlet,HeaderTube")
            inheader.IsBrine = True
            outnipple = CreateTubeData(cgroup + ",BrineDefrostConnSystemOutlet,NippleTube")
            outnipple.IsBrine = True
            innipple = CreateTubeData(cgroup + ",BrineDefrostConnSystemInlet,NippleTube")
            innipple.IsBrine = True
        Else
            'all header tubes should have same material
            hmaterial = GetValue("FluidCircuit1,ConnectionSystemOutlet,HeaderTube", "MaterialCodeLetter")
            nmaterial = GetValue("FluidCircuit1,ConnectionSystemOutlet,NippleTube", "MaterialCodeLetter")
            If hmaterial = "" Then
                hmaterial = nmaterial
            End If
            If circuit.CircuitType <> "Subcooler" Then
                cgroup = "FluidCircuit" + consysnumber.ToString
                If GetValue("ThreadFlangeConnectionFC" + consysnumber.ToString, "ERPCode") <> "" Then
                    hasFTCon = True
                    If hasFTCon Then
                        contype = GetValue("ThreadFlangeConnectionFC" + consysnumber.ToString, "ConnectionType", "double")
                    End If
                End If
                If GetValue("HotgasDefrostTray", "Quantity") <> "" Then
                    hasHotgas = True
                End If
                hasSensor = GetControlsValue("Temperature")
                If GetValue("VentilationDrainBallValve", "Quantity") <> "" Then
                    hasBallValve = True
                End If
            Else
                cgroup = "FluidCircuit" + consysnumber.ToString + ",FluidCircuitPartSubcooler"
            End If
            outheader = CreateTubeData(cgroup + ",ConnectionSystemOutlet,HeaderTube")
            inheader = CreateTubeData(cgroup + ",ConnectionSystemInlet,HeaderTube")
            outnipple = CreateTubeData(cgroup + ",ConnectionSystemOutlet,NippleTube")
            With outnipple
                If .Quantity > 0 Then
                    If General.currentunit.ApplicationType = "Evaporator" Then
                        .Angle = -90
                        If circuit.ConnectionSide = "left" Then
                            .Angle = 90
                        End If
                    ElseIf General.currentunit.UnitDescription = "VShape" Then
                        .Angle = 0
                    ElseIf circuit.Pressure > 16 Then
                        .Angle = -90
                    End If
                End If
            End With
            innipple = CreateTubeData(cgroup + ",ConnectionSystemInlet,NippleTube")
            With innipple
                If .Quantity > 0 Then
                    If General.currentunit.ApplicationType = "Evaporator" Then
                        .Angle = -90
                        If circuit.ConnectionSide = "left" Then
                            .Angle = 90
                        End If
                    End If
                End If
            End With
            vtype = General.currentunit.ModelRangeSuffix.Substring(1, 1)
            If vtype = "G" Then
                vtype = "P"
            End If
            outconjunction = CreateTubeData(cgroup + ",ConnectionSystemOutlet,ConjunctionTube")
            With outconjunction
                If .Quantity = 0 Then
                    Dim q As Integer = 1
                    Do
                        outconjunction = CreateTubeData(cgroup + ",ConnectionSystemOutlet,ConjunctionTube" + q.ToString)
                        If outconjunction.Quantity = 0 Then
                            q = 0
                        Else
                            outconjunctions.Add(outconjunction)
                            q += 1
                        End If
                    Loop Until q = 0
                Else
                    For i As Integer = 1 To .Quantity
                        outconjunctions.Add(outconjunction)
                    Next
                End If
            End With

            inconjunction = CreateTubeData(cgroup + ",ConnectionSystemInlet,ConjunctionTube")
            With inconjunction
                If .Quantity = 0 Then
                    Dim q As Integer = 1
                    Do
                        inconjunction = CreateTubeData(cgroup + ",ConnectionSystemInlet,ConjunctionTube" + q.ToString)
                        If inconjunction.Quantity = 0 Then
                            q = 0
                        Else
                            inconjunctions.Add(inconjunction)
                            q += 1
                        End If
                    Loop Until q = 0
                Else
                    For i As Integer = 1 To .Quantity
                        inconjunctions.Add(inconjunction)
                    Next
                End If
            End With

            sifon = CreateTubeData(cgroup + ",ConnectionSystemInlet,OilSifon")
            If sifon.Quantity > 0 Then
                sifon.HeaderType = "pot"
            End If
        End If

        With newconsys
            .ConType = contype
            .HasFTCon = hasFTCon
            .HasHotgas = hasHotgas
            .HasSensor = hasSensor
            .HasBallValve = hasBallValve
            If .HasHotgas Then
                .HotGasConnectionDiameter = GetValue("PipingConnection", "ConnectionDiameter", "double")
                Dim hotgasheader As String = GetHotGasHeader()
                .HotGasData = New HotgasInfo With {.Headertype = hotgasheader}
            End If
            .HeaderAlignment = GetHeaderAlignment(General.currentunit.UnitDescription, consystype)
            .HeaderMaterial = hmaterial
            .VType = vtype
            outheader.HeaderType = "outlet"
            If outheader.Quantity > 0 Then
                outheader.TubeType = "header"
                If outnipple.Quantity > 0 Then
                    If outnipple.IsBrine Then
                        If circuit.ConnectionSide = "left" Then
                            outnipple.Angle = 90
                        Else
                            outnipple.Angle = -90
                        End If
                    ElseIf .HeaderAlignment = "horizontal" And circuit.Pressure > 16 Then
                        outnipple.Angle = 90
                    End If
                    If (circuit.NoPasses = 1 Or circuit.NoPasses = 3) And outnipple.Angle = 0 Then
                        outnipple.Angle = 180
                    End If
                End If
            Else
                If vtype = "P" And circuit.Pressure > 16 Then
                    With outnipple
                        outheader.Material = .Material
                        outheader.Materialcodeletter = .Materialcodeletter
                        outheader.WallThickness = .WallThickness
                        .Quantity = 0
                    End With
                    outheader.Quantity = 1
                End If
                outheader.Diameter = outnipple.Diameter
            End If
            .OutletHeaders.Add(New HeaderData With {.Tube = outheader, .Nippletubes = outnipple.Quantity})
            inheader.HeaderType = "inlet"
            If inheader.Quantity > 0 Then
                inheader.TubeType = "header"
                If innipple.Quantity > 0 Then
                    If innipple.IsBrine Then
                        If circuit.ConnectionSide = "left" Then
                            innipple.Angle = 90
                        Else
                            innipple.Angle = -90
                        End If
                    End If
                End If
            Else
                If vtype = "P" And circuit.Pressure > 16 Then
                    With innipple
                        inheader.Material = .Material
                        inheader.Materialcodeletter = .Materialcodeletter
                        inheader.WallThickness = .WallThickness
                        .Quantity = 0
                    End With
                    inheader.Quantity = 1
                End If
                inheader.Diameter = innipple.Diameter
            End If
            .InletHeaders.Add(New HeaderData With {.Tube = inheader, .Nippletubes = innipple.Quantity})
            If outnipple.Quantity > 0 Then
                outnipple.HeaderType = "outlet"
                outnipple.TubeType = "nipple"
            End If
            .OutletNipples.Add(outnipple)
            If innipple.Quantity > 0 Then
                innipple.HeaderType = "inlet"
                innipple.TubeType = "nipple"
            End If
            .InletNipples.Add(innipple)
            If General.currentunit.UnitDescription = "Dual" Then
                If inconjunctions.Count > 0 Then
                    .InletConjunctions = inconjunctions
                Else
                    .InletConjunctions.Add(New TubeData)
                End If
                If outconjunctions.Count > 0 Then
                    .OutletConjunctions = outconjunctions
                Else
                    .OutletConjunctions.Add(New TubeData)
                End If
                .OilSifons.Add(New HeaderData With {.Tube = sifon})
            End If
        End With

        Return newconsys
    End Function

    Shared Function GetControlsValue(description As String) As Boolean
        Dim containsitem As Boolean = False

        Try
            For i As Integer = 0 To parentgroups.Count - 1
                If parentgroups(i).Contains("controls_solution") Then
                    For j As Integer = 0 To namelists(i).Count - 1
                        If namelists(i)(j) = "Description" Then
                            If valuelists(i)(j) = description Then
                                containsitem = True
                                Exit For
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception

        End Try
        Return containsitem
    End Function

    Shared Function GetValue(grnames As String, propname As String, Optional vartype As String = "string") As String
        Dim propvalue As String = ""
        Dim namelist, valuelist As List(Of String)
        Dim indexgroup, indexname As Integer
        Dim groups() As String

        Try
            If vartype.ToLower.Contains("string") Then
                propvalue = ""
            Else
                propvalue = "0"
            End If

            'Getting the path for the value
            groups = grnames.Split({","}, 0)
            indexgroup = GetGroupIndex(groups)

            If indexgroup > -1 Then
                'Getting the attributes of the group
                namelist = namelists(indexgroup)
                valuelist = valuelists(indexgroup)

                'Getting the value of the searched attribute
                indexname = namelist.IndexOf(propname)

                If indexname > -1 Then
                    propvalue = valuelist(indexname)
                    If vartype = "double" Then
                        propvalue = General.TextToDouble(propvalue)
                    End If
                End If
            End If

        Catch ex As Exception

        End Try

        Return propvalue
    End Function

    Shared Function GetGroupIndex(groups() As String) As Integer
        Dim gindex As Integer
        Dim splittedgroups() As String

        gindex = -1

        For i As Integer = 0 To parentgroups.Count - 1
            Dim gcount As Integer = 0
            For Each g In groups
                'If g.ToLower.Contains("ntube") Then
                '    If parentgroups(i) = g Then
                '        gcount += 1
                '    End If
                'Else
                'End If
                If parentgroups(i).Contains(g) Then
                    If g.ToLower.Contains("ntube") Then
                        'check for exact match
                        splittedgroups = parentgroups(i).Split({","}, 0)
                        For Each s In splittedgroups
                            If s = g Then
                                gcount += 1
                                Exit For
                            End If
                        Next
                    Else
                        gcount += 1
                    End If
                End If
            Next
            If gcount = groups.Count Then
                gindex = i
                Exit For
            End If
        Next

        Return gindex
    End Function

    Shared Function GetApplicationType() As String
        Dim apptype As String = GetValue("USU", "BaseUnitFunction")
        Dim unittype As String

        If apptype = "1" Or apptype = "2" Or apptype = "3" Then
            unittype = "Evaporator"
        Else
            unittype = "Condenser"
        End If
        Return unittype
    End Function

    Shared Function ConvertCTMat(ctcode As String) As String
        Dim ctmat As String = ""
        Dim codeletter As String = ctcode.Substring(1, 1)

        If (ctcode.Contains(".4") And ctcode.Contains("RC")) Or ctcode.Contains("V") Or ctcode.Contains("W") Then
            If ctcode.Substring(0, 2) = "RV" Then
                ctmat = "V2A (SP03-1)"
            ElseIf ctcode.Substring(0, 2) = "RW" Then
                ctmat = "V4A (SP03-2)"
            Else
                ctmat = "CU (SP01-A1 (K65))"
            End If
        Else
            Select Case codeletter
                Case "C"
                    ctmat = "CU (SP01-A)"
                Case "R"
                    ctmat = "CU (SP01-B1 (R))"
                Case "X"
                    ctmat = "CU (SP01-C (X))"
                Case "F"
                    ctmat = "F (SP02-A)"
                Case "V"
                    ctmat = "V2A (SP03-1)"
                Case "W"
                    ctmat = "V4A (SP03-2)"
            End Select
        End If

        Return ctmat
    End Function

    Shared Function GetHeaderAlignment(unittype As String, circtype As String) As String
        Dim alignment, alignno, alignangle As String

        If circtype = "brine" Then
            alignno = GetValue("BrineDefrostCircuitingScheme", "HeaderAlignment")
            If alignno = "2" Then
                alignment = "horizontal"
            Else
                alignment = "vertical"
            End If
        Else
            If unittype.Contains("VShape") Then
                alignment = "vertical"
            Else
                alignno = GetValue("CoreTubeCircuitingScheme", "HeaderAlignment")
                alignangle = GetValue("USU", "HeatExchangerAlignmentAngle")

                If alignno = "1" Then
                    If alignangle.Contains("90") Then
                        alignment = "vertical"
                    Else
                        alignment = "horizontal"
                    End If
                Else
                    If alignangle.Contains("90") Then
                        alignment = "horizontal"
                    Else
                        alignment = "vertical"
                    End If
                End If
            End If
        End If

        Return alignment
    End Function

    Shared Function GetSubcooler() As Boolean
        Dim hassubcooler As Boolean = False
        If GetValue("USU", "HasIntegratedSubcooler") = "1" Then
            hassubcooler = True
        End If
        Return hassubcooler
    End Function

    Shared Function GetBrineDefrost() As Boolean
        Dim hasbrine As Boolean = False
        If GetValue("BrineDefrostCircuitingScheme", "PDMID") <> "" Then
            hasbrine = True
        End If
        Return hasbrine
    End Function

    Shared Function GetCoilAlignment() As String
        Dim alignment As String
        If General.currentunit.UnitDescription = "Flat" Then
            alignment = "horizontal"
        Else
            alignment = "vertical"
        End If
        Return alignment
    End Function

    Shared Function GetConSide(conside As String) As String
        Dim side As String = "right"
        '1 - in air direction left (extra for GACV) // 2 - in air direction right (default) - confirmed M. Neumeyer 
        If conside = "1" Then
            side = "left"
        End If
        Return side
    End Function

    Shared Function GetPEDCat(isbrine As Boolean) As Integer
        Dim cat As Integer

        If isbrine Or GetValue("USU", "PEDClassification").Contains("inf") Then
            cat = 0
        Else
            cat = GetValue("USU", "PEDClassification", "double")
        End If
        Return cat
    End Function

    Shared Function GetHotGasHeader() As String
        Dim headertype As String
        If GetValue("PipingConnection", "HotgasFeed") = "1" Then
            headertype = "outlet"
        Else
            headertype = "inlet"
        End If
        Return headertype
    End Function

    Shared Function CreateTubeData(groupname As String) As TubeData
        Dim newtube As New TubeData

        With newtube
            .Quantity = GetValue(groupname, "Quantity", "double")
            If .Quantity > 0 Then
                .Length = GetValue(groupname, "Length", "double")
                If .Length > 0 Then
                    If General.currentunit.UnitDescription = "VShape" Then
                        .Quantity /= 2
                    End If
                    .Diameter = GetValue(groupname, "OuterDiameter", "double")
                    .Materialcodeletter = GetValue(groupname, "MaterialCodeLetter")
                    .RawMaterial = GetValue(groupname, "ERPCode")
                    .Material = ConvertCTMat(.RawMaterial)
                    .WallThickness = GetValue(groupname, "WallThickness", "double")
                Else
                    .Quantity = 0
                End If
            End If
        End With
        Return newtube
    End Function

    Shared Function CreateBrineCircuit(maincircuit As CircuitData) As CircuitData
        Dim brinecircuit As New CircuitData
        Dim brinetube As TubeData = CreateTubeData("BrineTube")

        With brinecircuit
            .Coilnumber = 1
            .CircuitNumber = 2
            .Quantity = 1
            .Pressure = 10
            .CoreTube = brinetube
            If brinetube.Materialcodeletter = "C" Then
                .CoreTubeOverhang = 30
            Else
                .CoreTubeOverhang = 35
            End If
            .CircuitType = "Defrost"
            .FinType = "F"
            .NoDistributions = GetValue("BrineDefrostCircuitingScheme", "NoDistributions")
            .NoPasses = GetValue("BrineDefrostCircuitingScheme", "NoPasses")
            .PEDCategory = 0
            .PitchX = 50
            .PitchY = 50
            .PDMID = GetValue("BrineDefrostCircuitingScheme", "PDMID")
            If maincircuit.ConnectionSide = "right" Then
                .ConnectionSide = "left"
            Else
                .ConnectionSide = "right"
            End If
            .SupportPDMID = GetValue("BrineDefSupportTubePositions", "PDMID")
        End With

        Return brinecircuit
    End Function

End Class
