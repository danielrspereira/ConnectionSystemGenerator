Imports System.IO

Public Class SEPart

    Shared Sub CreateCoreTube(ByRef coretube As TubeData, no As Integer, tubename As String, orbitalwelding As Boolean)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim partname As String = General.currentjob.Workspace + "\" + tubename + no.ToString + ".par"
        Dim refplane As SolidEdgePart.RefPlane
        Dim objsketch As SolidEdgePart.Sketch
        Dim objprofile, profilearr() As SolidEdgePart.Profile
        Dim waypoints As New List(Of Double)

        Try

            If Not File.Exists(partname) Then
                partdoc = General.seapp.Documents.Add(ProgID:="SolidEdge.PartDocument")
                General.seapp.DoIdle()
                partdoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered
                partdoc.SaveAs(partname)
                coretube.FileName = partname

                'create body
                refplane = partdoc.RefPlanes.Item(1)
                objsketch = partdoc.Sketches.Add
                objprofile = objsketch.Profiles.Add(refplane)
                objprofile.Visible = False

                'create lines
                If orbitalwelding Then
                    With coretube
                        'create lines
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000 - .WallThickness / 1000, 0, .Diameter / 2000, 0)
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000, 0, .Diameter / 2000, 0.08)
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000, 0.08, .Diameter / 2000 - .WallThickness / 1000, 0.08)
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000 - .WallThickness / 1000, 0.08, .Diameter / 2000 - .WallThickness / 1000, 0)
                    End With

                Else
                    With coretube
                        'calculate the waypoint
                        Dim keypoint() As Double = { .Diameter / 2000 + 0.001 - .WallThickness / 1000, Math.Round(0.006 - Math.Sin(15 * Math.PI / 180) / 2000, 7)}

                        'create lines
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000 + 0.001 - .WallThickness / 1000, 0, .Diameter / 2000 + 0.001, 0)
                        objprofile.Lines2d.AddBy2Points(.Diameter / 2000 + 0.001, 0, .Diameter / 2000 + 0.001, 0.006)
                        objprofile.Lines2d.AddByPointAngleLength(.Diameter / 2000 + 0.001, 0.006, 120 * Math.PI / 180, 0.002)

                        Dim xe, ye As Double

                        objprofile.Lines2d.Item(3).GetEndPoint(xe, ye)

                        objprofile.Lines2d.AddBy2Points(xe, ye, xe, 0.08)
                        objprofile.Lines2d.AddBy2Points(xe, 0.08, xe - .WallThickness / 1000, 0.08)

                        objprofile.Lines2d.AddBy2Points(keypoint(0), 0, keypoint(0), keypoint(1))
                        objprofile.Lines2d.AddByPointAngleLength(keypoint(0), keypoint(1), 120 * Math.PI / 180, 0.002)

                        objprofile.Lines2d.Item(objprofile.Lines2d.Count).GetEndPoint(xe, ye)

                        objprofile.Lines2d.AddBy2Points(xe, ye, xe, 0.08)
                    End With
                End If

                Dim ax As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, 0, 0, 0.01)

                Dim rotaxis As SolidEdgePart.RefAxis = objprofile.SetAxisOfRevolution(ax)

                objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                profilearr = {objprofile}
                objsketch.Name = "Profile"

                partdoc.Models.AddFiniteRevolvedProtrusion(1, profilearr, rotaxis, SolidEdgePart.FeaturePropertyConstants.igRight, 2 * Math.PI)
                Dim feat As SolidEdgePart.RevolvedProtrusion = partdoc.Models.Item(1).RevolvedProtrusions.Item(1)
                feat.Name = "Coretube"

                'assign material
                SetMaterial(General.seapp, partdoc, coretube.Materialcodeletter, "part")

                'custom file props needed? here just CSG & order info
                Order.AddOrderDatatoCustomProps(partdoc, "par", "Kernrohr", "Coretube", tubename.Substring(0, 4))

                General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)
            Else

            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateFin(ByRef coil As CoilData, ByRef circuit As CircuitData, tubesheet As SheetData, finname As String, no As Integer)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim partname As String = General.currentjob.Workspace + "\" + finname + coil.Number.ToString + no.ToString + ".par"
        Dim objexpros As SolidEdgePart.ExtrudedProtrusions
        Dim objfeature As SolidEdgePart.ExtrudedCutout
        Dim refplane As SolidEdgePart.RefPlane
        Dim objsheet As SolidEdgeDraft.Sheet
        Dim frameprofile, ctprofile, profilearr() As SolidEdgePart.Profile
        Dim objlines As New List(Of SolidEdgeFrameworkSupport.Line2d)
        Dim points(), h, d As Double
        Dim origin As String = ""

        Try

            If Not File.Exists(partname) Then
                partdoc = General.seapp.Documents.Add(ProgID:="SolidEdge.PartDocument")
                General.seapp.DoIdle()
                partdoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered
                partdoc.SaveAs(partname)

                'create body
                refplane = partdoc.RefPlanes.Item(3)
                frameprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
                frameprofile.Visible = False

                If coil.Alignment = "horizontal" Then
                    h = coil.FinnedDepth
                    d = coil.FinnedHeight
                Else
                    h = coil.FinnedHeight
                    d = coil.FinnedDepth
                End If
                points = {0, h / 1000, d / 1000}

                objlines.Add(frameprofile.Lines2d.AddBy2Points(points(0), points(0), points(0), points(1)))
                objlines.Add(frameprofile.Lines2d.AddBy2Points(points(0), points(1), points(2), points(1)))
                objlines.Add(frameprofile.Lines2d.AddBy2Points(points(2), points(1), points(2), points(0)))
                objlines.Add(frameprofile.Lines2d.AddBy2Points(points(2), points(0), points(0), points(0)))

                frameprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

                partdoc.Models.AddFiniteExtrudedProtrusion(1, {frameprofile}, 1, tubesheet.Thickness / 1000)
                objexpros = partdoc.Models.Item(1).ExtrudedProtrusions
                objexpros.Item(1).Name = "Fin"

                If coil.Gap > 0 And General.currentunit.UnitDescription = "Dual" Then
                    Dim secondframe As SolidEdgePart.Profile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
                    secondframe.Visible = False
                    'coil is always horizontal
                    points = {Math.Round(coil.Gap + d) / 1000, h / 1000, Math.Round((2 * d + coil.Gap) / 1000, 6)}
                    objlines.Add(secondframe.Lines2d.AddBy2Points(points(0), 0, points(0), points(1)))
                    objlines.Add(secondframe.Lines2d.AddBy2Points(points(0), points(1), points(2), points(1)))
                    objlines.Add(secondframe.Lines2d.AddBy2Points(points(2), points(1), points(2), 0))
                    objlines.Add(secondframe.Lines2d.AddBy2Points(points(2), 0, points(0), 0))

                    secondframe.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                    partdoc.Models.AddFiniteExtrudedProtrusion(1, {secondframe}, 1, tubesheet.Thickness / 1000)
                    objexpros = partdoc.Models.Item(1).ExtrudedProtrusions
                    objexpros.Item(1).Name = "FinL"
                    objexpros.Item(2).Name = "FinR"
                End If

                If circuit.CustomCirc Then
                    'circuiting not in workspace yet, fin has to be finished in this step for coil assembly → load circuiting

                    WSM.CheckoutCircs(circuit.PDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

                    'open drawing, find sheet of drawing view
                    Dim circfile As String = General.GetFullFilename(General.currentjob.Workspace, circuit.PDMID, ".dft")
                    objsheet = SEDraft.FindSheet(SEDraft.OpenDFT(circfile))

                    'check bottom left position
                    If SEDraft.ObjectAtLoc(objsheet, {circuit.PitchX / 2, circuit.PitchY / 2}, "circle") Then
                        '→ cannot be top left 
                        ctprofile = CreateCustomCTProfile(circuit, partdoc, refplane, Math.Round(circuit.PitchX * 3 / 2, 2, MidpointRounding.AwayFromZero), h - circuit.PitchY / 2)
                        origin = "topright"
                    Else
                        '→ must be top left
                        ctprofile = CreateCustomCTProfile(circuit, partdoc, refplane, circuit.PitchX / 2, h - circuit.PitchY / 2)
                        origin = "topleft"
                    End If

                    'close drawing 
                    General.seapp.Documents.CloseDocument(circfile, SaveChanges:=False, DoIdle:=True)

                    'add ct cutouts
                    profilearr = {ctprofile}
                    objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                    objfeature.Name = "CT1"
                Else
                    'add ct cutouts
                    ctprofile = CreateCTProfile(circuit, partdoc, refplane, coil.Alignment, 1, h, d, 0)
                    profilearr = {ctprofile}
                    objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                    objfeature.Name = "CT1"
                End If

                'create pattern
                CreateCTPattern(partdoc, objfeature, coil, circuit, refplane, "P1CT")

                If General.currentunit.UnitDescription = "Dual" And coil.Gap > 0 Then
                    objfeature.Name = "CT1L"
                    ctprofile = CreateCTProfile(circuit, partdoc, refplane, coil.Alignment, 1, h, d, coil.Gap)
                    profilearr = {ctprofile}
                    objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                    objfeature.Name = "CT1R"

                    CreateCTPattern(partdoc, objfeature, coil, circuit, refplane, "P1CT2")
                End If

                If circuit.CustomCirc Then
                    If origin = "topleft" Then
                        ctprofile = CreateCustomCTProfile(circuit, partdoc, refplane, circuit.PitchX * 3 / 2, h - circuit.PitchY * 3 / 2)
                    Else
                        ctprofile = CreateCustomCTProfile(circuit, partdoc, refplane, circuit.PitchX / 2, h - circuit.PitchY * 3 / 2)
                    End If
                    profilearr = {ctprofile}

                    objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                    objfeature.Name = "CT2"

                    CreateCTPattern(partdoc, objfeature, coil, circuit, refplane, "P2CT")
                Else
                    If circuit.FinType <> "N" And circuit.FinType <> "M" And coil.FinnedDepth >= Math.Max(circuit.PitchX, circuit.PitchY) * 2 Then
                        ctprofile = CreateCTProfile(circuit, partdoc, refplane, coil.Alignment, 2, h, d, 0)
                        profilearr = {ctprofile}

                        objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                        objfeature.Name = "CT2"

                        CreateCTPattern(partdoc, objfeature, coil, circuit, refplane, "P2CT")

                        If General.currentunit.UnitDescription = "Dual" And coil.Gap > 0 Then
                            objfeature.Name = "CT2L"
                            ctprofile = CreateCTProfile(circuit, partdoc, refplane, coil.Alignment, 2, h, d, coil.Gap)
                            profilearr = {ctprofile}
                            objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAll(ctprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igLeft)
                            objfeature.Name = "CT2R"

                            CreateCTPattern(partdoc, objfeature, coil, circuit, refplane, "P2CT2")
                        End If
                    End If
                End If

                'assign material
                SetMaterial(General.seapp, partdoc, tubesheet.MaterialCodeLetter, "part")

                'custom file props needed? here just CSG & order info
                Order.AddOrderDatatoCustomProps(partdoc, "par", "Lamelle", "Fin", finname.Substring(0, 4))

                General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateTube(ByRef tube As TubeData, name As String, tubetype As String, pressure As Integer)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim CDB_description() As String = WSM.GetPDMDescription(name)
        Dim objvariables As SolidEdgeFramework.Variables

        Try

            'no features, only basic information and file creation
            tube.TubeFile.Fullfilename = name
            tube.FileName = name
            tube.TubeType = tubetype

            partdoc = General.seapp.Documents.Add(ProgID:="SolidEdge.PartDocument")
            General.seapp.DoIdle()
            partdoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered
            partdoc.SaveAs(name)

            objvariables = partdoc.Variables
            objvariables.Add("Length", "100")
            objvariables.Add("Diameter", "22")
            objvariables.Add("WallThickness", "1")
            SetMaterial(General.seapp, partdoc, tube.Materialcodeletter, "part")

            tube.TubeFile.Shortname = partdoc.Name

            'get AGP number and raw material
            If tube.RawMaterial = "" Then
                'get raw material from database
                tube.RawMaterial = Database.GetTubeERP(tube.Diameter, "Headertube", pressure, tube.Materialcodeletter)
            End If

            With tube.TubeFile
                .CDB_Material = WSM.CDB_Material(tube.Materialcodeletter)
                .CDB_de = CDB_description(0)
                .CDB_en = CDB_description(1)
                .Filetype = "par"
                .Orderno = General.currentjob.OrderNumber
                .Orderpos = General.currentjob.OrderPosition
                .Projectno = General.currentjob.ProjectNumber
                .AGPno = WSM.GetAGPNumber(CDB_description(0), tube.Materialcodeletter, False)
                .CDB_Zusatzbenennung = tube.HeaderType
                If tube.IsBrine Then
                    .CDB_z_Bemerkung = "Brine Defrost"
                End If
            End With

            WriteCostumProps(partdoc, tube.TubeFile)

            General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AdjustFinDefrost(coil As CoilData, circuit As CircuitData, partname As String)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        Dim refplane As SolidEdgePart.RefPlane
        Dim profilearr() As SolidEdgePart.Profile
        Dim objsketch As SolidEdgePart.Sketch
        Dim profilelist As New List(Of SolidEdgePart.Profile)
        Dim objfeature As SolidEdgePart.ExtrudedCutout
        Dim tempcircuit As New CircuitData With {.PitchX = 50, .PitchY = 50, .FinType = "N", .ConnectionSide = circuit.ConnectionSide, .CoreTube = New TubeData With {.Diameter = 11}}

        Try
            'get the positions for needed tubes
            'use the support tube drawing
            WSM.CheckoutCircs(circuit.SupportPDMID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            dftdoc = SEDraft.OpenDFT(General.GetFullFilename(General.currentjob.Workspace, circuit.SupportPDMID, ".dft"))

            'get coretube positions
            Unit.supportcoords = SEDraft.GetSupportPositions(dftdoc, circuit.ConnectionSide, circuit.CircuitType.ToLower)

            partdoc = General.seapp.Documents.Open(partname)
            General.seapp.DoIdle()

            refplane = partdoc.RefPlanes.Item(3)
            objsketch = partdoc.Sketches.Add

            'create 1 profile for each position
            For i As Integer = 0 To Unit.supportcoords(0).Count - 1
                Dim ctprofile As SolidEdgePart.Profile
                ctprofile = objsketch.Profiles.Add(refplane)
                ctprofile.Visible = False
                ctprofile.Circles2d.AddByCenterRadius(Unit.supportcoords(0)(i) / 1000, Unit.supportcoords(1)(i) / 1000, tempcircuit.CoreTube.Diameter / 2000)
                ctprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                profilelist.Add(ctprofile)
            Next

            profilearr = profilelist.ToArray
            objfeature = partdoc.Models.Item(1).ExtrudedCutouts.AddThroughAllMulti(profilelist.Count, profilearr, SolidEdgePart.FeaturePropertyConstants.igLeft)
            objfeature.Name = "ST"

            General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AdjustFinEDefrost(heatingpos As List(Of Double)(), filename As String)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim refplanes As SolidEdgePart.RefPlanes
        Dim refplane As SolidEdgePart.RefPlane
        Dim objmodel As SolidEdgePart.Model
        Dim objsketches As SolidEdgePart.Sketchs
        Dim objsketch As SolidEdgePart.Sketch
        Dim objprofiles As SolidEdgePart.Profiles
        Dim objcutout As SolidEdgePart.ExtrudedCutout
        Dim objprofilearray As New List(Of SolidEdgePart.Profile)

        Try
            partdoc = General.seapp.Documents.Open(filename)
            General.seapp.DoIdle()
            objmodel = partdoc.Models.Item(1)
            refplanes = partdoc.RefPlanes
            refplane = refplanes.Item(3)
            objsketches = partdoc.Sketches
            objsketch = objsketches.Add

            objprofiles = objsketch.Profiles

            For i As Integer = 0 To heatingpos(0).Count - 1
                Dim newprofile As SolidEdgePart.Profile = objprofiles.Add(refplane)
                newprofile.Circles2d.AddByCenterRadius(heatingpos(0)(i) / 1000, heatingpos(1)(i) / 1000, 0.0055)
                newprofile.Visible = False
                objprofilearray.Add(newprofile)
            Next

            objcutout = objmodel.ExtrudedCutouts.AddThroughNextMulti(heatingpos(0).Count, objprofilearray.ToArray, SolidEdgePart.FeaturePropertyConstants.igLeft)
            objcutout.Name = "SupportTubes"

            General.seapp.Documents.CloseDocument(filename, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function CreateCTProfile(circuit As CircuitData, partdoc As SolidEdgePart.PartDocument, refplane As SolidEdgePart.RefPlane, alignment As String, row As Integer, fheight As Double,
                                    fdepth As Double, gap As Double) As SolidEdgePart.Profile
        Dim ctprofile As SolidEdgePart.Profile = Nothing
        Dim objcircle As SolidEdgeFrameworkSupport.Circle2d
        Dim x0, y0 As Double

        Try
            ctprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            ctprofile.Visible = False

            If row = 1 Then
                y0 = circuit.PitchY / 2
                If General.currentunit.UnitDescription = "Dual" Then
                    If (gap = 0 And circuit.ConnectionSide = "right") Or (gap > 0 And circuit.ConnectionSide = "left") Then
                        'top right, first coil
                        x0 = Math.Round(circuit.PitchX * 3 / 2, 2, MidpointRounding.AwayFromZero)
                    Else
                        'top left, second coil
                        x0 = circuit.PitchX / 2 + gap + fdepth
                    End If
                ElseIf (circuit.ConnectionSide = "right" And (General.currentunit.ApplicationType = "Evaporator" Or alignment = "horizontal" Or General.currentunit.UnitDescription = "VShape")) Or
                    circuit.FinType = "N" Or circuit.FinType = "M" Or (General.currentunit.ApplicationType = "Condenser" And (fdepth = 75 Or fdepth = 64.95)) Or
                    (circuit.ConnectionSide = "left" And Not General.currentunit.UnitDescription = "VShape" And General.currentunit.ApplicationType = "Condenser" And alignment = "vertical") Then
                    x0 = circuit.PitchX / 2
                Else
                    x0 = Math.Round(circuit.PitchX * 3 / 2, 2, MidpointRounding.AwayFromZero)
                End If
            Else
                y0 = circuit.PitchY * 3 / 2
                If General.currentunit.UnitDescription = "Dual" Then
                    If (gap = 0 And circuit.ConnectionSide = "right") Or (gap > 0 And circuit.ConnectionSide = "left") Then
                        'top left, second coil
                        x0 = circuit.PitchX / 2
                    Else
                        'top right, first coil
                        x0 = Math.Round(circuit.PitchX * 3 / 2, 2, MidpointRounding.AwayFromZero) + gap + fdepth
                    End If
                ElseIf circuit.ConnectionSide = "right" And (General.currentunit.ApplicationType = "Evaporator" Or alignment = "horizontal" Or General.currentunit.UnitDescription = "VShape") Or
                       (General.currentunit.ApplicationType = "Condenser" And (fdepth = 75 Or fdepth = 64.95)) Or
                       (circuit.ConnectionSide = "left" And Not General.currentunit.UnitDescription = "VShape" And General.currentunit.ApplicationType = "Condenser" And alignment = "vertical") Then
                    x0 = Math.Round(circuit.PitchX * 3 / 2, 2, MidpointRounding.AwayFromZero)
                Else
                    x0 = circuit.PitchX / 2
                End If
            End If

            GetSetVariableValue("X0", partdoc.Variables, "add", x0, 1)
            GetSetVariableValue("Y0", partdoc.Variables, "add", y0, 1)

            objcircle = ctprofile.Circles2d.AddByCenterRadius(x0 / 1000, (fheight - y0) / 1000, circuit.CoreTube.Diameter / 2000)
            ctprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

        Catch ex As Exception

        End Try

        Return ctprofile
    End Function

    Shared Function CreateCustomCTProfile(circuit As CircuitData, partdoc As SolidEdgePart.PartDocument, refplane As SolidEdgePart.RefPlane, xpos As Double, ypos As Double) As SolidEdgePart.Profile
        Dim ctprofile As SolidEdgePart.Profile = Nothing
        Dim objcircle As SolidEdgeFrameworkSupport.Circle2d

        Try
            ctprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            ctprofile.Visible = False

            GetSetVariableValue("X0", partdoc.Variables, "add", xpos, 1)
            GetSetVariableValue("Y0", partdoc.Variables, "add", ypos, 1)

            objcircle = ctprofile.Circles2d.AddByCenterRadius(xpos / 1000, ypos / 1000, circuit.CoreTube.Diameter / 2000)
            ctprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

        Catch ex As Exception

        End Try

        Return ctprofile
    End Function

    Shared Sub CreateCTPattern(partdoc As SolidEdgePart.PartDocument, objfeature As SolidEdgePart.ExtrudedCutout, coil As CoilData, circuit As CircuitData, refplane As SolidEdgePart.RefPlane, pattername As String)
        Dim patternprofile As SolidEdgePart.Profile

        Try
            'pattern profile sketch
            patternprofile = CreatePatternProfile(partdoc, refplane, objfeature, coil, circuit, pattername)

            Dim firstpat As SolidEdgePart.Pattern = partdoc.Models.Item(1).Patterns.Add(1, {objfeature}, patternprofile, SolidEdgePart.PatternTypeConstants.seFastPattern)
            firstpat.Name = pattername

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function CreatePatternProfile(partdoc As SolidEdgePart.PartDocument, refplane As SolidEdgePart.RefPlane, objfeature As SolidEdgePart.ExtrudedCutout, coil As CoilData, circuit As CircuitData, pname As String) As SolidEdgePart.Profile
        Dim objsketch As SolidEdgePart.Sketch
        Dim patternprofile As SolidEdgePart.Profile
        Dim objprofile As SolidEdgePart.Profile = objfeature.Profile
        Dim d, h, x, y, pitchx, pitchy As Double

        objprofile.Circles2d.Item(1).GetCenterPoint(x, y)

        If circuit.FinType <> "N" And circuit.FinType <> "M" Then
            pitchx = 2 * circuit.PitchX
            pitchy = 2 * circuit.PitchY
        Else
            pitchx = circuit.PitchX
            pitchy = circuit.PitchY
        End If
        If coil.Alignment = "horizontal" Then
            d = coil.FinnedHeight
            h = coil.FinnedDepth
            If General.IntegerRem(h, pitchy) > 0 And ((h / 1000 - y) < circuit.PitchY / 1000 Or h = pitchy) Then
                h += pitchy / 2
            End If
        Else
            d = coil.FinnedDepth
            h = coil.FinnedHeight
            If General.IntegerRem(d, pitchx) > 0 And (x < circuit.PitchX / 1000 Or d = pitchx) Then
                d += pitchx / 2
            End If
        End If

        objsketch = partdoc.Sketches.Add
        objsketch.Name = pname + "Sketch"
        patternprofile = objsketch.Profiles.Add(refplane)
        patternprofile.Visible = False
        patternprofile.RectangularPatterns2d.Add(x, y, 'Origin
                                                 (h - pitchy) / 1000, (d - pitchx) / 1000, 3 * Math.PI / 2, 'Width, Height, Angle
                                                SolidEdgeFrameworkSupport.PatternOffsetTypeConstants.sePatternFillOffset,
                                                1, 1, 'Quantity X & Y - not important when using filloffset methode
                                                pitchy / 1000, pitchx / 1000)

        Return patternprofile
    End Function

    Shared Sub SetMaterial(seapp As SolidEdgeFramework.Application, obj As Object, materialcode As String, filetype As String)
        Dim sematerialname As String = ""
        Dim objmattable As SolidEdgeFramework.MatTable
        Dim partdoc As SolidEdgePart.PartDocument
        Dim psmdoc As SolidEdgePart.SheetMetalDocument

        Try
            'going to be replaced
            Select Case materialcode
                Case "C"
                    sematerialname = "C (SP01-A) / Cu DHP R220"
                    'Cu_DHP_R220 / CW024A
                Case "D"
                    sematerialname = "C (SP01-A1) / CuFe2P (K65)"
                    'CuFe2P / CW107C (K65)
                Case "X"
                    sematerialname = "X (SP01-C) / Cu-DHP R220"
                    'Cu_DHP_R220 / CW024A
                Case "R"
                    sematerialname = "R (SP01-B1) / Cu-DHP R220"
                    'Cu_DHP_R220 / CW024A
                Case "K"
                    sematerialname = "C (SP01-A1) / CuFe2P (K65)"
                    'CuFe2P / CW107C (K65)
                Case "F"
                    sematerialname = "F (SP02-A) / P195TR2"
                Case "G"
                    sematerialname = "AlMg3 (SP12)/EN AW-5754/3.3535.10"
                    'SP12_3.3535.10
                Case "S"
                    sematerialname = "Steel galv (SP15-1)/ DX51D+Z / 1.0917+Z200-N-A"
                    'SP15-1_DX51D+Z
                Case "V"
                    sematerialname = "V (SP03-1) / V2A"
                    'X2CrNi18-9 / 1.4307
                Case "W"
                    sematerialname = "W (SP03-2) / V4A"
                    'X2CrNiMo17-13-2 / 1.4404
                Case "SG"
                    sematerialname = "AlMg3 (SP12)/EN AW-5754/3.3535.10"
                    'SP12_3.3535.10
                Case "SS"
                    sematerialname = "Steel galv (SP15-1)/ DX51D+Z / 1.0917+Z200-N-A"
                    'SP15-1_DX51D+Z
                Case "SV"
                    sematerialname = "Stainless Steel (SP16-4)/X2CrTiNb18/1.4509"
                    'SP16-4_1.4509
                Case "SW"
                    sematerialname = "Stainless Steel (SP16-2)/X2CrNiMo17-13-2/1.4404"
                    'SP16-2_1.4404
            End Select

            objmattable = seapp.GetMaterialTable
            If filetype = "part" Then
                partdoc = TryCast(obj, SolidEdgePart.PartDocument)
                objmattable.ApplyMaterial(partdoc, sematerialname)
            Else
                psmdoc = TryCast(obj, SolidEdgePart.SheetMetalDocument)
                objmattable.ApplyMaterial(psmdoc, sematerialname)
            End If

        Catch ex As Exception
            If materialcode = "S" Then
                Try
                    sematerialname = "Steel galv (SP15-1)/ DX51D+Z / 1.0226+Z275-N-A"
                    objmattable = seapp.GetMaterialTable
                    objmattable.ApplyMaterial(obj, sematerialname)
                Catch ex2 As Exception
                    General.CreateLogEntry(ex.ToString)
                End Try
            Else
                General.CreateLogEntry(ex.ToString)
            End If
        End Try

    End Sub

    Shared Function GetFeature(partdoc As SolidEdgePart.PartDocument, featurename As String, featuretype As String) As Object
        Dim returnobj As Object = Nothing

        Select Case featuretype
            Case "cutout"
                Dim objcuts As SolidEdgePart.ExtrudedCutouts
                Dim objcut As SolidEdgePart.ExtrudedCutout = Nothing

                objcuts = partdoc.Models.Item(1).ExtrudedCutouts
                For Each cutout As SolidEdgePart.ExtrudedCutout In objcuts
                    If featurename.Contains("_") Then
                        If cutout.SystemName = featurename Then
                            objcut = cutout
                        End If
                    Else
                        If cutout.Name.Contains(featurename) Then
                            objcut = cutout
                        End If
                    End If
                Next
                returnobj = objcut

            Case "extrusion"
                Dim objextrudes As SolidEdgePart.ExtrudedProtrusions
                Dim objextrude As SolidEdgePart.ExtrudedProtrusion = Nothing

                objextrudes = partdoc.Models.Item(1).ExtrudedProtrusions
                For Each extrude As SolidEdgePart.ExtrudedProtrusion In objextrudes
                    If featurename.Contains("_") Then
                        If extrude.SystemName = featurename Then
                            objextrude = extrude
                        End If
                    Else
                        If extrude.Name = featurename Then
                            objextrude = extrude
                        End If
                    End If
                Next
                returnobj = objextrude
            Case "revcutout"

            Case "revprotrusion"
                Dim objrevpros As SolidEdgePart.RevolvedProtrusions
                Dim objrevpro As SolidEdgePart.RevolvedProtrusion = Nothing

                objrevpros = partdoc.Models.Item(1).RevolvedProtrusions
                For Each protrude As SolidEdgePart.RevolvedProtrusion In objrevpros
                    If featurename.Contains("_") Then
                        If protrude.SystemName = featurename Then
                            objrevpro = protrude
                        End If
                    Else
                        If protrude.Name = featurename Then
                            objrevpro = protrude
                        End If
                    End If
                Next
                returnobj = objrevpro
            Case Else

        End Select

        Return returnobj
    End Function

    Shared Function GetSideFace(partdoc As SolidEdgePart.PartDocument, featuretyp As String, featureno As Integer, sidefaceno As Integer) As SolidEdgeGeometry.Face
        Dim objmodel As SolidEdgePart.Model
        Dim objcuts As SolidEdgePart.ExtrudedCutouts
        Dim objcut As SolidEdgePart.ExtrudedCutout
        Dim objpros As SolidEdgePart.ExtrudedProtrusions
        Dim objpro As SolidEdgePart.ExtrudedProtrusion
        Dim objrevpros As SolidEdgePart.RevolvedProtrusions
        Dim objrevpro As SolidEdgePart.RevolvedProtrusion
        Dim objsidefaces As SolidEdgeGeometry.Faces = Nothing
        Dim objface As SolidEdgeGeometry.Face = Nothing

        Try
            objmodel = partdoc.Models.Item(1)

            Select Case featuretyp
                Case "cutout"
                    objcuts = objmodel.ExtrudedCutouts
                    objcut = objcuts.Item(featureno)
                    objsidefaces = objcut.SideFaces
                Case "protrusion"
                    objpros = objmodel.ExtrudedProtrusions
                    objpro = objpros.Item(featureno)
                    objsidefaces = objpro.SideFaces
                Case "revprotrusion"
                    objrevpros = objmodel.RevolvedProtrusions
                    objrevpro = objrevpros.Item(featureno)
                    objsidefaces = objrevpro.SideFaces
            End Select

            If objsidefaces IsNot Nothing Then
                objface = objsidefaces.Item(sidefaceno)
            End If

        Catch ex As Exception
            objface = Nothing
        End Try

        Return objface
    End Function

    Shared Function GetRefPlane(partdoc As SolidEdgePart.PartDocument, Optional planenumber As Integer = 1, Optional planename As String = "") As SolidEdgePart.RefPlane
        Dim objrefplanes As SolidEdgePart.RefPlanes
        Dim objrefplane As SolidEdgePart.RefPlane = Nothing

        objrefplanes = partdoc.RefPlanes
        If planename = "" Then
            objrefplane = objrefplanes.Item(planenumber)
        Else
            For Each refplane As SolidEdgePart.RefPlane In objrefplanes
                If refplane.Global = True Then
                    If refplane.Name = planename Then
                        objrefplane = refplane
                    End If
                End If
            Next
        End If

        Return objrefplane
    End Function

    Shared Function GetFaceNo(objocc As SolidEdgeAssembly.Occurrence, fin As String, orientation As String, onebranch As Boolean) As Integer
        Dim faceno As Integer = 0
        Dim partdoc As SolidEdgePart.PartDocument
        Dim partmodels As SolidEdgePart.Models
        Dim partmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim objface As SolidEdgeGeometry.Face
        Dim minrange(2), maxrange(2), xmin, ymin, zmin, xmax, ymax, zmax As Double
        Dim diameter As Double = GNData.GetTubeDiameter(fin)
        Dim i As Integer = 0
        Dim loopexit As Boolean = False

        Try
            partdoc = objocc.PartDocument
            partmodels = partdoc.Models
            partmodel = partmodels.Item(1)
            objbody = partmodel.Body
            objloops = objbody.Loops

            If onebranch Then
                'check the diameter from variable table
                diameter = GetSetVariableValue("Diameter", partdoc.Variables, "get")
            End If

            Do
                objloop = objloops(i)
                objface = objloop.Face

                If objface.GeometryForm = 10 Then
                    objface.GetExactRange(minrange, maxrange)
                    xmin = Math.Round(minrange(0) * 1000, 2)
                    ymin = Math.Round(minrange(1) * 1000, 2)
                    zmin = Math.Round(minrange(2) * 1000, 2)
                    xmax = Math.Round(maxrange(0) * 1000, 2)
                    ymax = Math.Round(maxrange(1) * 1000, 2)
                    zmax = Math.Round(maxrange(2) * 1000, 2)
                    If orientation = "reverse" Then
                        If Math.Abs(zmin) = Math.Abs(zmax) And ymax <> ymin And Math.Round(xmax, 3) <> diameter / 2 And Math.Abs(xmin) <> Math.Abs(xmax) Then
                            faceno = i + 1
                            loopexit = True
                        End If
                    Else
                        If Math.Abs(zmax) = Math.Abs(zmin) And ymax <> ymin And Math.Abs(xmax) = diameter / 2 And Math.Abs(xmin) = diameter / 2 Then
                            faceno = i + 1
                            loopexit = True
                        End If
                    End If

                End If
                i += 1
                If i = objloops.Count Then
                    loopexit = True
                End If
            Loop Until loopexit

        Catch ex As Exception

        End Try

        Return faceno
    End Function

    Shared Function GetSetVariableValue(varname As String, objvariables As SolidEdgeFramework.Variables, opmode As String, Optional newvalue As Double = 1, Optional multiplier As Double = 1000) As Double
        Dim objvariable As SolidEdgeFramework.variable
        Dim objvalue As Double = 0

        Try
            If opmode = "add" Then
                objvariables.Add(varname, newvalue / multiplier)
            Else
                objvariable = objvariables.Item(varname)
                If opmode.ToLower = "get" Then
                    If objvariable.UnitsType = 2 Then
                        objvalue = Math.Round(objvariable.Value * 180 / Math.PI, 1)
                    Else
                        objvalue = Math.Round(objvariable.Value * multiplier, 4)
                    End If
                ElseIf opmode.ToLower = "set" Then
                    objvariable.Value = newvalue / multiplier
                    objvalue = newvalue / multiplier
                End If
            End If
        Catch ex As Exception
            If varname = "WKL" Then
                objvalue = GetSetVariableValue("WKL1", objvariables, opmode)
            End If
        End Try

        Return objvalue
    End Function

    Shared Function CreateNewBow(workspace As String, bowkey As String, tempid As String, fin As String, material As String, pressure As Integer, fordefrost As Boolean) As String
        Dim type, length, diameter, wallthickness, L1, varvalues(), rawlength As Double
        Dim splittedkey(), materialcode, currentfilename, counterstring, newfilename As String
        Dim bowid As String = ""
        Dim filecounter As Integer = 1
        Dim tempfile, erpcode, tempdft As String
        Dim radius As Double
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objvariables As SolidEdgeFramework.Variables
        Dim varnames As New List(Of String) From {"CTDiameter", "WallThickness", "ABV", "L1"}
        Dim bowtube As New TubeData

        Try
            splittedkey = bowkey.Split({"\"}, 0)
            L1 = CDbl(splittedkey(0))
            type = CDbl(splittedkey(2))
            length = CDbl(splittedkey(1))

            If L1 < 60 And tempid.Contains(Library.TemplateParts.BOW9) Then
                L1 = 60
            End If

            tempfile = General.GetFullFilename(workspace, tempid, ".par")
            tempdft = General.GetFullFilename(workspace, tempid, ".dft")
            If tempfile = "" Or tempdft = "" Then
                WSM.CheckoutCircs(tempid, General.currentjob.OrderDir, workspace, General.batfile)
                General.WaitForFile(workspace, tempid, "par", 100)
                tempfile = General.GetFullFilename(workspace, tempid, "par")
            End If

            General.seapp.Documents.Open(tempfile)
            General.seapp.DoIdle()

            partdoc = General.seapp.ActiveDocument

            currentfilename = partdoc.Name
            objvariables = partdoc.Variables

            'Gather data
            diameter = GNData.GetTubeDiameter(fin)
            materialcode = GNData.GetMaterialcode(material, "bow")
            wallthickness = Database.GetTubeThickness("Bow", diameter, materialcode, pressure)
            varvalues = {diameter, wallthickness, length, L1}

            'Assign material
            materialcode = GNData.GetMaterialcode(material, "tube")
            SetMaterial(General.seapp, partdoc, materialcode, "part")

            'Assign general parameters (diameter, wallthickness, abv, length)
            For i As Integer = 0 To varvalues.Count - 1
                GetSetVariableValue(varnames(i), objvariables, "set", varvalues(i))
            Next

            If type = 9 Then
                'Change L2 value to diameter value
                If fordefrost Then
                    diameter = 18
                    GetSetCustomProp(partdoc, "", "Brine Defrost", "write")
                End If
                GetSetVariableValue("L2", objvariables, "set", diameter)
                'Iterate through Korrekturfaktor 
                'Check the values hv_L1 = L1, hv_L2 > Ø, hv_Radius > Ø/2
                ChangeType9(L1, diameter, objvariables)
            Else
                'check if radius > abv / 2
                radius = GetSetVariableValue("Radius", objvariables, "get")
                If radius * 2 > length Then
                    GetSetVariableValue("Radius", objvariables, "set", length / 2)
                End If
                'needed for more than 1 new bow → would apply the radius of the first new bow instead of default value 25mm
                If length >= 50 Then
                    GetSetVariableValue("Radius", objvariables, "set", 25)
                End If
            End If

            'Create new filename
            newfilename = currentfilename.Replace("-", "1")
            newfilename = workspace + "\" + newfilename

            Do
                If File.Exists(newfilename) Then
                    filecounter += 1
                    counterstring = filecounter.ToString + ".par"
                    newfilename = newfilename.Replace((filecounter - 1).ToString + ".par", counterstring)
                End If
            Loop Until File.Exists(newfilename) = False Or filecounter = 30

            'do the property thing / key as property?
            With bowtube.TubeFile
                .AGPno = WSM.GetAGPNumber("Rohrbogen", materialcode, False)
                .Orderno = General.currentjob.OrderNumber
                .Orderpos = General.currentjob.OrderPosition
                .Projectno = General.currentjob.ProjectNumber
                .Plant = General.currentjob.Plant
                .CDB_Zusatzbenennung = fin + "/" + materialcode + "/" + bowkey
                .CDB_Material = WSM.CDB_Material(materialcode)
                .CDB_de = "Rohrbogen"
                .CDB_en = "Tube bend"
            End With

            erpcode = Database.GetTubeERP(diameter, "Bow", pressure, materialcode)
            GetSetCustomProp(partdoc, "Raw_Material", erpcode, "write")
            rawlength = Calculation.GetTubeLength(type, partdoc, diameter)
            GetSetCustomProp(partdoc, "Raw_Length", rawlength.ToString.Replace(",", "."), "write")

            If fordefrost Then
                bowtube.TubeFile.CDB_z_Bemerkung = "Brine Defrost"
                GetSetCustomProp(partdoc, "CDB_z_bemerkung", bowtube.TubeFile.CDB_z_Bemerkung, "write")
            End If

            WriteCostumProps(partdoc, bowtube.TubeFile)

            partdoc.SaveAs(newfilename)

            bowid = partdoc.Name
            bowid = bowid.Replace(".par", "")

            'create the drawing
            SEDraft.CreateDummyDrawing(General.GetFullFilename(workspace, tempid + "0001-", ".dft"), newfilename.Replace(".par", ".dft"), bowtube.TubeFile)

            General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)

            General.ReleaseObject(partdoc)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return bowid
    End Function

    Shared Sub ChangeType9(L1 As Double, diameter As Double, objvariables As SolidEdgeFramework.Variables)
        Dim hv_L1, hv_L2, hv_radius, kf, deltavalue, L3 As Double
        Dim loopcount As Integer = 0
        Dim loopexit As Boolean = False

        L3 = L1 / 2 - 10
        GetSetVariableValue("L3", objvariables, "Set", newvalue:=L3)

        Do
            'First check for radius, if it's too low the model will fail
            hv_radius = GetSetVariableValue("hv_radius", objvariables, "Get")
            If hv_radius - 1 < diameter / 2 Then
                deltavalue = diameter / 2 - hv_radius
                kf = GetSetVariableValue("Korrekturfaktor", objvariables, "Get", multiplier:=1)
                kf = Math.Abs(deltavalue) / 5 + kf
                GetSetVariableValue("Korrekturfaktor", objvariables, "Set", kf, 1)
            Else
                'Second check for L2, heating rod must pass the bow 
                hv_L2 = GetSetVariableValue("hv_L2", objvariables, "Get")
                If hv_L2 < diameter Then
                    kf = GetSetVariableValue("Korrekturfaktor", objvariables, "Get", multiplier:=1)
                    Do
                        kf -= 0.05
                        kf = GetSetVariableValue("Korrekturfaktor", objvariables, "Set", kf, 1)
                        hv_L2 = GetSetVariableValue("hv_L2", objvariables, "Get")
                    Loop Until hv_L2 > diameter Or kf < 1.8
                Else
                    'Last check for L1, minimum length in case another bow is below
                    hv_L1 = GetSetVariableValue("hv_L1", objvariables, "Get")
                    If hv_L1 < L1 - 1 Then
                        'Increase L3
                        L3 = GetSetVariableValue("L3", objvariables, "Get")
                        Do
                            L3 += 1
                            L3 = GetSetVariableValue("L3", objvariables, "Set", newvalue:=L3)
                        Loop Until hv_L1 > L1
                    Else
                        loopexit = True
                    End If
                End If
            End If

            loopcount += 1
        Loop Until loopexit Or loopcount > 9

    End Sub

    Shared Function CreatePatternSketch(asmdoc As SolidEdgeAssembly.AssemblyDocument, quantity As Integer, pitchx As Double, pitchy As Double, circsize As Double(), alignment As String) As SolidEdgePart.Profile
        Dim objLayouts As SolidEdgeAssembly.Layouts = asmdoc.Layouts
        Dim objLayout As SolidEdgeAssembly.Layout
        Dim objProfile As SolidEdgePart.Profile = Nothing

        Try
            objLayout = objLayouts.Add(asmdoc.AsmRefPlanes.Item(3))
            objProfile = objLayout.Profile
            If alignment = "horizontal" Then
                objProfile.RectangularPatterns2d.Add(0, 0, 1, 1, 2 * Math.PI, SolidEdgeFrameworkSupport.PatternOffsetTypeConstants.sePatternFixedOffset, quantity, 1, circsize(0) / 1000 / quantity, 2 * pitchy / 1000)
            Else
                objProfile.RectangularPatterns2d.Add(0, 0, 1, 1, 0, SolidEdgeFrameworkSupport.PatternOffsetTypeConstants.sePatternFixedOffset, 1, quantity, 2 * pitchx / 1000, circsize(1) / 1000 / quantity)
            End If
            objProfile.Visible = False

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
        Return objProfile
    End Function

    Shared Function CreateNewStutzen(workspace As String, stutzenkey As String, tempid As String, material As String, fin As String, figure As Integer, pressure As Integer, abv As Double) As String
        Dim tempfile, tempdft As String
        Dim stutzenid As String = ""
        Dim filecounter As Integer = 1
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objvariables As SolidEdgeFramework.Variables
        Dim objvariable As SolidEdgeFramework.variable
        Dim varnames As New List(Of String) From {"Diameter", "WallThickness", "ABV", "L2", "WKL", "L1"}
        Dim splittedkey(), spezification, materialcode, varname, currentfilename, counterstring, newfilename, erpcode As String
        Dim diameter, wallthickness, l1, l2, angle, varvalues(), GNAngles(), figvalues(), rawlength As Double
        Dim stutzentube As New TubeData

        Try
            splittedkey = stutzenkey.Split({"\"}, 0)
            l1 = CDbl(splittedkey(1))
            l2 = CDbl(splittedkey(2))
            angle = CDbl(splittedkey(3))

            tempfile = General.GetFullFilename(workspace, tempid + "0001-", "par")
            tempdft = General.GetFullFilename(workspace, tempid + "0001-", ".dft")

            If tempfile = "" Or tempdft = "" Then
                WSM.CheckoutCircs(tempid, General.currentjob.OrderDir, workspace, General.batfile)
                General.WaitForFile(workspace, tempid, "par", 100)
                tempfile = General.GetFullFilename(workspace, tempid, "par")
            End If
            General.seapp.Documents.Open(tempfile)
            General.seapp.DoIdle()

            partdoc = General.seapp.ActiveDocument

            currentfilename = partdoc.Name
            objvariables = partdoc.Variables

            'Gather data
            diameter = GNData.GetTubeDiameter(fin)
            spezification = GNData.GetSpecification(pressure, material, "CT", fin)
            'Get Wallthickness from DB_Tubes
            wallthickness = Database.GetTubeThickness("Stub", diameter, material.Substring(0, 1), pressure)

            'Assign material
            materialcode = GNData.GetMaterialcode(material, "bow")
            SetMaterial(General.seapp, partdoc, materialcode, "part")

            erpcode = Database.GetTubeERP(diameter, "Stub", pressure, materialcode)

            If figure = 4 Then
                GNAngles = GNData.AllowedAngles("C", 4)
                'key for fig4 = l1 / abv / defaultangle
                abv = l2
                figvalues = Calculation.Fig4Parameters(l1, abv, GNAngles)
                If figvalues(0) = 0 Then
                    General.CreateLogEntry("Manual adjustment for Stutzen needed")
                End If

                l2 = figvalues(0)
                angle = figvalues(1)
            Else
                abv = 0
            End If

            'Assign general parameters
            varvalues = {diameter, wallthickness, abv, l2, angle, l1}
            For i As Integer = 0 To varnames.Count - 1
                If varvalues(i) <> 0 Then
                    varname = varnames(i)
                    objvariable = objvariables.Item(varname)
                    objvariable.Value = varvalues(i) / 1000
                End If
            Next

            'Create new filename
            newfilename = currentfilename.Replace("-", "1")
            newfilename = workspace + "\" + newfilename

            Do
                If File.Exists(newfilename) Then
                    filecounter += 1
                    counterstring = filecounter.ToString + ".par"
                    newfilename = newfilename.Replace((filecounter - 1).ToString + ".par", counterstring)
                End If
            Loop Until File.Exists(newfilename) = False Or filecounter = 20

            With stutzentube.TubeFile
                .AGPno = WSM.GetAGPNumber("Stutzen", materialcode, False)
                .Orderno = General.currentjob.OrderNumber
                .Orderpos = General.currentjob.OrderPosition
                .Projectno = General.currentjob.ProjectNumber
                .Plant = General.currentjob.Plant
                .CDB_Zusatzbenennung = fin + "/" + materialcode + "/" + stutzenkey
                .CDB_Material = WSM.CDB_Material(materialcode)
                .CDB_de = "Stutzen"
                .CDB_en = "Connection Piece"
            End With

            If General.currentjob.Plant <> "Beji" Then
                GetSetCustomProp(partdoc, "Raw_Material", erpcode, "write")
                rawlength = Calculation.GetTubeLength(figure, partdoc, diameter)
                GetSetCustomProp(partdoc, "Raw_Length", rawlength.ToString.Replace(",", "."), "write")
            End If
            WriteCostumProps(partdoc, stutzentube.TubeFile)

            partdoc.SaveAs(newfilename)

            stutzenid = partdoc.Name
            stutzenid = stutzenid.Replace(".par", "")

            'create drawing too
            SEDraft.CreateDummyDrawing(General.GetFullFilename(workspace, tempid, ".dft"), newfilename.Replace(".par", ".dft"), stutzentube.TubeFile)

            General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)

            General.ReleaseObject(partdoc)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return stutzenid
    End Function

    Shared Function GetSetCustomProp(doc As Object, propname As String, ByRef propvalue As String, opmode As String) As String
        Dim objpsets As SolidEdgeFramework.PropertySets
        Dim objgsets As SolidEdgeFramework.Properties
        Dim objprop As SolidEdgeFramework.Property

        Try
            objpsets = doc.Properties
            objgsets = objpsets.Item(4)

            If opmode = "write" Then
                Try
                    objprop = objgsets.Add(propname, propvalue.Replace(",", "."))
                Catch ex As Exception
                    objprop = objgsets.Item(propname)
                    objprop.Value = propvalue.Replace(",", ".")
                End Try
            Else
                For Each occprop As SolidEdgeFramework.Property In objgsets
                    If occprop.Name = propname Then
                        propvalue = occprop.Value
                        Exit For
                    End If
                Next
            End If

        Catch ex As Exception
            Debug.Print("Error settting value for costum prop " + propname)
        End Try

        Return propvalue
    End Function

    Shared Sub CreateHeaderTube(ByRef header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData, coil As CoilData)
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objvariables As SolidEdgeFramework.Variables
        Dim objpsets As SolidEdgeFramework.PropertySets
        Dim objgsets As SolidEdgeFramework.Properties
        Dim varname As String
        Dim nippleneeded As Boolean = False
        Dim valuelist1, valuelist2 As List(Of Double)
        Dim headercoord, angle, currentabv, position, ctdiameter, disver, hnull, vertoffset, shorten, deltaangle As Double
        Dim index, mp As Integer

        Try
            'change mp for mirrrored
            If circuit.ConnectionSide = "left" Then
                mp = -1
            Else
                mp = 1
            End If

            ctdiameter = GNData.GetTubeDiameter(circuit.FinType)
            If (header.Tube.Materialcodeletter = "V" Or header.Tube.Materialcodeletter = "W") And (circuit.CoreTube.Materialcodeletter = "V" Or circuit.CoreTube.Materialcodeletter = "W") Then
                ctdiameter += 1
            End If
            shorten = 0

            If consys.HeaderAlignment = "horizontal" Then
                valuelist1 = header.Xlist
                valuelist2 = header.Ylist
                hnull = header.Origin(0)
                headercoord = header.Origin(2)
                disver = 0
            Else
                valuelist1 = header.Ylist
                valuelist2 = header.Xlist
                headercoord = header.Origin(0)
                hnull = header.Origin(2)
                disver = header.Displacever
            End If

            header.Tube.Length = valuelist1.Max - valuelist1.Min + header.Overhangbottom + header.Overhangtop
            If General.currentunit.ModelRangeName = "GACV" Then
                If consys.HeaderAlignment = "vertical" Then
                    header.Tube.Length += disver
                    If header.Tube.HeaderType = "inlet" And header.Tube.Diameter > 60 And circuit.FinType = "F" And header.Ylist.Min = 12.5 Then
                        header.Tube.Length -= 25
                        shorten = 25
                    End If
                Else
                    header.Tube.Length -= mp * header.Displacehor
                End If
            ElseIf General.currentunit.UnitDescription = "VShape" Then
                If header.Displacever <> 0 Then
                    header.Tube.Length -= Math.Abs(disver)
                    If header.Tube.HeaderType = "inlet" Then
                        shorten = Math.Abs(disver)
                    End If
                End If
            End If

            If header.Tube.HeaderType = "inlet" Then
                If consys.InletNipples.Count > 0 AndAlso consys.InletNipples.First.Quantity > 0 Then
                    nippleneeded = True
                End If
            Else
                If consys.OutletNipples.Count > 0 AndAlso consys.OutletNipples.First.Quantity > 0 Then
                    nippleneeded = True
                End If
            End If

            If nippleneeded Then 'consys.HeaderAlignment = "vertical" Or General.currentunit.ApplicationType = "Condenser" Or General.currentunit.ApplicationType = "GCO"
                For k As Integer = 1 To header.Nippletubes
                    Calculation.NipplePosition(header, circuit, consys, k, coil.FinnedHeight)
                Next
            End If

            If header.Tube.TubeFile.Fullfilename <> "" Then
                'open the headerfile in SE
                General.seapp.Documents.Open(header.Tube.TubeFile.Fullfilename)
                General.seapp.DoIdle()
                partdoc = General.seapp.ActiveDocument
                header.Tube.FileName = partdoc.Name
                objvariables = partdoc.Variables

                'get the erpcode
                objpsets = partdoc.Properties
                objgsets = objpsets.Item(4)

                'apply basic data (Ø, wt, length, material)
                GetSetVariableValue("Diameter", objvariables, "set", header.Tube.Diameter)
                GetSetVariableValue("Length", objvariables, "set", header.Tube.Length)
                GetSetVariableValue("WallThickness", objvariables, "set", header.Tube.WallThickness)

                'create tube body
                Dim objmodels As SolidEdgePart.Models = partdoc.Models
                Dim refplane As SolidEdgePart.RefPlane = partdoc.RefPlanes.Item(1)
                Dim objsketch As SolidEdgePart.Sketch = partdoc.Sketches.Add
                Dim formulas() As String = {"Diameter", "Diameter - 2*WallThickness"}
                Dim newprofilelist As New List(Of SolidEdgePart.Profile)
                Dim objprofile As SolidEdgePart.Profile

                For i As Integer = 0 To 1
                    objprofile = objsketch.Profiles.Add(refplane)
                    objprofile.Visible = False
                    Dim objcircle As SolidEdgeFrameworkSupport.Circle2d = objprofile.Circles2d.AddByCenterRadius(0, 0, header.Tube.Diameter / 1000 - 2 * i * header.Tube.WallThickness / 1000)
                    Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objprofile.Dimensions
                    Dim circdim As SolidEdgeFrameworkSupport.Dimension = objdims.AddCircularDiameter(objcircle)
                    circdim.Constraint = True
                    circdim.Formula = formulas(i)

                    objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                    newprofilelist.Add(objprofile)
                Next

                objmodels.AddFiniteExtrudedProtrusion(2, newprofilelist.ToArray, 2, 0.1)
                Dim objexpros As SolidEdgePart.ExtrudedProtrusions = partdoc.Models.Item(1).ExtrudedProtrusions
                Dim objcyl As SolidEdgePart.ExtrudedProtrusion = objexpros(0)
                objcyl.Name = "Header"
                Dim varcount As Integer
                Dim dims(0) As Object
                objcyl.GetDimensions(varcount, dims)
                Dim lengthdim As SolidEdgeFrameworkSupport.Dimension
                lengthdim = TryCast(dims(0), SolidEdgeFrameworkSupport.Dimension)
                If lengthdim IsNot Nothing Then
                    lengthdim.Formula = "Length"
                End If

                Dim capcount As Integer = 0
                Dim needbottom As Boolean = True
                Dim needtop As Boolean = True
                If header.Tube.Materialcodeletter = "C" And General.currentunit.ModelRangeSuffix = "CP" And consys.HeaderAlignment = "horizontal" Then
                    'CP CU inlet and outlet header have the SV → requires a tube closing feature
                    If circuit.ConnectionSide = "left" Then
                        header.Tube.BottomCapNeeded = True
                        header.Tube.TopCapNeeded = False
                    Else
                        header.Tube.BottomCapNeeded = False
                        header.Tube.TopCapNeeded = True
                    End If
                    'check for tube ends  
                ElseIf GNData.CheckCaps(header.Tube.Materialcodeletter, header.Tube.Diameter, header.Tube.WallThickness, False, circuit.Pressure, "header") Then
                    header.Tube.BottomCapNeeded = True
                    header.Tube.TopCapNeeded = True
                End If

                If Not header.Tube.BottomCapNeeded Or Not header.Tube.TopCapNeeded Or (header.Tube.HeaderType = "outlet" And header.Tube.IsBrine And consys.HeaderAlignment = "horizontal") Then
                    CreateTubeEnds(partdoc, consys.HeaderAlignment, circuit.ConnectionSide, header.Tube, 1)
                End If
                If header.Tube.TopCapNeeded Then
                    capcount += 1
                End If
                If header.Tube.BottomCapNeeded Then
                    capcount += 1
                End If

                'create the angle planes
                Dim usedangles, uniqueabvs, uniqueangles As New List(Of Double)
                Dim angleplanedicc As New Dictionary(Of Double, Integer)

                For i As Integer = 0 To header.StutzenDatalist.Count - 1
                    Dim oddangle As Boolean = False
                    angle = header.StutzenDatalist(i).Angle
                    'if air direction comes from right or top, then abv switches sign
                    angle = CheckAngle(angle, header.Tube.HeaderType, consys.HeaderAlignment, header.StutzenDatalist(i).ABV)
                    'if position (y value) is > coil length, then header on the other side of the coil → change angle
                    If Math.Abs(header.Origin(1)) > coil.FinnedLength Or header.Tube.IsBrine Then
                        angle = 180 - angle
                        oddangle = True
                    End If

                    If uniqueabvs.IndexOf(header.StutzenDatalist(i).ABV) = -1 And header.StutzenDatalist(i).SpecialTag = "" Then
                        varname = "Angle" + CStr(uniqueabvs.Count + 1)
                        CreateAngRefPlane(partdoc, 3, angle, varname + "plane")
                        uniqueabvs.Add(header.StutzenDatalist(i).ABV)
                        If oddangle Then
                            If angle > 180 Then
                                uniqueangles.Add(Math.Abs(180 - angle))
                            Else
                                uniqueangles.Add(180 - angle)
                            End If
                        Else
                            uniqueangles.Add(angle)
                        End If
                    End If
                Next

                'set the cutout features
                For i As Integer = 0 To valuelist1.Count - 1
                    vertoffset = 0
                    'get abv
                    currentabv = Math.Round(headercoord - valuelist2(i), 2)

                    'get index from abv list
                    index = uniqueabvs.IndexOf(currentabv)

                    If index = -1 Or uniqueangles.IndexOf(header.StutzenDatalist(i).Angle) = -1 Or header.StutzenDatalist(i).SpecialTag <> "" Then
                        If header.StutzenDatalist(i).SpecialTag = "" Then
                            If General.currentunit.ModelRangeName = "GACV" And index = -1 Then
                                'abv was changed → currently only possible for GACV AP → abv = -55.9 / different for CU CP
                                If header.Tube.Materialcodeletter = "C" Then
                                    If Math.Abs(currentabv) = 25 Then
                                        currentabv = -29.7
                                    End If
                                Else
                                    If Math.Abs(currentabv) = 50 Then
                                        currentabv = -55.9
                                    End If
                                End If
                                index = uniqueabvs.IndexOf(currentabv)
                            Else
                                If header.StutzenDatalist(i).Angle = 0 Or header.StutzenDatalist(i).Angle = 180 Then
                                    index = uniqueangles.IndexOf(0)
                                Else
                                    Dim tempangle As Integer = CheckAngle(header.StutzenDatalist(i).Angle, header.Tube.HeaderType, consys.HeaderAlignment, header.StutzenDatalist(i).ABV)
                                    If -header.StutzenDatalist(i).Angle = tempangle Then
                                        index = uniqueangles.IndexOf(-header.StutzenDatalist(i).Angle)
                                    Else
                                        If header.StutzenDatalist(i).ABV < 0 Then
                                            angle = -Math.Abs(header.StutzenDatalist(i).Angle)
                                        Else
                                            angle = Math.Abs(header.StutzenDatalist(i).Angle)
                                        End If
                                        CreateAngRefPlane(partdoc, 3, angle, "Angle" + CStr(uniqueangles.Count).ToString + "plane")
                                        uniqueangles.Add(header.StutzenDatalist(i).Angle)
                                        uniqueabvs.Add(currentabv)
                                        index = uniqueangles.IndexOf(header.StutzenDatalist(i).Angle)
                                    End If
                                End If
                            End If
                        Else
                            If uniqueangles.IndexOf(header.StutzenDatalist(i).Angle) = -1 Then
                                uniqueangles.Add(header.StutzenDatalist(i).Angle)
                                If header.StutzenDatalist(i).ABV < 0 Then
                                    angle = -Math.Abs(header.StutzenDatalist(i).Angle)
                                Else
                                    angle = Math.Abs(header.StutzenDatalist(i).Angle)
                                End If
                                If Math.Abs(header.Origin(1)) > coil.FinnedLength Or header.Tube.IsBrine Then
                                    angle = 180 - angle
                                End If
                                CreateAngRefPlane(partdoc, 3, angle, "Angle" + CStr(uniqueangles.Count).ToString + "plane")
                                index = uniqueangles.Count - 1
                            Else
                                index = uniqueangles.IndexOf(header.StutzenDatalist(i).Angle)
                            End If
                        End If
                    End If

                    'get the refplane number with that index, starting at item(4)
                    index += 4

                    If header.StutzenDatalist(i).SpecialTag <> "" Then
                        If header.StutzenDatalist(i).SpecialTag = "s2star" Then
                            position = valuelist1(i) + header.Overhangbottom - valuelist1.Min + 25 * mp
                        Else
                            If header.Tube.HeaderType = "inlet" Then
                                If General.currentunit.ModelRangeName = "GACV" Then
                                    If header.StutzenDatalist(i).SpecialTag.Contains("s1star") Then
                                        position = header.Tube.Length - header.Overhangtop
                                    Else
                                        position = header.Overhangbottom
                                    End If
                                Else
                                    position = header.Overhangbottom + header.StutzenDatalist(i).HoleOffset
                                End If
                            Else
                                If General.currentunit.ModelRangeName = "GACV" Then
                                    position = valuelist1(i) + header.Overhangbottom - valuelist1.Min - shorten
                                    If header.StutzenDatalist(i).SpecialTag = "sOutT45r1" Then
                                        position -= 25
                                    ElseIf header.StutzenDatalist(i).SpecialTag = "s4starN" Then
                                        position += 25.6
                                    ElseIf header.StutzenDatalist(i).SpecialTag = "s4starE1" Then
                                        position -= 50
                                    ElseIf header.StutzenDatalist(i).SpecialTag = "s4starE2" Then
                                        position -= 25
                                    ElseIf header.StutzenDatalist(i).SpecialTag = "s4starAP" Then
                                        position += 5
                                    ElseIf header.StutzenDatalist(i).SpecialTag = "s4starF1" Then
                                        position -= 25
                                    Else
                                        position = header.Tube.Length - header.Overhangtop
                                    End If
                                Else
                                    position = header.Tube.Length - header.Overhangtop - header.StutzenDatalist(i).HoleOffset
                                End If
                            End If
                        End If
                    Else
                        position = valuelist1(i) + header.Overhangbottom - valuelist1.Min - shorten
                    End If

                    If header.Tube.Diameter = 60.3 And consys.HeaderAlignment = "horizontal" And General.currentunit.ModelRangeName = "GACV" And header.Tube.HeaderType = "inlet" And valuelist2(i) = valuelist2.Min Then
                        vertoffset = -10
                    End If

                    'create the feature
                    SetCutoutFeature(partdoc, index, ctdiameter, position, vertoffset:=vertoffset, cuttype:="")
                Next

                'nipple cutouts
                If header.Nipplepositions.Count > 0 Then

                    If header.Tube.HeaderType = "outlet" Then
                        deltaangle = consys.OutletNipples.First.Angle
                    Else
                        deltaangle = consys.InletNipples.First.Angle
                    End If

                    CreateNippleRefPlane(partdoc, header, circuit, consys, deltaangle, "NippleRefplane")

                    If header.Tube.Materialcodeletter = "C" Then
                        If header.Tube.HeaderType = "inlet" Then
                            CreateNippleHoles(partdoc, header, circuit, consys.InletNipples.First.Diameter, consys.HeaderAlignment, deltaangle)
                        Else
                            CreateNippleHoles(partdoc, header, circuit, consys.OutletNipples.First.Diameter, consys.HeaderAlignment, deltaangle)
                        End If
                    ElseIf consys.OutletNipples.First.Materialcodeletter = "D" Then
                        If header.Tube.HeaderType = "inlet" Then
                            CreateNippleHoles(partdoc, header, circuit, GNData.GetNippleDiameterK65(consys.InletNipples.First.Diameter), consys.HeaderAlignment, deltaangle)
                        Else
                            CreateNippleHoles(partdoc, header, circuit, GNData.GetNippleDiameterK65(consys.OutletNipples.First.Diameter), consys.HeaderAlignment, deltaangle)
                        End If
                    Else
                        Dim profilside As Integer = 1
                        mp = -1
                        If circuit.ConnectionSide = "left" And General.currentunit.UnitDescription <> "VShape" Then
                            profilside = 2
                            mp = 1
                        End If
                        If header.Tube.HeaderType = "outlet" Then
                            If header.Tube.Diameter = consys.OutletNipples.First.Diameter Then
                                CreateSimilarHeader(partdoc, header, circuit, mp, profilside, consys.HeaderAlignment, deltaangle)
                            Else
                                For i As Integer = 0 To header.Nipplepositions.Count - 1
                                    CreateProjections(partdoc, CreateNippleCutSketch(partdoc, consys.OutletNipples.First.Diameter, consys.OutletNipples.First.WallThickness, header.Nipplepositions(i), i + 1), partdoc.RefPlanes.Item("NippleRefplane"), 2)
                                    CreateBlueSurf(partdoc, True)
                                    SubtractBody(partdoc, SolidEdgePart.SESubtractDirection.igSubtractDirectionNone)
                                Next
                            End If
                        Else
                            If header.Tube.Diameter = consys.InletNipples.First.Diameter Then
                                CreateSimilarHeader(partdoc, header, circuit, mp, profilside, consys.HeaderAlignment, deltaangle)
                            Else
                                For i As Integer = 0 To header.Nipplepositions.Count - 1
                                    CreateProjections(partdoc, CreateNippleCutSketch(partdoc, consys.InletNipples.First.Diameter, consys.InletNipples.First.WallThickness, header.Nipplepositions(i), i + 1), partdoc.RefPlanes.Item("NippleRefplane"), 2)
                                    CreateBlueSurf(partdoc, True)
                                    SubtractBody(partdoc, SolidEdgePart.SESubtractDirection.igSubtractDirectionNone)
                                Next
                            End If
                        End If
                    End If
                End If

                If header.Tube.IsBrine Then
                    'cutouts for vents
                    CreateBrineVents(partdoc, header, consys.HeaderAlignment)
                    header.Ventsize = GNData.GetBrineVentSize(header.Tube.Materialcodeletter, header.Tube.Diameter)
                ElseIf consys.HasHotgas Then
                    If header.Tube.Diameter > 21.3 Then
                        Dim HGpos As Double
                        If consys.HeaderAlignment = "horizontal" Then
                            Dim hotgasangle As Double = 0
                            If circuit.ConnectionSide = "right" Then
                                HGpos = header.Overhangbottom - 42
                            Else
                                HGpos = header.Tube.Length - (header.Overhangtop - 42)
                            End If
                            If header.Tube.HeaderType = "outlet" And consys.HotGasData.Headertype = "outlet" Then
                                If header.Tube.Diameter > 60 Then
                                    hotgasangle = 45
                                ElseIf header.Tube.Diameter = 42.4 Then
                                    hotgasangle = 38
                                Else
                                    hotgasangle = 35
                                End If
                                'cutout between first stutzen and tube bottom, might be angled
                                Dim planeno As Integer = CreateAngRefPlane(partdoc, 2, 180 - hotgasangle, "HotgasRefplane")
                                SetCutoutFeature(partdoc, planeno, consys.HotGasConnectionDiameter, HGpos, direction:="right", cuttype:="HG", newname:="Hotgas")
                                consys.HotGasData.Angle = hotgasangle
                                consys.HotGasData.Diameter = consys.HotGasConnectionDiameter
                            ElseIf header.Tube.HeaderType = "inlet" And consys.HotGasData.Headertype = "inlet" Then
                                Dim planeno As Integer = CreateAngRefPlane(partdoc, 2, 180, "HotgasRefplane")
                                SetCutoutFeature(partdoc, planeno, consys.HotGasConnectionDiameter, HGpos, direction:="right", cuttype:="HG", newname:="Hotgas")
                                consys.HotGasData.Angle = 180
                                consys.HotGasData.Diameter = consys.HotGasConnectionDiameter
                            End If
                        ElseIf header.Tube.HeaderType = "outlet" And consys.VType = "P" Then
                            'cutout in the middle of the header, only for AP / CP
                            If circuit.ConnectionSide = "left" Then
                                HGpos = (header.Tube.Length - header.Overhangbottom - header.Overhangtop) / 2 + header.Overhangtop
                            Else
                                HGpos = header.Tube.Length - ((header.Tube.Length - header.Overhangbottom - header.Overhangtop) / 2 + header.Overhangbottom)
                            End If
                            SetCutoutFeature(partdoc, 3, consys.HotGasConnectionDiameter, HGpos, direction:="right", cuttype:="HG", newname:="Hotgas")
                            consys.HotGasData.Angle = 90
                            consys.HotGasData.Diameter = consys.HotGasConnectionDiameter
                        End If
                    End If
                Else
                    If circuit.Pressure < 17 And General.currentunit.ModelRangeSuffix <> "OD" Then
                        If General.currentunit.ModelRangeName = "GACV" Then
                            If header.Tube.HeaderType = "inlet" Then
                                If GNData.GACVInletVent(consys.InletHeaders.First, circuit.FinType) Then
                                    consys.InletHeaders.First.Ventposition = Calculation.VentGACVPos(consys.InletHeaders.First, circuit, coil.NoRows)
                                End If
                            Else
                                consys.OutletHeaders.First.Ventposition = Calculation.VentGACVPos(consys.OutletHeaders.First, circuit, coil.NoRows)
                            End If
                            'vents, important are angle and direction depending of modelrange and header alignment
                            header.Ventsize = GNData.GetGACVVentsize(header, circuit, coil.NoRows)
                        Else
                            header.Ventposition = Calculation.VentWSPos(header, circuit.CoreTube.Diameter)
                            If consys.Valvesize <> "" Then
                                header.Ventsize = consys.Valvesize
                            Else
                                header.Ventsize = GNData.GetDryCoolerVentsize(header.Tube.Materialcodeletter, header.Tube.Diameter)
                            End If
                        End If

                        If header.Ventposition <> 0 Then
                            CreateVents(partdoc, header, circuit, coil)
                        End If
                    End If
                End If

                If consys.VType = "X" Then
                    'pressure pipe
                    SetCutoutFeature(partdoc, 3, 6, header.Tube.Length / 2, direction:="right", cuttype:="DAL", newname:="DAL")
                End If

                If header.Tube.SVPosition(0) = "header" And header.Tube.SVPosition(1) = "perp" Then
                    Dim svpos As Double
                    Dim svdir As String
                    Dim planeno As Integer = 2
                    If consys.HeaderAlignment = "horizontal" Then
                        svpos = (header.Nipplepositions(0) - header.Overhangbottom) / 2 + header.Overhangbottom
                        svdir = "left"
                    Else
                        svdir = "right"
                        If General.currentunit.UnitDescription = "VShape" Then
                            svpos = header.Tube.Length - ((header.Nipplepositions(0) - header.Overhangtop) / 2 + header.Overhangtop)
                            planeno = 3
                        Else
                            svpos = header.Tube.Length - 50
                        End If
                    End If
                    'only Condenser inlet
                    SetCutoutFeature(partdoc, planeno, 8, svpos, direction:=svdir, cuttype:="SV", newname:="svhole")
                End If

                'raw length and material
                Dim closingoffset As Double
                If header.Tube.SVPosition(1) = "axial" Then
                    closingoffset = capcount * GNData.TubeClosingOffset(header.Tube.Diameter, False, General.currentjob.Plant, header.Tube.WallThickness) + (2 - capcount) * GNData.TubeClosingOffset(header.Tube.Diameter, True, General.currentjob.Plant, header.Tube.WallThickness)
                Else
                    closingoffset = capcount * GNData.TubeClosingOffset(header.Tube.Diameter, False, General.currentjob.Plant, header.Tube.WallThickness) + (2 - capcount) * GNData.TubeClosingOffset(header.Tube.Diameter, False, General.currentjob.Plant, header.Tube.WallThickness)
                End If

                header.Tube.RawLength = (header.Tube.Length + closingoffset) / 1000

                header.Tube.RawMaterial = Database.GetTubeERP(header.Tube.Diameter, "Headertube", circuit.Pressure, header.Tube.Materialcodeletter)

                GetSetCustomProp(partdoc, "Raw_Material", header.Tube.RawMaterial, "write")
                GetSetCustomProp(partdoc, "Raw_Length", header.Tube.RawLength, "write")

                partdoc.Save()
                General.seapp.DoIdle()
                partdoc.Close()
                Threading.Thread.Sleep(1000)

                General.ReleaseObject(partdoc)

            Else
                General.CreateLogEntry("Missing headerfile")
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try


    End Sub

    Shared Sub CreateTubeEnds(partdoc As SolidEdgePart.PartDocument, alignment As String, conside As String, tube As TubeData, planeno As Integer)
        Dim objexpros As SolidEdgePart.ExtrudedProtrusions
        Dim objexpro, objtop, objbottom As SolidEdgePart.ExtrudedProtrusion
        Dim objcuts As SolidEdgePart.ExtrudedCutouts
        Dim objcut As SolidEdgePart.ExtrudedCutout
        Dim refplane As SolidEdgePart.RefPlane
        Dim objcircles As SolidEdgeFrameworkSupport.Circles2d
        Dim objrelations As SolidEdgeFrameworkSupport.Relations2d
        Dim objprofile As SolidEdgePart.Profile
        Dim objplane As SolidEdgePart.RefPlane
        Dim ventdiameter As Double
        Dim featurename As String

        Try
            refplane = partdoc.RefPlanes.Item(planeno)

            objexpros = partdoc.Models.Item(1).ExtrudedProtrusions
            objexpro = objexpros(1)

            objplane = partdoc.RefPlanes.AddParallelByDistance(refplane, tube.Length / 1000, SolidEdgePart.ReferenceElementConstants.igNormalSide, Local:=True)
            objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            objprofile.Visible = False

            objprofile.Circles2d.AddByCenterRadius(0, 0, tube.Diameter / 2000)

            'at origin = "bottom"
            objbottom = objexpros.AddFinite(objprofile, 1, 2, tube.WallThickness / 1000)
            objbottom.Name = "BottomCap"

            '"top"
            objtop = objexpros.AddFromTo(objprofile, 1, refplane, objplane)
            objtop.SetFromFaceOffsetData(objplane, 1, tube.WallThickness / 1000)
            objtop.Name = "TopCap"

            If alignment = "horizontal" And General.currentunit.ApplicationType = "Evaporator" And Not (tube.HeaderType = "inlet" And tube.IsBrine) Then
                objcuts = partdoc.Models.Item(1).ExtrudedCutouts

                'add a profile 
                objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
                objprofile.Visible = False

                objcircles = objprofile.Circles2d
                objrelations = objprofile.Relations2d

                If (tube.IsBrine And conside = "left") Or (Not tube.IsBrine And conside = "right") Then
                    SetCutoutFeature(partdoc, 1, 8, 0, direction:="right", newname:="svhole")
                Else
                    ventdiameter = 8
                    featurename = "svhole"
                    'objcircles.AddByCenterRadius((tube.WallThickness + 1) / 1000, 0, ventdiameter / 2000) 'x = tube.diameter/2?
                    'objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

                    'objcut = objcuts.AddFromTo(objprofile, 1, refplane, objplane)

                    'objcut.SetFromFaceOffsetData(refplane, 1, -tube.Length + tube.WallThickness)

                    objcircles.AddByCenterRadius(0, 0, ventdiameter / 2000) 'x = tube.diameter/2?
                    objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                    objcut = objcuts.AddFromTo(objprofile, 1, refplane, objplane)

                    objcut.SetFromFaceOffsetData(refplane, 2, (tube.WallThickness + 1) / 1000)

                    objcut.Name = featurename
                End If
            End If

            objbottom.Suppress = tube.BottomCapNeeded
            If tube.TubeType = "nipple" Then
                objbottom.Suppress = True
            End If
            objtop.Suppress = tube.TopCapNeeded

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateNippleHoles(partdoc As SolidEdgePart.PartDocument, header As HeaderData, circuit As CircuitData, diameter As Double, headeralignment As String, deltaangle As Double)
        Dim parentplaneno As Integer = 3
        Dim planeno As Integer
        Dim offset As Double
        Dim direction As String = "left"

        Try
            'deltaangle = CSData.GetNippleAngle(header.Tube, circuit, headeralignment)
            'If General.currentunit.ApplicationType = "Condenser" And circuit.Pressure > 100 Then
            '    'deltaangle -= 180
            'End If

            planeno = CreateAngRefPlane(partdoc, parentplaneno, deltaangle, "NippleRefplane")

            'If circuit.NoPasses = 1 Or circuit.NoPasses = 3 And circuit.Pressure < 17 Then
            '    If header.OddLocation = "front" Then
            '        direction = "left"
            '    Else
            '        direction = "right"
            '    End If
            'End If
            If header.Tube.Materialcodeletter = "C" Then
                offset = GNData.GetCutoutOffset(header.Tube.Diameter, diameter)
            End If

            Dim i As Integer = 1
            For Each pos In header.Nipplepositions
                Debug.Print("Direction: " + direction + " // Angle: " + deltaangle.ToString)
                SetNippleCutout(partdoc, planeno, diameter, pos, offset, direction:=direction, newname:="Nippletube" + i.ToString)
                i += 1
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetCutoutFeature(partdoc As SolidEdgePart.PartDocument, planeno As Integer, diameter As Double, position As Double, Optional direction As String = "left", Optional cuttype As String = "CT",
                                Optional newname As String = "", Optional vertoffset As Double = 0)
        Dim refplanes As SolidEdgePart.RefPlanes
        Dim refplane As SolidEdgePart.RefPlane
        Dim objprofile As SolidEdgePart.Profile
        Dim objcircles As SolidEdgeFrameworkSupport.Circles2d
        Dim objrelations As SolidEdgeFrameworkSupport.Relations2d
        Dim objmodel As SolidEdgePart.Model
        Dim objcutouts As SolidEdgePart.ExtrudedCutouts
        Dim newcutout As SolidEdgePart.ExtrudedCutout
        Dim objplaneside As SolidEdgePart.FeaturePropertyConstants
        Dim offsetside As SolidEdgePart.OffsetSideConstants

        Try
            'get the refplane
            refplanes = partdoc.RefPlanes
            refplane = refplanes.Item(planeno)

            'add a profile for the sketch 
            objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            objprofile.Visible = False

            objcircles = objprofile.Circles2d
            objrelations = objprofile.Relations2d

            'add the circle
            If cuttype = "CT" Then
                objcircles.AddByCenterRadius(position / 1000, vertoffset / 1000, diameter / 2000)
            Else
                objcircles.AddByCenterRadius(0, position / 1000, diameter / 2000)
            End If

            'select the side for the cutout 
            If direction = "left" Then
                objplaneside = SolidEdgePart.FeaturePropertyConstants.igLeft
                offsetside = SolidEdgePart.OffsetSideConstants.seOffsetLeft
            Else
                objplaneside = SolidEdgePart.FeaturePropertyConstants.igRight
                offsetside = SolidEdgePart.OffsetSideConstants.seOffsetRight
            End If
            'close the profile
            objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

            objmodel = partdoc.Models.Item(1)

            objcutouts = objmodel.ExtrudedCutouts

            'add the cutout
            newcutout = objcutouts.AddThroughNext(objprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, objplaneside)

            If newname <> "" Then
                newcutout.Name = newname
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function CheckAngle(angle As Double, headertype As String, alignment As String, abv As Double) As Double
        Dim newangle As Double = angle

        If angle >= 180 Then
            newangle = 0
        Else

            If abv < 0 Then
                newangle = -newangle
            End If
            If headertype = "inlet" And alignment = "horizontal" Then
                newangle = -newangle
            End If
            If headertype = "outlet" And alignment = "horizontal" Then
                newangle = -newangle
            End If

        End If

        Return newangle
    End Function

    Shared Sub CreateNippleRefPlane(partdoc As SolidEdgePart.PartDocument, header As HeaderData, circuit As CircuitData, consys As ConSysData, deltaangle As Double, planename As String)
        Dim parentplaneno As Integer = 3
        Dim planeno As Integer

        'deltaangle = CSData.GetNippleAngle(header.Tube, circuit, consys.HeaderAlignment)
        'If General.currentunit.ApplicationType = "Condenser" Then   'And circuit.Pressure > 100 
        '    If header.OddLocation = "back" And circuit.Pressure <= 16 Then
        '        deltaangle -= 180
        '    End If
        'End If
        'If header.Tube.HeaderType = "outlet" Then
        '    deltaangle += consys.NippleAngleOut
        'Else
        '    deltaangle += consys.NippleAngleIn
        'End If 
        planeno = CreateAngRefPlane(partdoc, parentplaneno, deltaangle, planename)
    End Sub

    Shared Function CreateAngRefPlane(partdoc As SolidEdgePart.PartDocument, parentplaneno As Integer, angle As Double, planename As String) As Integer
        Dim refplanes As SolidEdgePart.RefPlanes = partdoc.RefPlanes
        Dim newplane, parentplane, secundaryplane, refplane As SolidEdgePart.RefPlane
        Dim radangle As Double = Math.Round(angle * Math.PI / 180, 4)
        Dim secplaneno As Integer = 5 - parentplaneno
        Dim newplanenno As Integer = -1
        Dim i As Integer = 1

        Try
            For Each refplane In refplanes
                If refplane.Global Then
                    If refplane.Name = planename Then
                        newplanenno = i
                    End If
                End If
                i += 1
            Next

            If newplanenno = -1 Then
                parentplane = refplanes.Item(parentplaneno)
                secundaryplane = refplanes.Item(secplaneno)
                newplane = refplanes.AddAngularByAngle(parentplane, radangle, SolidEdgePart.ReferenceElementConstants.igNormalSide, secundaryplane, SolidEdgePart.ReferenceElementConstants.igAngular, Local:=False)
                newplane.Name = planename
                newplane.Visible = False
                refplanes = Nothing
                refplanes = partdoc.RefPlanes
                newplanenno = i
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return newplanenno
    End Function

    Shared Function GetFigure(asmdoc As SolidEdgeAssembly.AssemblyDocument, stutzenno As Integer, abv As Double) As Integer
        Dim loopcount As Integer
        Dim figure As Integer = 0
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops

        Try
            If abv = 0 Then
                figure = 8
            Else
                partdoc = asmdoc.Occurrences.Item(stutzenno).PartDocument
                objmodel = partdoc.Models.Item(1)
                objbody = objmodel.Body
                objloops = objbody.Loops
                loopcount = objloops.Count

                If loopcount = 16 Or loopcount = 20 Then
                    '16 loops only for the model with 2 legs and 1 bending 
                    figure = 5
                ElseIf loopcount = 32 Then
                    figure = 45
                Else
                    If objmodel.Thinwalls.Count > 0 Or objmodel.CopiedParts.Count > 0 Then
                        figure = 45
                    Else
                        figure = 4
                    End If
                End If

            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return figure
    End Function

    Shared Function GetStutzenAlignment(asmdoc As SolidEdgeAssembly.AssemblyDocument, stutzenno As Integer, abv As Double, header As HeaderData, ctoverhang As Double) As String
        Dim stutzenvalues(), value As Double
        Dim varnames() As String = {"L1", "L2", "WKL"}
        Dim varname, alignment As String
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objvariables As SolidEdgeFramework.Variables
        Dim objvariable As SolidEdgeFramework.variable

        alignment = "normal"
        stutzenvalues = {0, 0, 0}

        Try
            partdoc = asmdoc.Occurrences.Item(stutzenno).PartDocument
            objvariables = partdoc.Variables

            'get all parameters from the model
            For i As Integer = 0 To 2
                varname = varnames(i)
                objvariable = objvariables.Item(varname)
                If objvariable.UnitsType = 2 Then
                    value = Math.Round(objvariable.Value / Math.PI * 180, 1)
                Else
                    value = Math.Round(objvariable.Value * 1000, 2)
                End If
                stutzenvalues(i) = value
            Next

            If header.Tube.TubeType = "stutzen" Then
                'abv = a - overhang
                If Math.Abs(stutzenvalues(1) - abv) < Math.Abs(stutzenvalues(0) - abv) Then
                    alignment = "reverse"
                End If
            Else
                alignment = Calculation.GetStutzen5Alignment(stutzenvalues(0), stutzenvalues(1), Math.Abs(abv), stutzenvalues(2), header.Dim_a, header.Tube.Diameter, header.Tube.WallThickness, ctoverhang)
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return alignment
    End Function

    Shared Function GetFace(partdoc As SolidEdgePart.PartDocument, parttyp As String) As Object
        Dim objface As SolidEdgeGeometry.Face = Nothing
        Dim objmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim loopexit As Boolean
        Dim i As Integer = 0
        Dim minrange(2), maxrange(2), ymin, ymax As Double

        Try
            objmodel = partdoc.Models.Item(1)
            objbody = objmodel.Body
            objloops = objbody.Loops

            Do
                objloop = objloops(i)
                objface = objloop.Face
                If objface.GeometryForm = 9 Then
                    objface.GetRange(minrange, maxrange)
                    ymin = Math.Round(minrange(1) * 1000, 2)
                    ymax = Math.Round(maxrange(1) * 1000, 2)
                    If parttyp = "stutzen" Then
                        If ymin <> 0 And ymax <> 0 Then
                            loopexit = True
                        End If
                    Else
                        If ymin = 0 And ymax = 0 Then
                            loopexit = True
                        End If
                    End If
                End If

                i += 1
            Loop Until loopexit Or i = objloops.Count

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objface
    End Function

    Shared Function GetPlaneNo(headeralignment As String) As Integer
        Dim planeno As Integer

        If headeralignment = "vertical" Then
            planeno = 1
        Else
            planeno = 2
        End If

        Return planeno
    End Function

    Shared Sub CreateBrineVents(partdoc As SolidEdgePart.PartDocument, ByRef header As HeaderData, alignment As String)
        Dim ventdiameter, position, distance As Double
        Dim planeno As Integer
        Dim direction As String = "left"

        Try
            ventdiameter = GNData.GetVentDiameter(header.Tube.Materialcodeletter, GNData.GetBrineVentSize(header.Tube.Materialcodeletter, header.Tube.Diameter))
            If header.Tube.Materialcodeletter = "C" Then
                distance = 18
            Else
                If header.Tube.Diameter < 33 Then
                    distance = 19
                ElseIf header.Tube.Diameter > 76 Then
                    distance = 32
                Else
                    distance = 22
                End If
            End If

            If alignment = "horizontal" Then
                planeno = 3
                position = 28
            Else
                'direction changing for outlet if mirrored
                If header.Tube.HeaderType = "outlet" Then
                    planeno = 2
                Else
                    planeno = 3
                    distance = header.Tube.Length - distance
                End If
                position = distance
            End If
            position = header.Tube.Length - position
            header.Ventposition = position
            SetCutoutFeature(partdoc, planeno, ventdiameter, 0, direction:=direction, newname:="Vent", vertoffset:=position)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateVents(partdoc As SolidEdgePart.PartDocument, ByRef header As HeaderData, circuit As CircuitData, coil As CoilData)
        Dim planeno As Integer
        Dim parentplaneno As Integer = 3
        Dim deltaangle As Double
        Dim holediameter As Double

        Try
            If General.currentunit.ModelRangeName = "GACV" Then
                deltaangle = CSData.GetVentAngle(header.Tube.Diameter, circuit.FinType + coil.NoRows.ToString, header.Tube.Materialcodeletter, circuit.FinType, coil.FinnedHeight, circuit.ConnectionSide, header.Tube.HeaderType)
            Else
                If coil.Alignment = "vertical" Then
                    If header.Tube.HeaderType = "inlet" Then
                        deltaangle = 90
                    Else
                        deltaangle = -90
                    End If
                    If General.currentunit.UnitDescription = "VShape" And circuit.ConnectionSide = "right" Then
                        deltaangle *= -1
                    End If
                Else
                    If header.Tube.HeaderType = "inlet" Then
                        deltaangle = -90
                    Else
                        deltaangle = 90
                    End If
                End If
            End If

            planeno = CreateAngRefPlane(partdoc, parentplaneno, deltaangle, "Ventplane")
            holediameter = GNData.GetVentDiameter(header.Tube.Materialcodeletter, header.Ventsize)

            Debug.Print("Header: " + header.Tube.HeaderType + " // Angle: " + deltaangle.ToString)
            SetCutoutFeature(partdoc, planeno, holediameter, header.Ventposition, direction:="right", cuttype:="vent", newname:="Vent")
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetNippleCutout(partdoc As SolidEdgePart.PartDocument, planeno As Integer, diameter As Double, position As Double, offset As Double, direction As String, newname As String)
        Dim refplanes As SolidEdgePart.RefPlanes
        Dim refplane, nippletoplane As SolidEdgePart.RefPlane
        Dim objprofile As SolidEdgePart.Profile
        Dim objcircles As SolidEdgeFrameworkSupport.Circles2d
        Dim objrelations As SolidEdgeFrameworkSupport.Relations2d
        Dim objmodel As SolidEdgePart.Model
        Dim objcutouts As SolidEdgePart.ExtrudedCutouts
        Dim newcutout As SolidEdgePart.ExtrudedCutout
        Dim objplaneside As SolidEdgePart.FeaturePropertyConstants
        Dim offsetside As SolidEdgePart.OffsetSideConstants

        Try
            'get the refplane
            refplanes = partdoc.RefPlanes
            refplane = refplanes.Item(planeno)

            'add a profile for the sketch 
            objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            objprofile.Visible = False

            objcircles = objprofile.Circles2d
            objrelations = objprofile.Relations2d

            'add the circle
            objcircles.AddByCenterRadius(0, position / 1000, diameter / 2000)

            'select the side for the cutout 
            If direction = "left" Then
                objplaneside = SolidEdgePart.FeaturePropertyConstants.igLeft
                offsetside = SolidEdgePart.OffsetSideConstants.seOffsetLeft
            Else
                objplaneside = SolidEdgePart.FeaturePropertyConstants.igRight
                offsetside = SolidEdgePart.OffsetSideConstants.seOffsetRight
            End If
            'close the profile
            objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

            objmodel = partdoc.Models.Item(1)

            objcutouts = objmodel.ExtrudedCutouts

            'add the cutout
            'first add another refplane or find if it was already created
            nippletoplane = CreateParRefPlane(refplanes, refplane, direction)

            'create the cutout from to
            newcutout = objcutouts.AddFromTo(objprofile, SolidEdgePart.FeaturePropertyConstants.igLeft, refplane, nippletoplane)

            'set the offset 
            newcutout.SetFromFaceOffsetData(refplane, offsetside, -offset / 1000)

            newcutout.Name = newname

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function CreateParRefPlane(refplanes As SolidEdgePart.RefPlanes, nippleplane As SolidEdgePart.RefPlane, direction As String) As SolidEdgePart.RefPlane
        Dim newplane As SolidEdgePart.RefPlane = Nothing
        Dim offset As Double

        Try
            For Each refplane As SolidEdgePart.RefPlane In refplanes
                If refplane.Global Then
                    If refplane.Name = "ToPlane" Then
                        newplane = refplane
                    End If
                End If
            Next

            If direction = "right" Then
                offset = -0.1
            Else
                offset = 0.1
            End If

            If newplane Is Nothing Then
                newplane = refplanes.AddParallelByDistance(nippleplane, offset, SolidEdgePart.ReferenceElementConstants.igNormalSide, Local:=False)
                Debug.Print(nippleplane.Name)
                newplane.Name = "ToPlane"
            End If
            newplane.Visible = False
        Catch ex As Exception

        End Try

        Return newplane
    End Function

    Shared Sub CreateNippleTube(headerdiameter As Double, ByRef nippletube As TubeData, pressure As Integer, consys As ConSysData, connectionside As String, finneddepth As Double, nopasses As Integer)
        Dim angle As Double = 45
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objvariables As SolidEdgeFramework.Variables
        Dim controllength, totallength As Double

        Try
            If File.Exists(nippletube.TubeFile.Fullfilename) Then
                partdoc = General.seapp.Documents.Open(nippletube.TubeFile.Fullfilename)
                General.seapp.DoIdle()
                nippletube.FileName = partdoc.Name
                objvariables = partdoc.Variables

                If consys.ControlNipples Then
                    'control nipple length
                    If General.currentunit.ModelRangeName = "GACV" Then
                        If nippletube.HeaderType = "outlet" Then
                            controllength = Calculation.ControlNippleLengthGACV(nippletube, consys.ConType, consys.OutletHeaders.First.Xlist, connectionside, consys.OutletHeaders.First.Displacehor, finneddepth)
                        Else
                            controllength = Calculation.ControlNippleLengthGACV(nippletube, consys.ConType, consys.InletHeaders.First.Xlist, connectionside, consys.InletHeaders.First.Displacehor, finneddepth)
                        End If
                        If Math.Abs(nippletube.Length - controllength) > 5 Then
                            nippletube.Length = Math.Round(controllength)
                        End If
                    ElseIf General.currentjob.ModelRange = "GCO" Or General.currentjob.ModelRange = "NNNN" Then
                        'do not change nipple tube length, keep the raw input
                    ElseIf General.currentunit.UnitDescription = "VShape" Then
                        If pressure > 100 Then
                            If General.currentunit.ModelRangeName.Substring(3, 1) = "V" Then
                                totallength = 350
                            Else
                                If nippletube.HeaderType = "outlet" Then
                                    totallength = consys.OutletHeaders.First.Dim_a + headerdiameter + 120
                                Else
                                    totallength = consys.InletHeaders.First.Dim_a + headerdiameter + 120
                                End If
                            End If
                        Else
                            If consys.HeaderMaterial = "V" Then
                                If nopasses = 2 Then
                                    totallength = 450
                                Else
                                    totallength = 400
                                End If
                            ElseIf pressure <= 16 Then
                                totallength = 400
                            Else
                                totallength = 350
                            End If
                            'consider flange connections
                            If consys.FlangeID <> "" Then

                            End If
                        End If
                        If nippletube.HeaderType = "outlet" Then
                            controllength = Calculation.ControlNippleLengthVShape(nippletube.Length, totallength, consys.ConType, consys.OutletHeaders.First)
                        Else
                            controllength = Calculation.ControlNippleLengthVShape(nippletube.Length, totallength, consys.ConType, consys.InletHeaders.First)
                        End If
                        If Math.Abs(nippletube.Length - controllength) > 5 Then
                            nippletube.Length = Math.Round(controllength)
                        End If
                    ElseIf General.currentunit.ApplicationType = "Condenser" Then
                        If consys.ConType = 2 Or consys.ConType >= 6 Then
                            'loose flange
                            totallength = 350
                            'change of formula, removed the flange dimensions, nippletube was 10mm too short
                            If nippletube.HeaderType = "inlet" Then
                                controllength = totallength - (consys.InletHeaders.First.Dim_a + consys.InletHeaders.First.Tube.Diameter / 2 + 3)
                            Else
                                controllength = totallength - (consys.OutletHeaders.First.Dim_a + consys.OutletHeaders.First.Tube.Diameter / 2 + 3)
                            End If
                            If Math.Abs(nippletube.Length - controllength) > 5 Then
                                nippletube.Length = Math.Round(controllength)
                            End If
                        ElseIf pressure >= 100 Then
                            totallength = 120
                            If nippletube.Materialcodeletter = "D" Then
                                controllength = 90
                            Else
                                controllength = totallength - headerdiameter / 2
                            End If
                            If Math.Abs(nippletube.Length - controllength) > 5 Then
                                nippletube.Length = Math.Round(controllength)
                            End If
                        ElseIf consys.ConType = 0 Then
                            If nippletube.Materialcodeletter <> "C" Then
                                If nippletube.HeaderType = "inlet" Then
                                    If nippletube.Diameter <= 60.3 Then
                                        controllength = 120
                                    Else
                                        controllength = 150
                                    End If
                                Else
                                    If nippletube.Diameter <= 60.3 And finneddepth < 150 Then
                                        controllength = 200
                                    Else
                                        controllength = 130
                                    End If
                                End If
                                controllength += headerdiameter / 2
                            Else
                                If nippletube.Diameter <= 64 Then
                                    controllength = 120
                                Else
                                    If nippletube.HeaderType = "inlet" Then
                                        controllength = 150
                                    Else
                                        controllength = 130
                                    End If
                                End If
                            End If
                            If Math.Abs(nippletube.Length - controllength) > 5 Then
                                nippletube.Length = Math.Round(controllength)
                            End If
                        End If
                    End If
                End If

                'apply basic data (Ø, wt, length, material)
                objvariables.Add("HeaderDiameter", "28")
                Dim objvar As SolidEdgeFramework.variable = objvariables.Add("Angle", "45")
                objvar.Formula = "ATN( Diameter / HeaderDiameter )*180/3.14159"
                objvariables.Add("e", "1")

                GetSetVariableValue("Diameter", objvariables, "set", nippletube.Diameter)
                GetSetVariableValue("HeaderDiameter", objvariables, "set", headerdiameter)
                GetSetVariableValue("Length", objvariables, "set", nippletube.Length)
                GetSetVariableValue("WallThickness", objvariables, "set", nippletube.WallThickness)
                GetSetVariableValue("Angle", objvariables, "set", angle)

                'create tube body on x-z plane
                Dim objmodels As SolidEdgePart.Models = partdoc.Models
                Dim refplane As SolidEdgePart.RefPlane = partdoc.RefPlanes.Item(3)
                Dim objsketch As SolidEdgePart.Sketch = partdoc.Sketches.Add
                Dim formulas() As String = {"Diameter", "Diameter - 2*WallThickness", "HeaderDiameter /2- e", "2*( HeaderDiameter /2- e )*tan( Angle *3.14159/180)"}
                Dim newprofilelist As New List(Of SolidEdgePart.Profile)
                Dim objprofile As SolidEdgePart.Profile

                For i As Integer = 0 To 1
                    objprofile = objsketch.Profiles.Add(refplane)
                    objprofile.Visible = False
                    Dim objcircle As SolidEdgeFrameworkSupport.Circle2d = objprofile.Circles2d.AddByCenterRadius(0, 0, nippletube.Diameter / 1000 - 2 * i * nippletube.WallThickness / 1000)
                    Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objprofile.Dimensions
                    Dim circdim As SolidEdgeFrameworkSupport.Dimension = objdims.AddCircularDiameter(objcircle)
                    circdim.Constraint = True
                    circdim.Formula = formulas(i)

                    objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                    newprofilelist.Add(objprofile)
                Next

                objmodels.AddFiniteExtrudedProtrusion(2, newprofilelist.ToArray, 2, 0.1)
                Dim objexpros As SolidEdgePart.ExtrudedProtrusions = partdoc.Models.Item(1).ExtrudedProtrusions
                Dim objcyl As SolidEdgePart.ExtrudedProtrusion = objexpros(0)
                objcyl.Name = "Nippletube"
                Dim varcount As Integer
                Dim dims(0) As Object
                objcyl.GetDimensions(varcount, dims)
                Dim lengthdim As SolidEdgeFrameworkSupport.Dimension
                lengthdim = TryCast(dims(0), SolidEdgeFrameworkSupport.Dimension)
                If lengthdim IsNot Nothing Then
                    lengthdim.Formula = "Length"
                End If

                SetMaterial(General.seapp, partdoc, nippletube.Materialcodeletter, "part")

                'create sketch for header cutout
                If nippletube.Materialcodeletter = "C" Then
                    CreateHCut(partdoc, headerdiameter, nippletube, angle, formulas)
                Else
                    If headerdiameter = nippletube.Diameter Then
                        CreateSimilarNipple(partdoc, nippletube)
                    ElseIf nippletube.Materialcodeletter <> "D" Then
                        CreateProjections(partdoc, CreateHeaderCutSketch(partdoc, headerdiameter, nippletube.WallThickness), partdoc.RefPlanes.Item(1), 6)
                        CreateBlueSurf(partdoc, False)
                        SubtractBody(partdoc, SolidEdgePart.SESubtractDirection.igSubtractDirectionRight)
                    End If
                End If

                'create tube end
                If GNData.CheckCaps(nippletube.Materialcodeletter, nippletube.Diameter, nippletube.WallThickness, False, pressure, "nipple") And Not consys.HasFTCon Then
                    nippletube.TopCapNeeded = True
                End If

                If Not nippletube.TopCapNeeded And Not consys.HasFTCon Then
                    CreateTubeEnds(partdoc, "", "", nippletube, 3)
                End If

                Dim hasSV As Boolean = False
                If nippletube.SVPosition(1) = "axial" Then
                    hasSV = True
                    'create sv cutout in cap feature
                    SetCutoutFeature(partdoc, 3, 8.5, 0, direction:="right", cuttype:="SV", newname:="svhole")
                ElseIf nippletube.SVPosition(1) = "perp" Then
                    'nipple perp only for evaporators
                    Dim position As Double
                    If pressure <= 16 Then
                        position = -(nippletube.Length - 17)
                    ElseIf pressure > 50 And General.currentunit.ModelRangeSuffix = "CP" Then
                        position = -(nippletube.Length - 9)
                    Else
                        position = -(nippletube.Length - 70)
                    End If
                    SetCutoutFeature(partdoc, 1, 8.5, position, direction:="right", cuttype:="SV", newname:="svhole")
                End If

                Dim closingoffset As Double

                If Not consys.HasFTCon Then
                    closingoffset = GNData.TubeClosingOffset(nippletube.Diameter, hasSV, General.currentjob.Plant, nippletube.WallThickness)
                End If

                nippletube.RawLength = nippletube.Length + closingoffset

                If nippletube.Materialcodeletter <> "D" Then
                    nippletube.RawLength -= headerdiameter / 2 - GNData.GetTubeOffset(nippletube.Diameter, nippletube.Materialcodeletter, General.currentjob.Plant)
                End If

                nippletube.RawMaterial = Database.GetTubeERP(nippletube.Diameter, "Headertube", pressure, nippletube.Materialcodeletter)

                GetSetCustomProp(partdoc, "Raw_Material", nippletube.RawMaterial, "write")
                GetSetCustomProp(partdoc, "Raw_Length", nippletube.RawLength / 1000, "write")

                General.seapp.Documents.CloseDocument(partdoc.FullName, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateHCut(partdoc As SolidEdgePart.PartDocument, headerdiameter As Double, nippletube As TubeData, angle As Double, formulas() As String)
        Try
            Dim cutdepth As Double = headerdiameter / 2 - GNData.GetCutoutOffset(headerdiameter, nippletube.Diameter)
            Debug.Print("e=" + cutdepth.ToString)
            GetSetVariableValue("e", partdoc.Variables, "set", cutdepth)
            Dim objprofile As SolidEdgePart.Profile = partdoc.ProfileSets.Add().Profiles.Add(partdoc.RefPlanes.Item(2))

            Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objprofile.Dimensions

            Dim ypos As Double = Math.Round(2 * (headerdiameter / 2 - cutdepth) * Math.Tan(angle * Math.PI / 180))
            Dim objrels As SolidEdgeFrameworkSupport.Relations2d = objprofile.Relations2d
            Dim objline1 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(-(headerdiameter / 2 - cutdepth) / 1000, -ypos / 2000, -(headerdiameter / 2 - cutdepth) / 1000, ypos / 2000)
            objrels.AddVertical(objline1)
            Dim linedim As SolidEdgeFrameworkSupport.Dimension = objdims.AddLength(objline1)
            linedim.Constraint = True
            linedim.Formula = formulas(3)

            'vertical at tube end
            Dim objline2 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, headerdiameter / 2000, 0, -headerdiameter / 2000)
            objrels.AddFix(objline2)

            Dim distancedim As SolidEdgeFrameworkSupport.Dimension = objdims.AddDistanceBetweenObjects(objline2, 0, 0, 0, False, objline1, 0, 0, 0, False)
            distancedim.MeasurementAxisDirection = True
            distancedim.Constraint = True
            distancedim.Formula = formulas(2)

            'line at the bottom
            Dim objline3 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, -headerdiameter / 2000, -0.1, -headerdiameter / 2000)
            objrels.AddKeypoint(objline2, 1, objline3, 0)
            objrels.AddHorizontal(objline3)

            'line at the top
            Dim objline4 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, headerdiameter / 2000, -0.1, headerdiameter / 2000)
            objrels.AddKeypoint(objline2, 0, objline4, 0)
            objrels.AddHorizontal(objline4)

            'connector top
            Dim objline5 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(-0.1, headerdiameter / 2000, -0.09, headerdiameter / 2000 + 0.01)
            objrels.AddKeypoint(objline4, 1, objline5, 0)
            objrels.AddKeypoint(objline5, 1, objline1, 1)

            'connector bottom
            Dim objline6 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(-0.1, -headerdiameter / 2000, -0.09, -headerdiameter / 2000 + 0.01)
            objrels.AddKeypoint(objline3, 1, objline6, 0)
            objrels.AddKeypoint(objline6, 1, objline1, 0)
            Dim angledim As SolidEdgeFrameworkSupport.Dimension = objdims.AddAngle(objline6)
            angledim.Constraint = True
            angledim.Formula = "Angle"

            objrels.AddEqual(objline6, objline5)
            objrels.AddEqual(objline4, objline3)
            objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
            objprofile.Visible = False

            Dim objcutouts As SolidEdgePart.ExtrudedCutouts = partdoc.Models.Item(1).ExtrudedCutouts
            Dim objcutout As SolidEdgePart.ExtrudedCutout = objcutouts.AddThroughAll(objprofile, ProfileSide:=SolidEdgePart.FeaturePropertyConstants.igRight, ProfilePlaneSide:=SolidEdgePart.FeaturePropertyConstants.igLeft)
            objcutout.ExtentType = 16
            objcutout.ExtentSide = 6
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub CreateSimilarHeader(partdoc As SolidEdgePart.PartDocument, header As HeaderData, circuit As CircuitData, mp As Integer, profileside As Integer, headeralignment As String, deltaangle As Double)
        Dim objsketches As SolidEdgePart.Sketchs = partdoc.Sketches
        Dim refplane As SolidEdgePart.RefPlane
        Dim objprofile As SolidEdgePart.Profile
        Dim wt As Double
        Dim planeno As Integer
        Dim yfillet As Integer

        Try
            'create the refplane first - 90° to nipple plane
            If deltaangle = 0 Or deltaangle = 180 Then
                planeno = CreateAngRefPlane(partdoc, 3, -90, "NipplePlane")
            ElseIf General.currentunit.ApplicationType = "Evaporator" And deltaangle = 90 Then
                planeno = CreateAngRefPlane(partdoc, 3, Math.Abs(deltaangle) + deltaangle, "NipplePlane")
            Else
                planeno = CreateAngRefPlane(partdoc, 3, Math.Abs(deltaangle) - deltaangle, "NipplePlane")
            End If

            refplane = partdoc.RefPlanes.Item(planeno)
            objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objprofile.Dimensions
            Dim objrels As SolidEdgeFrameworkSupport.Relations2d = objprofile.Relations2d

            With header
                wt = Math.Round(.Tube.WallThickness / 1000, 4)
                For Each pos In header.Nipplepositions
                    Dim line1 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(- .Tube.Diameter * mp / 2000, (pos - .Tube.Diameter / 2) / 1000, - .Tube.Diameter * mp / 2000, (pos + .Tube.Diameter / 2) / 1000)
                    Dim l1dim As SolidEdgeFrameworkSupport.Dimension = objdims.AddLength(line1)
                    l1dim.Constraint = True
                    objrels.AddVertical(line1)

                    Dim line2 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(-mp * wt, pos / 1000 - wt, - .Tube.Diameter * mp / 2000, (pos - .Tube.Diameter / 2) / 1000)
                    objrels.AddKeypoint(line1, 0, line2, 1)

                    Dim line3 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(-mp * wt, pos / 1000 + wt, - .Tube.Diameter * mp / 2000, (pos + .Tube.Diameter / 2) / 1000)
                    objrels.AddKeypoint(line3, 1, line1, 1)

                    Dim angledim As SolidEdgeFrameworkSupport.Dimension = objdims.AddAngle(line2)
                    angledim.Constraint = True

                    yfillet = pos / Math.Abs(pos)

                    Dim objarc As SolidEdgeFrameworkSupport.Arc2d = objprofile.Arcs2d.AddByStartAlongEnd(-mp * wt, pos / 1000 - wt, 0, pos / 1000, -mp * wt, pos / 1000 + wt)
                    Dim arcdim As SolidEdgeFrameworkSupport.Dimension = objdims.AddRadius(objarc)
                    arcdim.Constraint = True
                    arcdim.Value = .Tube.WallThickness / 1000
                    objrels.AddKeypoint(line2, 0, objarc, 1)
                    objrels.AddKeypoint(line3, 0, objarc, 2)
                    objrels.AddTangent(line2, objarc)
                    objrels.AddTangent(line3, objarc)

                    objrels.AddEqual(line2, line3)
                    objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                Next

                Dim objcutouts As SolidEdgePart.ExtrudedCutouts = partdoc.Models.Item(1).ExtrudedCutouts
                'outlet default = 1, mirrored = 2 // 
                Dim objcutout As SolidEdgePart.ExtrudedCutout = objcutouts.AddThroughAll(objprofile, ProfileSide:=profileside, ProfilePlaneSide:=SolidEdgePart.FeaturePropertyConstants.igRight)
                objcutout.ExtentType = 16
                objcutout.ExtentSide = 6
            End With

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreateSimilarNipple(partdoc As SolidEdgePart.PartDocument, nippletube As TubeData)
        Dim objsketches As SolidEdgePart.Sketchs = partdoc.Sketches
        Dim refplane As SolidEdgePart.RefPlane = partdoc.RefPlanes.Item(2)
        Dim objprofile As SolidEdgePart.Profile

        Try
            'objprofile = objsketches.Add().Profiles.Add(refplane)
            objprofile = partdoc.ProfileSets.Add().Profiles.Add(refplane)
            objprofile.Visible = False
            Dim objdims As SolidEdgeFrameworkSupport.Dimensions = objprofile.Dimensions
            Dim objrels As SolidEdgeFrameworkSupport.Relations2d = objprofile.Relations2d

            With nippletube
                Dim line1 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, .Diameter / 2000, 0, - .Diameter / 2000)
                Dim l1dim As SolidEdgeFrameworkSupport.Dimension = objdims.AddLength(line1)
                l1dim.Constraint = True
                objrels.AddVertical(line1)

                'top line
                Dim line2 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, .Diameter / 2000, - .Diameter / 2000, .Diameter / 2000)
                objrels.AddKeypoint(line1, 0, line2, 0)
                Dim l2dim As SolidEdgeFrameworkSupport.Dimension = objdims.AddLength(line2)
                l2dim.Constraint = True
                objrels.AddHorizontal(line2)

                'bottom line
                Dim line3 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(0, - .Diameter / 2000, - .Diameter / 2000, - .Diameter / 2000)
                objrels.AddKeypoint(line1, 1, line3, 0)
                Dim l3dim As SolidEdgeFrameworkSupport.Dimension = objdims.AddLength(line3)
                l3dim.Constraint = True
                objrels.AddHorizontal(line3)


                Dim line4 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(- .Diameter / 2000, .Diameter / 2000, 0, 0)
                objrels.AddKeypoint(line2, 1, line4, 0)

                Dim line5 As SolidEdgeFrameworkSupport.Line2d = objprofile.Lines2d.AddBy2Points(- .Diameter / 2000, - .Diameter / 2000, 0, 0)
                objrels.AddKeypoint(line3, 1, line5, 0)

                Dim angledim As SolidEdgeFrameworkSupport.Dimension = objdims.AddAngle(line4)
                angledim.Constraint = True

                'dim objfillet as SolidEdgeGeometry.f
                Dim objfillet As SolidEdgeFrameworkSupport.Arc2d = objprofile.Arcs2d.AddAsFillet(line4, line5, .WallThickness / 1000, -1, -1)
                Dim arcdim As SolidEdgeFrameworkSupport.Dimension = objdims.AddRadius(objfillet)
                arcdim.Constraint = True

                objrels.AddEqual(line4, line5)
                objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

                Dim objcutouts As SolidEdgePart.ExtrudedCutouts = partdoc.Models.Item(1).ExtrudedCutouts
                Dim objcutout As SolidEdgePart.ExtrudedCutout = objcutouts.AddThroughAll(objprofile, ProfileSide:=SolidEdgePart.FeaturePropertyConstants.igRight, ProfilePlaneSide:=SolidEdgePart.FeaturePropertyConstants.igRight)
                objcutout.ExtentType = 16
                objcutout.ExtentSide = 6
            End With

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function GetFaces(occurance As SolidEdgeAssembly.Occurrence, partname As String, ftype As String, figure As Integer, alignment As String) As SolidEdgeGeometry.Face()
        Dim partdoc As SolidEdgePart.PartDocument
        Dim partmodel As SolidEdgePart.Model
        Dim plface, axface, facereturn() As SolidEdgeGeometry.Face
        Dim allfaces As New List(Of SolidEdgeGeometry.Face)
        Dim expros As SolidEdgePart.ExtrudedProtrusions
        Dim expro As SolidEdgePart.ExtrudedProtrusion
        Dim featureno As Integer

        'returns a face each for axial relation and for planar relation
        Try
            partdoc = occurance.PartDocument
            partmodel = partdoc.Models.Item(1)
            Select Case partname
                Case "Stutzen"
                    plface = GetFaceFromLoop(partmodel, "Stutzen", "planar", figure, alignment)
                    axface = GetFaceFromLoop(partmodel, "Stutzen", "axial", figure, alignment)
                Case "SV"
                    expros = partmodel.ExtrudedProtrusions
                    expro = expros.Item(7)
                    plface = expro.BottomCap
                    axface = GetSideFace(partdoc, "protrusion", 7, 1)
                Case "nadapter"
                    axface = GetFaceFromLoop(partmodel, "nadapter", "axial", figure, alignment)
                    If ftype = "fixed" Then
                        plface = GetFace(partdoc, "")
                    Else
                        plface = GetFace(partdoc, "stutzen")
                    End If
                Case "CUplate"
                    'get cutoutnumber
                    featureno = GetCutoutNo(partdoc)
                    axface = GetSideFace(partdoc, "cutout", featureno, 1)
                    plface = GetFace(partdoc, "")
                Case Else
                    axface = GetSideFace(partdoc, "protrusion", 1, 1)
                    If ftype = "fixed" Then
                        plface = GetFace(partdoc, "")
                    Else
                        plface = GetFace(partdoc, "stutzen")
                    End If
            End Select

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        facereturn = {plface, axface}

        Return facereturn
    End Function

    Shared Function GetFaceFromLoop(partmodel As SolidEdgePart.Model, parttype As String, facetype As String, figure As String, alignment As String) As SolidEdgeGeometry.Face
        Dim fgroup As SolidEdgePart.FeatureGroup
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim objface, plface, axface, returnface As SolidEdgeGeometry.Face
        Dim allfaces As New List(Of SolidEdgeGeometry.Face)
        Dim xminlist, xmaxlist, yminlist, ymaxlist, zminlist, zmaxlist As New List(Of Double)
        Dim formlist As New List(Of Integer)
        Dim minrange(2), maxrange(2) As Double
        Dim geoform As Integer
        Dim x, y, z, xcheck As Double

        objbody = partmodel.Body
        objloops = objbody.Loops

        For j As Integer = 0 To objloops.Count - 1
            objloop = objloops(j)
            objloop.GetRange(minrange, maxrange)
            x = Math.Round(minrange(0) * 1000, 3)
            y = Math.Round(minrange(1) * 1000, 3)
            z = Math.Round(minrange(2) * 1000, 3)
            xminlist.Add(x)
            yminlist.Add(y)
            zminlist.Add(z)

            x = Math.Round(maxrange(0) * 1000, 3)
            y = Math.Round(maxrange(1) * 1000, 3)
            z = Math.Round(maxrange(2) * 1000, 3)
            xmaxlist.Add(x)
            ymaxlist.Add(y)
            zmaxlist.Add(z)

            objface = objloop.Face
            geoform = objface.GeometryForm
            formlist.Add(geoform)
            allfaces.Add(objface)
        Next

        If parttype = "Stutzen" Then
            If figure = 3 Then
                fgroup = partmodel.FeatureGroups.Item(1)
                If fgroup.Suppress = True Then
                    'search for min x value
                    xcheck = xminlist.Min
                Else
                    'search for max x value
                    xcheck = xmaxlist.Max
                End If
            Else
                If alignment = "normal" Then
                    xcheck = xmaxlist.Max
                Else
                    xcheck = xminlist.Min
                End If
            End If

            For j As Integer = 0 To formlist.Count - 1
                If facetype = "planar" Then
                    If formlist(j) = 9 Then
                        If zminlist(j) = zminlist.Min And (xminlist(j) = xcheck Or xmaxlist(j) = xcheck) Then
                            'found planar face
                            Debug.Print("Planar face is allfaces(" + j.ToString + ")")
                            plface = allfaces(j)
                        End If
                    End If
                Else
                    If formlist(j) = 10 Then
                        If figure = 3 Or alignment = "reverse" Then
                            If yminlist(j) = yminlist.Min And zminlist(j) = zminlist.Min Then
                                'found axial face
                                Debug.Print("Axial face is allfaces(" + j.ToString + ")")
                                axface = allfaces(j)
                            End If
                        Else
                            If ymaxlist(j) = 0 And yminlist(j) = 0 Then
                                'found axial face
                                Debug.Print("Axial face is allfaces(" + j.ToString + ")")
                                axface = allfaces(j)
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf parttype = "pipe" Then
            axface = allfaces.Last
        Else
            'nadapter
            For j As Integer = 0 To formlist.Count - 1
                If formlist(j) = 10 And facetype = "axial" Then
                    Debug.Print("Axial face is allfaces(" + j.ToString + ")")
                    axface = allfaces(j)
                    j = formlist.Count - 1
                End If
            Next
        End If

        If facetype = "axial" Then
#Disable Warning BC42104 ' Die axface-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
            returnface = axface
#Enable Warning BC42104 ' Die axface-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Else
#Disable Warning BC42104 ' Die plface-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
            returnface = plface
#Enable Warning BC42104 ' Die plface-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        End If

        Return returnface
    End Function

    Shared Function GetCutoutNo(partdoc As SolidEdgePart.PartDocument) As Integer
        Dim objmodel As SolidEdgePart.Model = partdoc.Models.Item(1)
        Dim objcutouts As SolidEdgePart.ExtrudedCutouts = objmodel.ExtrudedCutouts
        Dim objcutout As SolidEdgePart.ExtrudedCutout
        Dim featureno As Integer

        For i As Integer = 0 To objcutouts.Count - 1
            objcutout = objcutouts(i)
            If objcutout.Suppress = False Then
                featureno = i + 1
            End If
        Next

        Return featureno
    End Function

    Shared Function GetCapAxFace(asmdoc As SolidEdgeAssembly.AssemblyDocument, diameter As Double) As SolidEdgeGeometry.Face
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim objocc As SolidEdgeAssembly.Occurrence
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim objface, axface As SolidEdgeGeometry.Face
        Dim facetype, faceno As Integer
        Dim minrange(2), maxrange(2), xmax, ymax, zmax As Double

        axface = Nothing

        Try
            objoccs = asmdoc.Occurrences
            objocc = objoccs.Item(objoccs.Count)
            partdoc = objocc.PartDocument

            objmodel = partdoc.Models.Item(1)
            objbody = objmodel.Body
            objloops = objbody.Loops

            faceno = -1

            'find a suiting face
            For i As Integer = 0 To objloops.Count - 1
                If axface Is Nothing Then
                    objloop = objloops(i)
                    objface = objloop.Face
                    facetype = objface.GeometryForm
                    If facetype = 10 Then    'cylindric face
                        objface.GetRange(minrange, maxrange)
                        xmax = Math.Round(maxrange(0) * 1000, 2)
                        ymax = Math.Round(maxrange(1) * 1000, 2)
                        zmax = Math.Round(maxrange(2) * 1000, 2)
                        If xmax >= diameter / 2 Then
                            faceno = i
                        ElseIf ymax >= diameter / 2 Then
                            faceno = i
                        End If
                        If faceno > -1 Then
                            axface = objface
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return axface
    End Function

    Shared Function GetHeaderFace(headerocc As SolidEdgeAssembly.Occurrence, facetype As String) As SolidEdgeGeometry.Face
        Dim partdoc As SolidEdgePart.PartDocument
        Dim objmodel As SolidEdgePart.Model
        Dim objexpros As SolidEdgePart.ExtrudedProtrusions
        Dim objexpro As SolidEdgePart.ExtrudedProtrusion
        Dim objfaces As SolidEdgeGeometry.Faces
        Dim headerface As SolidEdgeGeometry.Face

        partdoc = headerocc.PartDocument
        objmodel = partdoc.Models.Item(1)
        objexpros = objmodel.ExtrudedProtrusions
        objexpro = objexpros.Item(1)


        Select Case facetype
            Case "axial"
                objfaces = objexpro.SideFaces
            Case "bottom"
                objfaces = objexpro.BottomCaps
            Case Else
                objfaces = objexpro.TopCaps
        End Select

        headerface = objfaces.Item(1)

        Return headerface
    End Function

    Shared Function GetCapPlanFace(partdoc As SolidEdgePart.PartDocument, diameter As Double, pressure As Integer, materialcodeletter As String) As Object
        Dim objmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim objface As SolidEdgeGeometry.Face
        Dim facelist As New List(Of SolidEdgeGeometry.Face)
        Dim objrefplanes As SolidEdgePart.RefPlanes
        Dim refplane As SolidEdgePart.RefPlane
        Dim facetype, faceno, index As Integer
        Dim facearea, minrange(2), maxrange(2), xmin, xmax, ymin, ymax, zmin, zmax As Double
        Dim arealist As New List(Of Double)
        Dim captype As String
        Dim objplanface As Object

        If (diameter = 133 And pressure = 16) Or pressure > 16 Or materialcodeletter <> "C" Then
            captype = "Deckel"
        Else
            captype = "Kappe"
        End If

        objplanface = Nothing

        Try

            objmodel = partdoc.Models.Item(1)
            objbody = objmodel.Body
            objloops = objbody.Loops

            faceno = -1

            'find a suiting face
            For i As Integer = 0 To objloops.Count - 1
                If objplanface Is Nothing Then
                    objloop = objloops(i)
                    objface = objloop.Face
                    facetype = objface.GeometryForm
                    If captype = "Deckel" Then
                        If facetype = 9 Then    'cylindric face
                            facearea = objface.Area
                            arealist.Add(facearea)
                            facelist.Add(objface)
                        End If
                    Else
                        If facetype = 10 And faceno < 0 Then
                            'check the maxpoints to see which refplane can be used
                            objface.GetRange(minrange, maxrange)
                            xmin = Math.Round(minrange(0) * 1000, 2)
                            xmax = Math.Round(maxrange(0) * 1000, 2)
                            ymin = Math.Round(minrange(1) * 1000, 2)
                            ymax = Math.Round(maxrange(1) * 1000, 2)
                            zmin = Math.Round(minrange(2) * 1000, 2)
                            zmax = Math.Round(maxrange(2) * 1000, 2)
                            If maxrange.Max * 1000 > diameter / 2 Then
                                If xmax = ymax Then
                                    faceno = 1
                                ElseIf xmax = zmax Then
                                    faceno = 3
                                Else 'ymax = zmax
                                    faceno = 2
                                End If
                                objrefplanes = partdoc.RefPlanes
                                refplane = objrefplanes.Item(faceno)
                                objplanface = refplane
                            End If
                        End If
                    End If
                End If
            Next

            If captype = "Deckel" Then
                index = arealist.IndexOf(arealist.Min)
                objplanface = facelist(index)
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objplanface
    End Function

    Shared Sub CreateProjections(partdoc As SolidEdgePart.PartDocument, objsketch As SolidEdgePart.Sketch, refplane As SolidEdgePart.RefPlane, direction As Integer)

        Try
            'get profile
            Dim objprofile As SolidEdgePart.Profile = objsketch.Profile
            Dim profile_curve_body As SolidEdgeGeometry.CurveBody
            profile_curve_body = objprofile.CurveBody
            Dim objCurves As Object = profile_curve_body.Curves
            Dim minarr1(2), maxarr1(2), minarr2(2), maxarr2(2) As Double

            profile_curve_body.Curves.Item(1).GetRange(minarr1, maxarr1)

            profile_curve_body.Curves.Item(3).GetRange(minarr2, maxarr2)

            'get face
            Dim objheader As SolidEdgePart.ExtrudedProtrusion = partdoc.Models.Item(1).ExtrudedProtrusions.Item(1)
            Dim objsidefaces As SolidEdgeGeometry.Faces = objheader.SideFaces
            Dim objface1 As SolidEdgeGeometry.Face = objsidefaces.Item(1)
            Dim objface2 As SolidEdgeGeometry.Face = objsidefaces.Item(2)
            Dim outerface, innerface As SolidEdgeGeometry.Face

            objface1.GetExactRange(minarr1, maxarr1)

            objface2.GetExactRange(minarr2, maxarr2)

            If Math.Abs(maxarr1(1)) > Math.Abs(maxarr2(1)) Then
                innerface = objface2
                outerface = objface1
            Else
                innerface = objface1
                outerface = objface2
            End If

            Dim objconstruction As SolidEdgePart.Constructions = partdoc.Constructions
            Dim objprojcurveout, objprojcurvein As SolidEdgePart.ProjectCurve

            Dim objplane As SolidEdgePart.RefPlane = refplane

            Dim obj1() As Object = {objCurves(1), objCurves(2)}
            Dim obj2() As Object = {objCurves(3), objCurves(4)}

            Dim objsec1() As Object = {outerface, outerface}
            Dim objsec2() As Object = {innerface, innerface}

            objprojcurveout = objconstruction.ProjectCurves.AddArray(obj1, objsec1, objplane, direction, SolidEdgePart.FeaturePropertyConstants.igProjectOptionProject)
            objprojcurveout.Name = "FirstLine"
            objprojcurvein = objconstruction.ProjectCurves.AddArray(obj2, objsec2, objplane, direction, SolidEdgePart.FeaturePropertyConstants.igProjectOptionProject)
            objprojcurvein.Name = "SecondLine"
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function CreateNippleCutSketch(partdoc As SolidEdgePart.PartDocument, nipplediameter As Double, wallthickness As Double, position As Double, counter As Integer) As SolidEdgePart.Sketch
        Dim nippleplane As SolidEdgePart.RefPlane
        Dim objsketch As SolidEdgePart.Sketch

        objsketch = partdoc.Sketches.Add()

        Try
            nippleplane = partdoc.RefPlanes.Item("NippleRefplane")

            If nippleplane IsNot Nothing Then
                Dim objprofile As SolidEdgePart.Profile = objsketch.Profiles.Add(nippleplane)
                objsketch.Name = "NippleGeometry" + counter.ToString

                objprofile.Arcs2d.AddByStartAlongEnd(0, (position + nipplediameter / 2) / 1000, nipplediameter / 2000, position / 1000, 0, (position - nipplediameter / 2) / 1000)
                objprofile.Arcs2d.AddByStartAlongEnd(0, (position + nipplediameter / 2) / 1000, -nipplediameter / 2000, position / 1000, 0, (position - nipplediameter / 2) / 1000)

                objprofile.Arcs2d.AddByStartAlongEnd(0, (position + nipplediameter / 2 - wallthickness) / 1000, (nipplediameter / 2 - wallthickness) / 1000, position / 1000, 0, (position - (nipplediameter / 2 - wallthickness)) / 1000)
                objprofile.Arcs2d.AddByStartAlongEnd(0, (position + nipplediameter / 2 - wallthickness) / 1000, -(nipplediameter / 2 - wallthickness) / 1000, position / 1000, 0, (position - (nipplediameter / 2 - wallthickness)) / 1000)

                objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
                objprofile.Visible = False
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objsketch
    End Function

    Shared Function CreateHeaderCutSketch(partdoc As SolidEdgePart.PartDocument, headerdiameter As Double, wallthickness As Double) As SolidEdgePart.Sketch
        Dim nippleplane As SolidEdgePart.RefPlane
        Dim objsketch As SolidEdgePart.Sketch

        objsketch = partdoc.Sketches.Add

        Try
            nippleplane = partdoc.RefPlanes.Item(1)

            Dim objprofile As SolidEdgePart.Profile = objsketch.Profiles.Add(nippleplane)
            objsketch.Name = "HeaderGeometry"

            objprofile.Arcs2d.AddByStartAlongEnd(0, headerdiameter / 2000, headerdiameter / 2000, 0, 0, -headerdiameter / 2000)
            objprofile.Arcs2d.AddByStartAlongEnd(0, headerdiameter / 2000, -headerdiameter / 2000, 0, 0, -headerdiameter / 2000)

            objprofile.Arcs2d.AddByStartAlongEnd(0, (headerdiameter - wallthickness) / 2000, (headerdiameter - wallthickness) / 2000, 0, 0, -(headerdiameter - wallthickness) / 2000)
            objprofile.Arcs2d.AddByStartAlongEnd(0, (headerdiameter - wallthickness) / 2000, -(headerdiameter - wallthickness) / 2000, 0, 0, -(headerdiameter - wallthickness) / 2000)

            objprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)
            objprofile.Visible = False
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objsketch
    End Function

    Shared Sub CreateBlueSurf(partdoc As SolidEdgePart.PartDocument, closedEnds As Boolean)
        Dim objCSections(0 To 1) As Object
        Dim objOrigins(0 To 1) As Object
        Dim objGuideCurves(0 To 1) As Object
        Dim objbluesurf As SolidEdgePart.BlueSurf
        Dim objConstructions As SolidEdgePart.Constructions = partdoc.Constructions
        Dim edgesCol As SolidEdgeGeometry.Edges

        Try
            edgesCol = objConstructions.ProjectCurves.Item(1).Edges(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
            objCSections(0) = edgesCol

            edgesCol = objConstructions.ProjectCurves.Item(2).Edges(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
            objCSections(1) = edgesCol

            objOrigins(0) = objCSections(0).Item(1).StartVertex
            objOrigins(1) = objCSections(1).Item(1).StartVertex

            objGuideCurves(0) = Nothing
            objGuideCurves(1) = Nothing

            objbluesurf = objConstructions.BlueSurfs.Add(2, objCSections, objOrigins, 113, 0.0#, 113, 0.0#, 0, objGuideCurves, 113, 0.0#, 113, 0.0#, closedEnds, False)
            objbluesurf.Visible = False

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SubtractBody(partdoc As SolidEdgePart.PartDocument, direction As SolidEdgePart.SESubtractDirection)

        Try
            Dim objmodel As SolidEdgePart.Model = partdoc.Models.Item(1)

            Dim objtarget() As Object = {objmodel.Body}
            Dim objtool() As Object = {partdoc.Constructions.Item(partdoc.Constructions.Count).Body}
            Dim objdir() As SolidEdgePart.SESubtractDirection = {direction}

            Dim objsub As SolidEdgePart.Subtract = objmodel.Subtracts.Add(1, objtarget, 1, objtool, objdir, 0, 0)
        Catch ex As Exception

        End Try
    End Sub

    Shared Function GetRotationAxis(partdoc As SolidEdgePart.PartDocument, featuretyp As String) As SolidEdgePart.RefAxis
        Dim objmodel As SolidEdgePart.Model
        Dim objrevpros As SolidEdgePart.RevolvedProtrusions
        Dim objrevpro As SolidEdgePart.RevolvedProtrusion
        Dim objrevcuts As SolidEdgePart.RevolvedCutouts
        Dim objrevcut As SolidEdgePart.RevolvedCutout
        Dim objaxis As SolidEdgePart.RefAxis

        Try
            objmodel = partdoc.Models.Item(1)

            If featuretyp = "cutout" Then
                objrevcuts = objmodel.RevolvedCutouts
                objrevcut = objrevcuts.Item(1)
                objaxis = objrevcut.Axis
            Else
                objrevpros = objmodel.RevolvedProtrusions
                objrevpro = objrevpros.Item(1)
                objaxis = objrevpro.Axis
            End If

        Catch ex As Exception
            objaxis = Nothing
        End Try

        Return objaxis
    End Function

    Shared Sub WriteCostumProps(partdoc As SolidEdgePart.PartDocument, f As FileData)
        GetSetCustomProp(partdoc, "CSG", "1", "write")
        GetSetCustomProp(partdoc, "Auftragsnummer", f.Orderno, "write")
        GetSetCustomProp(partdoc, "Position", f.Orderpos, "write")
        GetSetCustomProp(partdoc, "Order_Projekt", f.Projectno, "write")
        GetSetCustomProp(partdoc, "AGP_Nummer", f.AGPno + ".", "write")
        GetSetCustomProp(partdoc, "CDB_Benennung_de", f.CDB_de, "write")
        GetSetCustomProp(partdoc, "CDB_Benennung_en", f.CDB_en, "write")
        GetSetCustomProp(partdoc, "CDB_Zusatzbenennung", f.CDB_Zusatzbenennung, "write")
        GetSetCustomProp(partdoc, "CDB_z_bemerkung", f.CDB_z_Bemerkung, "write")
        GetSetCustomProp(partdoc, "CDB_Material", f.CDB_Material, "write")
    End Sub

    Shared Sub CreateCutout(nippledata As NippleCutoutData, psmdoc As SolidEdgePart.SheetMetalDocument, conside As String)
        Dim psmmodel As SolidEdgePart.Model = psmdoc.Models.Item(1)
        Dim newprofile As SolidEdgePart.Profile
        Dim refplane As SolidEdgePart.RefPlane
        Dim objsketch As SolidEdgePart.Sketch
        Dim objlines As New List(Of SolidEdgeFrameworkSupport.Line2d)
        Dim points(3) As Double
        Dim ncutout As SolidEdgePart.NormalCutout
        Dim profileside As SolidEdgePart.FeaturePropertyConstants = SolidEdgePart.FeaturePropertyConstants.igRight
        Dim profilearr() As SolidEdgePart.Profile

        Try
            refplane = psmdoc.RefPlanes.Item(2)
            objsketch = psmdoc.Sketches.Add
            newprofile = objsketch.Profiles.Add(refplane)

            'add 4 lines around the center point, consider displacement
            points(0) = nippledata.YPos - nippledata.CutsizeY / 2
            points(1) = points(0) + nippledata.CutsizeY
            points(2) = nippledata.ZPos + nippledata.CutsizeZ / 2 + nippledata.Displacement
            points(3) = points(2) - nippledata.CutsizeZ

            If nippledata.HasFlange Then
                points(0) = 0
            End If

            For i As Integer = 0 To 3
                points(i) = -points(i) / 1000
            Next

            If conside = "left" Then
                points(0) = -points(0)
                points(1) = -points(1)
                profileside = SolidEdgePart.FeaturePropertyConstants.igLeft
            End If

            objlines.Add(newprofile.Lines2d.AddBy2Points(points(0), points(2), points(1), points(2)))
            objlines.Add(newprofile.Lines2d.AddBy2Points(points(1), points(2), points(1), points(3)))
            objlines.Add(newprofile.Lines2d.AddBy2Points(points(1), points(3), points(0), points(3)))
            objlines.Add(newprofile.Lines2d.AddBy2Points(points(0), points(3), points(0), points(2)))

            newprofile.End(SolidEdgePart.ProfileValidationType.igProfileClosed)

            profilearr = {newprofile}
            ncutout = psmmodel.NormalCutouts.AddThroughAllMulti(1, profilearr, SolidEdgePart.FeaturePropertyConstants.igLeft, SolidEdgePart.FeaturePropertyConstants.igSMClearanceCutout)
            ncutout.ExtentSide = SolidEdgePart.FeaturePropertyConstants.igRight
            ncutout.ProfileSide = profileside

            ReOrderFeature(ncutout, psmdoc)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub ReOrderFeature(ncutfeature As SolidEdgePart.NormalCutout, psmdoc As SolidEdgePart.SheetMetalDocument)
        Dim sketch As SolidEdgePart.Sketch
        Dim sketchs As SolidEdgePart.Sketchs
        Dim cflange As SolidEdgePart.ContourFlange

        Try
            cflange = psmdoc.Models.Item(1).ContourFlanges.Item(1)

            sketchs = psmdoc.Sketches
            sketch = sketchs.Item(sketchs.Count)
            sketch.Reorder(cflange, False)

            ncutfeature.Reorder(sketch, False)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub CoversheetProps(filetype As String)
        Dim psmdoc As SolidEdgePart.SheetMetalDocument
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        Dim objpsets As SolidEdgeFramework.PropertySets
        Dim objgsets As SolidEdgeFramework.Properties
        Dim objprop As SolidEdgeFramework.Property
        Dim category, csmat, matkey As String
        Dim propfound As Boolean = False

        Try
            'get cover sheet material
            csmat = PCFData.GetValue("ConnectionCoveringTypeA", "MaterialCodeLetter")

            If filetype = "part" Then
                psmdoc = General.seapp.ActiveDocument
                SetMaterial(General.seapp, psmdoc, "S" + csmat, "psm")
                'update flatten model
                Try
                    psmdoc.FlatPatternModels.Item(1).Update()
                Catch ex As Exception
                End Try
                objpsets = psmdoc.Properties
                category = "3D-Einzelteil"
            Else
                dftdoc = General.seapp.ActiveDocument
                objpsets = dftdoc.Properties
                category = "2D-Einzelteilzeichnung"
            End If
            objgsets = objpsets.Item(4)

            For i As Integer = 0 To objgsets.Count - 1
                objprop = objgsets(i)
                If objprop.Name = "CDB_Material" Then
                    propfound = True
                    Debug.Print("CDB Material found: " + objprop.Value)
                    Exit For
                End If
            Next

            If Not propfound Then
                objprop = objgsets.Add("CDB_Material", "")
            End If

            For i As Integer = 0 To objgsets.Count - 1
                objprop = objgsets(i)
                If objprop.Name = "CDB_ERP_Artnr." Then
                    'delete ERPCode
                    objprop.Value = ""
                    objprop.Delete()
                End If
                If objprop.Name = "CDB_Material" Then
                    matkey = ""
                    Select Case csmat
                        Case "G"
                            matkey = "AlMg (SP12)"
                        Case "S"
                            matkey = "St galv (SP15-1)"
                        Case "V"
                            matkey = "St stainl V (SP16-4)"
                        Case "W"
                            matkey = "St stainl V (SP16-2)"
                    End Select
                    objprop.Value = matkey
                End If
            Next

            objprop = objgsets.Add("Z_Kategorie", category)

            If filetype = "dft" Then
                If General.currentjob.Plant = "Tata" Then
                    objgsets.Add("AGP_Nummer", "1.")
                ElseIf General.currentjob.Plant = "Sibiu" Then
                    objgsets.Add("AGP_Nummer", "1.5")
                ElseIf General.currentjob.Plant = "Beji" Then
                    objgsets.Add("AGP_Nummer", "1.7")
                End If

                objgsets.Add("AGP_Plant", General.currentjob.Plant)
            End If
            objgsets.Add("Auftragsnummer", General.currentjob.OrderNumber)
            objgsets.Add("Position", General.currentjob.OrderPosition)
            objgsets.Add("Order_Projekt", General.currentjob.ProjectNumber)
            objgsets.Add("CSG", "1")

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function RenameCoversheet(ByRef oldfilename As String) As String
        Dim newfilename As String = ""
        Dim tempname As String
        Try
            tempname = "CS" + oldfilename.Substring(oldfilename.LastIndexOf("\") + 1, 10) + ".psm"
            newfilename = General.currentjob.Workspace + "\" + tempname
            My.Computer.FileSystem.RenameFile(oldfilename, tempname)
        Catch ex As Exception
            ex.ToString()
        End Try
        Return newfilename
    End Function

End Class
