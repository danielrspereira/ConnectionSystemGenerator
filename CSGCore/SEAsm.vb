Public Class SEAsm

    Shared Sub CreateSubAssembly(coil As CoilData, thickness As Double, ctoverhang As Double, asmname As String, circuitno As Integer)
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim ctno, finno As Integer
        Dim ctnolist As New List(Of Integer)
        Dim ctdoc, findoc As SolidEdgePart.PartDocument
        Dim offset() As Double = {0, 0, -coil.FinnedLength - thickness}
        Dim ctfeature As SolidEdgePart.RevolvedProtrusion
        Dim finfeature As SolidEdgePart.ExtrudedCutout
        Dim skipfront As Boolean = False
        Dim skipback As Boolean = False
        Dim ctname, finname As String

        Try
            asmdoc = General.seapp.Documents.Add(ProgID:="SolidEdge.AssemblyDocument")
            General.seapp.DoIdle()
            asmdoc.SaveAs(asmname)
            General.seapp.DoIdle()

            If asmname.Contains("\Consys") Then
                'skip unneccessary core tube and fin
                skipback = True
                Dim no As String = asmdoc.Name.Substring(6, 1)
                If coil.Circuits(CInt(no) - 1).CircuitType.Contains("Defrost") Then
                    'skip front
                    skipfront = True
                    skipback = False
                ElseIf coil.Circuits(CInt(no) - 1).NoPasses = 1 Or coil.Circuits(CInt(no) - 1).NoPasses = 3 Then
                    skipback = False
                End If
                SEPart.GetSetCustomProp(asmdoc, "GUE_Block", "0", "write")
                ctname = "ConsysCoretube" + circuitno.ToString
                finname = "ConsysFin" + coil.Number.ToString + circuitno.ToString
            Else
                'set a costum prop for BOM process
                WriteCustomProps(asmdoc, coil.CoilFile)
                SEPart.GetSetCustomProp(asmdoc, "GUE_Block", "1", "write")
                SEPart.GetSetVariableValue("RotationDirection", asmdoc.Variables, "add", coil.RotationDirection, 1)
                SEPart.GetSetVariableValue("FinnedLength", asmdoc.Variables, "add", coil.FinnedLength, 1)
                SEPart.GetSetVariableValue("FinnedHeight", asmdoc.Variables, "add", coil.FinnedHeight, 1)
                SEPart.GetSetVariableValue("FinnedDepth", asmdoc.Variables, "add", coil.FinnedDepth, 1)
                ctname = "CoilCoretube" + coil.Number.ToString
                finname = "CoilFin" + coil.Number.ToString + "1"
            End If

            'first fin
            finno = AddOccurance(asmdoc, finname, General.currentjob.Workspace, "par", coil.Occlist)
            If asmname.Contains("\Consys") Then
                coil.Occlist.RemoveAt(coil.Occlist.Count - 1)
            End If
            RemoveFromBOM(finno, asmdoc.Occurrences)
            asmdoc.Relations3d.AddGround(asmdoc.Occurrences.Item(finno))
            findoc = asmdoc.Occurrences.Item(finno).PartDocument

            'second fin
            AddOccurance(asmdoc, finname, General.currentjob.Workspace, "par", coil.Occlist)
            If asmname.Contains("\Consys") Then
                coil.Occlist.RemoveAt(coil.Occlist.Count - 1)
            End If
            RemoveFromBOM(finno + 1, asmdoc.Occurrences)

            For i As Integer = 1 To 3
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(finno), asmdoc.Occurrences.Item(finno + 1), i, i, False, offset(i - 1), True)
            Next

            If Not skipfront Then
                'first coretube
                ctno = AddOccurance(asmdoc, ctname, General.currentjob.Workspace, "par", coil.Occlist)
                If asmname.Contains("\Consys") Then
                    coil.Occlist.RemoveAt(coil.Occlist.Count - 1)
                End If
                ctnolist.Add(ctno)
                ctdoc = asmdoc.Occurrences.Item(ctno).PartDocument
                asmdoc.Occurrences.Item(ctno).IncludeInBom = False

                'assemble coretube with fin
                ctfeature = SEPart.GetFeature(ctdoc, "Coretube", "revprotrusion")
                finfeature = SEPart.GetFeature(findoc, "CT1", "cutout")

                If ctfeature IsNot Nothing And finfeature IsNot Nothing Then
                    Dim ctface As SolidEdgeGeometry.Face = SEPart.GetSideFace(ctdoc, "revprotrusion", 1, 1)
                    Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 1, 1)

                    SetAxialFaces(asmdoc, asmdoc.Occurrences.Item(finno), asmdoc.Occurrences.Item(ctno), finface, ctface, False, False)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(finno), asmdoc.Occurrences.Item(ctno), 1, 1, False, 0, False)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(finno), asmdoc.Occurrences.Item(ctno), 3, 3, False, ctoverhang, True)
                End If
            End If

            If Not skipback Then
                'second coretube
                ctno = AddOccurance(asmdoc, ctname, General.currentjob.Workspace, "par", coil.Occlist)
                If asmname.Contains("\Consys") Then
                    coil.Occlist.RemoveAt(coil.Occlist.Count - 1)
                End If
                ctnolist.Add(ctno)
                ctdoc = asmdoc.Occurrences.Item(ctno).PartDocument
                asmdoc.Occurrences.Item(ctno).IncludeInBom = False

                findoc = asmdoc.Occurrences.Item(finno + 1).PartDocument

                'assemble coretube with fin
                ctfeature = SEPart.GetFeature(ctdoc, "Coretube", "revprotrusion")
                finfeature = SEPart.GetFeature(findoc, "CT1", "cutout")

                If ctfeature IsNot Nothing And finfeature IsNot Nothing Then
                    Dim ctface As SolidEdgeGeometry.Face = SEPart.GetSideFace(ctdoc, "revprotrusion", 1, 1)
                    Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 1, 1)

                    SetAxialFaces(asmdoc, asmdoc.Occurrences.Item(finno + 1), asmdoc.Occurrences.Item(ctno), finface, ctface, True, False)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(finno + 1), asmdoc.Occurrences.Item(ctno), 1, 1, True, 0, False)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(finno + 1), asmdoc.Occurrences.Item(ctno), 3, 3, True, -ctoverhang - thickness, True)
                    CloneComponents(asmdoc, ctnolist, finno, finface, finno, "Coretubes")
                End If
            Else
                Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 1, 1)
                CloneComponents(asmdoc, ctnolist, finno, finface, finno, "Coretubes")
            End If

            If skipback Then
                asmdoc.Occurrences.Item(2).Delete()
            ElseIf skipfront Then
                asmdoc.Occurrences.Item(1).Delete()
                Dim objrels As SolidEdgeAssembly.Relations3d = asmdoc.Relations3d
                objrels.AddGround(asmdoc.Occurrences.Item(1))
            End If

            'create level views
            Dim blankconfig As String = CreateBlankView(asmdoc, ctname)

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AdjustCoilAssemblyDefrost(ByRef coil As CoilData, circuit As CircuitData, thickness As Double)
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim stno As Integer
        Dim stdoc, findoc As SolidEdgePart.PartDocument
        Dim stfeature As SolidEdgePart.RevolvedProtrusion
        Dim finfeature As SolidEdgePart.ExtrudedCutout
        Dim stnolist As New List(Of Integer)

        Try
            asmdoc = General.seapp.Documents.Open(coil.CoilFile.Fullfilename)
            General.seapp.DoIdle()

            'add first supporttube
            stno = AddOccurance(asmdoc, circuit.CoreTube.FileName, General.currentjob.Workspace, "par", coil.Occlist)
            stnolist.Add(stno)
            asmdoc.Occurrences.Item(stno).IncludeInBom = False
            findoc = asmdoc.Occurrences.Item(1).PartDocument
            stdoc = asmdoc.Occurrences.Item(stno).PartDocument

            'assemble first supporttube with fin
            stfeature = SEPart.GetFeature(stdoc, "Coretube", "revprotrusion")
            finfeature = SEPart.GetFeature(findoc, "CT1", "cutout")

            If stfeature IsNot Nothing And finfeature IsNot Nothing Then
                Dim stface As SolidEdgeGeometry.Face = SEPart.GetSideFace(stdoc, "revprotrusion", 1, 1)
                Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 2, 1)

                SetAxialFaces(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), finface, stface, False, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), 1, 1, False, 0, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), 3, 3, False, circuit.CoreTubeOverhang, True)
            End If

            'add second supporttube
            stno = AddOccurance(asmdoc, circuit.CoreTube.FileName, General.currentjob.Workspace, "par", coil.Occlist)
            stnolist.Add(stno)
            stdoc = asmdoc.Occurrences.Item(stno).PartDocument
            asmdoc.Occurrences.Item(stno).IncludeInBom = False

            findoc = asmdoc.Occurrences.Item(2).PartDocument

            'assemble second supporttube with fin
            stfeature = SEPart.GetFeature(stdoc, "Coretube", "revprotrusion")
            finfeature = SEPart.GetFeature(findoc, "CT1", "cutout")

            If stfeature IsNot Nothing And finfeature IsNot Nothing Then
                Dim ctface As SolidEdgeGeometry.Face = SEPart.GetSideFace(stdoc, "revprotrusion", 1, 1)
                Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 2, 1)

                SetAxialFaces(asmdoc, asmdoc.Occurrences.Item(2), asmdoc.Occurrences.Item(stno), finface, ctface, True, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(2), asmdoc.Occurrences.Item(stno), 1, 1, True, 0, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(2), asmdoc.Occurrences.Item(stno), 3, 3, True, -circuit.CoreTubeOverhang - thickness, True)
                CloneComponents(asmdoc, stnolist, 1, finface, 1, "Supporttubes")
            End If

            CreateBlankView(asmdoc, "CoilSupporttube1")

            'change coil.ordercomment
            coil.CoilFile.CDB_z_Bemerkung += ";Brine Circuiting:*" + circuit.PDMID
            SEPart.GetSetCustomProp(asmdoc, "CDB_z_bemerkung", coil.CoilFile.CDB_z_Bemerkung, "write")

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AdjustConsysAssemblyDefrost(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef consys As ConSysData, circuit As CircuitData, thickness As Double)
        Dim stno As Integer
        Dim stdoc, findoc As SolidEdgePart.PartDocument
        Dim stfeature As SolidEdgePart.RevolvedProtrusion
        Dim finfeature As SolidEdgePart.ExtrudedCutout
        Dim stnolist As New List(Of Integer)

        Try

            ''add second supporttube
            stno = AddOccurance(asmdoc, circuit.CoreTube.FileName, General.currentjob.Workspace, "par", consys.Occlist)
            stnolist.Add(stno)
            stdoc = asmdoc.Occurrences.Item(stno).PartDocument
            asmdoc.Occurrences.Item(stno).IncludeInBom = False

            findoc = asmdoc.Occurrences.Item(1).PartDocument

            'assemble second supporttube with fin
            stfeature = SEPart.GetFeature(stdoc, "Coretube", "revprotrusion")
            finfeature = SEPart.GetFeature(findoc, "CT1", "cutout")

            If stfeature IsNot Nothing And finfeature IsNot Nothing Then
                Dim ctface As SolidEdgeGeometry.Face = SEPart.GetSideFace(stdoc, "revprotrusion", 1, 1)
                Dim finface As SolidEdgeGeometry.Face = SEPart.GetSideFace(findoc, "cutout", 2, 1)

                SetAxialFaces(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), finface, ctface, True, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), 1, 1, True, 0, False)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(stno), 3, 3, True, -circuit.CoreTubeOverhang - thickness, True)
                CloneComponents(asmdoc, stnolist, 1, finface, 1, "Supporttubes")
            End If

            CreateBlankView(asmdoc, "ConsysSupporttube1")

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function AddOccurance(asmdoc As SolidEdgeAssembly.AssemblyDocument, partname As String, workspace As String, fileending As String, ByRef Occlist As List(Of PartData), Optional cref As String = "") As Integer
        Dim i As Integer = 0
        Dim occname As String
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim newocc As SolidEdgeAssembly.Occurrence
        Dim addocc As SolidEdgeAssembly.Occurrence
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objfixed As SolidEdgeAssembly.GroundRelation3d

        Try
            If partname.Contains(Library.TemplateParts.BOW1) Then
                partname += ".par"
            End If
            occname = General.GetFullFilename(workspace, partname, fileending)
            If occname <> "" Then
                'add part by name
                objoccs = asmdoc.Occurrences
                newocc = objoccs.AddByFilename(occname)
                i = objoccs.Count
                Occlist.Add(New PartData With {.Occindex = i, .Occname = newocc.Name, .Configref = cref})
                addocc = objoccs.Item(i)
                objrelations = addocc.Relations3d

                'delete fixed relation
                objfixed = objrelations.Item(1)
                objfixed.Delete()
            End If

        Catch ex As Exception

        End Try

        Return i
    End Function

    Shared Sub SetPlanar(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstocc As SolidEdgeAssembly.Occurrence, secondocc As SolidEdgeAssembly.Occurrence, firstplane As Integer, secondplane As Integer,
                         normals As Boolean, offset As Double, fixed As Boolean)
        Dim firstpart, secondpart As SolidEdgePart.PartDocument
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim fixedreference, adjreference As Object
        Dim basep() As Double = {0, 0, 0}

        Try
            'get the part docs
            firstpart = firstocc.PartDocument
            secondpart = secondocc.PartDocument

            objrelations = asmdoc.Relations3d

            'get the reference objects
            fixedreference = asmdoc.CreateReference(firstocc, firstpart.RefPlanes.Item(firstplane))
            adjreference = asmdoc.CreateReference(secondocc, secondpart.RefPlanes.Item(secondplane))

            objplanar = objrelations.AddPlanar(fixedreference, adjreference, normals, basep, basep)
            If fixed Then
                objplanar.Offset = offset / 1000
            Else
                'creates a float relation
                objplanar.FixedOffset = False
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetPlanarAsm(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstocc As SolidEdgeAssembly.Occurrence, secondocc As SolidEdgeAssembly.Occurrence, firstplane As Integer, secondplane As Integer,
                         normals As Boolean, offset As Double, fixed As Boolean)
        Dim firstasm, secondasm As SolidEdgeAssembly.AssemblyDocument
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim fixedreference, adjreference As Object
        Dim basep() As Double = {0, 0, 0}

        Try
            firstasm = firstocc.PartDocument
            secondasm = secondocc.PartDocument

            objrelations = asmdoc.Relations3d

            fixedreference = asmdoc.CreateReference(firstocc, firstasm.AsmRefPlanes.Item(firstplane))
            adjreference = asmdoc.CreateReference(secondocc, secondasm.AsmRefPlanes.Item(secondplane))

            objplanar = objrelations.AddPlanar(fixedreference, adjreference, normals, basep, basep)
            If fixed Then
                objplanar.Offset = offset / 1000
            Else
                'creates a float relation
                objplanar.FixedOffset = False
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub SetPlanarAsmByOcc(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstocc As SolidEdgeAssembly.Occurrence, secondocc As SolidEdgeAssembly.Occurrence, firstplane As Integer, secondplane As Integer,
                         normals As Boolean, offset As Double, fixed As Boolean)
        Dim firstpart, secondpart As SolidEdgePart.PartDocument
        Dim subfirstpart, subsecondpart As SolidEdgeAssembly.SubOccurrence
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim fixedreference, adjreference, subfixedref, subadjref As Object
        Dim basep() As Double = {0, 0, 0}

        Try

            subfirstpart = firstocc.SubOccurrences.Item(1)
            subsecondpart = secondocc.SubOccurrences.Item(1)

            'get the part docs
            firstpart = subfirstpart.ThisAsOccurrence.PartDocument
            secondpart = subsecondpart.ThisAsOccurrence.PartDocument

            objrelations = asmdoc.Relations3d

            'get the reference objects
            subfixedref = asmdoc.CreateReference(subfirstpart.ThisAsOccurrence, firstpart.RefPlanes.Item(firstplane))
            subadjref = asmdoc.CreateReference(subsecondpart.ThisAsOccurrence, secondpart.RefPlanes.Item(secondplane))

            fixedreference = asmdoc.CreateReference(firstocc, subfixedref)
            adjreference = asmdoc.CreateReference(secondocc, subadjref)

            objplanar = objrelations.AddPlanar(fixedreference, adjreference, normals, basep, basep)
            If fixed Then
                objplanar.Offset = offset / 1000
            Else
                'creates a float relation
                objplanar.FixedOffset = False
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function SetAxialFaces(asmdoc As SolidEdgeAssembly.AssemblyDocument, fixedocc As SolidEdgeAssembly.Occurrence, adjocc As SolidEdgeAssembly.Occurrence, fixedface As SolidEdgeGeometry.Face,
                            adjface As SolidEdgeGeometry.Face, normals As Boolean, fixedrot As Boolean) As Integer
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim firstaxobj, secondaxobj As Object
        Dim relstatus As Integer = 0

        Try
            firstaxobj = asmdoc.CreateReference(fixedocc, fixedface)
            secondaxobj = asmdoc.CreateReference(adjocc, adjface)

            objrelations = asmdoc.Relations3d
            objaxial = objrelations.AddAxial(firstaxobj, secondaxobj, normals)
            objaxial.FixedRotate = fixedrot

            relstatus = objaxial.Status
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return relstatus
    End Function

    Shared Sub SetAngular(asmdoc As SolidEdgeAssembly.AssemblyDocument, occno As Integer, angle As Double, measures() As Boolean, refno As Integer, refplaneno As Integer, measureplane As Integer, placeplaneno As Integer)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objangular As SolidEdgeAssembly.AngularRelation3d
        Dim adjocc, refocc As SolidEdgeAssembly.Occurrence
        Dim adjpart, refpart As SolidEdgePart.PartDocument
        Dim refplanes, adjrefplanes As SolidEdgePart.RefPlanes
        Dim refplane, adjrefplane, measurerefplane As SolidEdgePart.RefPlane
        Dim refreference, adjreference, measureelement As Object

        Try
            refocc = objoccs.Item(refno)

            'get the references from the fin item
            refpart = refocc.PartDocument
            refplanes = refpart.RefPlanes
            refplane = refplanes.Item(refplaneno)
            refreference = asmdoc.CreateReference(refocc, refplane)
            measurerefplane = refplanes.Item(measureplane)
            measureelement = asmdoc.CreateReference(refocc, measurerefplane)

            'get the reference from the placed part
            adjocc = objoccs.Item(occno)
            adjpart = adjocc.PartDocument
            adjrefplanes = adjpart.RefPlanes
            adjrefplane = adjrefplanes.Item(placeplaneno)
            adjreference = asmdoc.CreateReference(adjocc, adjrefplane)

            objrelations = asmdoc.Relations3d

            objangular = objrelations.AddAngular(refreference, adjreference, False, False, measureelement, Nothing, angle, measures(0), measures(1), measures(2))
        Catch ex As Exception

        End Try

    End Sub

    Shared Sub SetAxial(asmdoc As SolidEdgeAssembly.AssemblyDocument, fixedno As Integer, adjno As Integer, facenumber As Integer, Optional fixrot As Boolean = False)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim fixedocc, adjocc As SolidEdgeAssembly.Occurrence
        Dim fixedpart, adjpart As SolidEdgePart.PartDocument
        Dim objrevpro As SolidEdgePart.RevolvedProtrusion
        Dim fixedmodel, adjmodel As SolidEdgePart.Model
        Dim objbody As SolidEdgeGeometry.Body
        Dim objloops As SolidEdgeGeometry.Loops
        Dim objloop As SolidEdgeGeometry.Loop
        Dim objfaces As SolidEdgeGeometry.Faces
        Dim fixedface, adjface As SolidEdgeGeometry.Face
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim fixedreference, adjreference As Object

        fixedocc = objoccs.Item(fixedno)
        adjocc = objoccs.Item(adjno)

        Try
            'find the cylindric face of the first revolved protrusion as reference
            fixedpart = fixedocc.PartDocument
            fixedmodel = fixedpart.Models.Item(1)
            objrevpro = fixedmodel.RevolvedProtrusions.Item(1)
            objfaces = objrevpro.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryCylinder)
            fixedface = objfaces.Item(2)

            fixedreference = asmdoc.CreateReference(fixedocc, fixedface)

            adjpart = adjocc.PartDocument
            adjmodel = adjpart.Models.Item(1)
            objbody = adjmodel.Body
            objloops = objbody.Loops
            If facenumber = 24 Then
                'standard type 1 bows
                facenumber = objloops.Count
            End If

            objloop = objloops.Item(facenumber)
            adjface = objloop.Face
            adjreference = asmdoc.CreateReference(adjocc, adjface)

            objrelations = asmdoc.Relations3d
            objaxial = objrelations.AddAxial(fixedreference, adjreference, True)
            objaxial.FixedRotate = fixrot
        Catch ex As Exception

        End Try

    End Sub

    Shared Sub CloneComponents(asmdoc As SolidEdgeAssembly.AssemblyDocument, masteroccnos As List(Of Integer), refoccno As Integer, refface As SolidEdgeGeometry.Face, cloneoccno As Integer, groupname As String)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim refocc, cloneocc As SolidEdgeAssembly.Occurrence
        Dim objComponenetsToClone As Array = Array.CreateInstance(GetType(Object), masteroccnos.Count)
        Dim objFaces As Array = Array.CreateInstance(GetType(Object), 1)
        Dim objCloneEnv As Array = Array.CreateInstance(GetType(Object), 1)
        Dim obj As Object
        Dim errorstate As Integer

        Try
            For i As Integer = 1 To masteroccnos.Count
                objComponenetsToClone.SetValue(objoccs.Item(masteroccnos(i - 1)), i - 1)
            Next

            refocc = objoccs.Item(refoccno)
            cloneocc = objoccs.Item(cloneoccno)

            obj = asmdoc.CreateReference(refocc, refface)
            objFaces.SetValue(obj, 0)

            objCloneEnv.SetValue(cloneocc, 0)

            asmdoc.CreateCloneComponents(objComponenetsToClone, objFaces, objCloneEnv, SolidEdgeAssembly.CloneComponentOptions.seCreateGroundRelationships, True,
                                         SolidEdgeAssembly.CloneMatchTypeOptions.CloneMatchTypeAutomatic, errorstate)
            If errorstate = 0 Then
                asmdoc.AssemblyGroups.Item(asmdoc.AssemblyGroups.Count).Name = groupname
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceBows(asmdoc As SolidEdgeAssembly.AssemblyDocument, bowprops() As List(Of Double), bowids As List(Of String), side As String, workspace As String, fin As String,
                           reftube As String, ctoverhang As Double, offset() As Double, ByRef circuit As CircuitData, ByRef coildata As CoilData)
        Dim x1list, y1list, x2list, y2list, l1list, typelist, levellist As List(Of Double)
        Dim dx1, dy1, dx2, dy2, level1, level2, origin(), angle As Double
        Dim hx1list, hy1list, hx2list, hy2list As New List(Of Double)
        Dim usedbows, similarbows, ctlist As New List(Of Integer)
        Dim tubeno, bowno, faceno, copyno, patterncount As Integer
        Dim pattername, refconfig As String
        Dim skipbow, measures(), showmessage As Boolean

        Try
            x1list = bowprops(0)
            y1list = bowprops(1)
            x2list = bowprops(2)
            y2list = bowprops(3)
            typelist = bowprops(5)
            levellist = bowprops(bowprops.Count - 2)
            l1list = bowprops.Last
            patterncount = 1
            showmessage = False
            If reftube.ToLower.Contains("support") Then
                refconfig = "s" + side
            Else
                refconfig = "c" + side
            End If
            For i As Integer = 0 To bowids.Count - 1
                skipbow = False
                'Check if bow was already placed by a pattern
                For Each entry As Integer In usedbows
                    If i = entry Then
                        skipbow = True
                    End If
                Next

                If skipbow = False Then
                    'Get dx, dy and level for comparison
                    dx1 = Math.Round(x1list(i) - x2list(i), 3)
                    dy1 = Math.Round(y1list(i) - y2list(i), 3)
                    level1 = l1list(i)
                    For j As Integer = i To bowids.Count - 1
                        dx2 = Math.Round(x1list(j) - x2list(j), 3)
                        dy2 = Math.Round(y1list(j) - y2list(j), 3)
                        level2 = l1list(j)
                        If dx1 = dx2 And dy1 = dy2 And level1 = level2 And j > i And typelist(i) = typelist(j) Then
                            'add similar bows to the current and total list
                            usedbows.Add(j)
                            similarbows.Add(j)
                        End If
                    Next

                    'add first bow
                    bowno = AddOccurance(asmdoc, bowids(i), workspace, "par", coildata.Occlist, refconfig)
                    If side = "back" Then
                        'check for hairpin
                        For Each hp In circuit.Hairpins
                            If bowids(i) = hp.RefBow Then
                                RemoveFromBOM(bowno, asmdoc.Occurrences)
                            End If
                        Next
                    End If

                    If bowno > 0 Then
                        'get the origin for placement → here consider 2nd circuit / offset
                        origin = Calculation.GetPartOrigin(x1list(i), y1list(i), side, ctoverhang, offset(0), offset(1))
                        Dim ishp As Boolean = False
                        For Each hp In circuit.Hairpins
                            If hp.PDMID = bowids(i) Then
                                origin(1) = -ctoverhang
                                ishp = True
                                Exit For
                            End If
                        Next

                        'find the core tube
                        tubeno = GetTubeNoNew(origin, circuit.CoreTubes.Last, coildata.Alignment)

                        If tubeno > 0 Then
                            'set planar relation
                            SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(bowno), 3, 3, False, -5, True)

                            'set axial relation
                            faceno = SEPart.GetFaceNo(asmdoc.Occurrences.Item(bowno), fin, "normal", False)

                            SetAxial(asmdoc, tubeno, bowno, faceno)

                            'get angle
                            If ishp Then
                                angle = Calculation.GetAngleRad(x1list(i), y1list(i), x2list(i), y2list(i), "fronthp")
                                If x1list(i) = x2list(i) Then
                                    'angle -= Math.Round(Math.PI, 4)
                                End If
                            Else
                                angle = Calculation.GetAngleRad(x1list(i), y1list(i), x2list(i), y2list(i), side)
                            End If '+90

                            'get measure bools
                            measures = Calculation.GetMeasures(side, "bow")

                            'set angular relation
                            SetAngular(asmdoc, bowno, angle, measures, 1, 1, 3, 1)

                            'now find all similar bows for the pattern
                            For Each bowdupeno As Integer In similarbows
                                origin = Calculation.GetPartOrigin(x1list(bowdupeno), y1list(bowdupeno), side, ctoverhang, offset(0), offset(1))
                                If ishp Then
                                    origin(1) = -ctoverhang
                                End If
                                copyno = GetTubeNoNew(origin, circuit.CoreTubes.Last, coildata.Alignment)
                                If copyno > 0 Then
                                    ctlist.Add(copyno)
                                End If
                            Next

                            'create the pattern
                            If similarbows.Count > 0 Then
                                If reftube.ToLower.Contains("support") Then
                                    pattername = side + "brinebows" + patterncount.ToString
                                Else
                                    pattername = side + "bows" + patterncount.ToString
                                End If
                                CreateDuplicates(asmdoc, bowno, tubeno, ctlist, pattername)
                                patterncount += 1

                                If side = "back" Then
                                    For Each hp In circuit.Hairpins
                                        If bowids(i) = hp.RefBow Then
                                            Dim asmpat As SolidEdgeAssembly.AssemblyPattern = asmdoc.AssemblyPatterns.Item(pattername)
                                            For n As Integer = 0 To asmpat.Count - 1
                                                Dim patocc As SolidEdgeAssembly.AssemblyPatternOccurrence = asmpat(n)
                                                Dim objarr(0) As Object
                                                patocc.GetOccurrences(objarr)
                                                Dim singleocc As SolidEdgeAssembly.Occurrence = objarr(0)
                                                RemoveFromBOM(singleocc.Index, asmdoc.Occurrences)
                                            Next
                                        End If
                                    Next
                                End If
                            End If
                        End If

                    End If

                End If
                ctlist.Clear()
                similarbows.Clear()
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetTubeNoNew(origin() As Double, coretubes As CoreTubeLocations, coilalignment As String, Optional attempt As Integer = 0)
        Dim xindex As Integer = coretubes.Xlist.IndexOf(Math.Round(origin(0), 4))
        Dim yindex As Integer = coretubes.Ylist.IndexOf(origin(1))
        Dim zindex As Integer = coretubes.Zlist.IndexOf(Math.Round(origin(2), 4))
        Dim ctindex As Integer = 0

        If xindex > -1 And yindex > -1 And zindex > -1 Then
            Dim ctorigin(2) As Double

            'got 3 lists with index numbers of occurances. only 1 occurance can be in all 3 lists 
            Dim result = coretubes.Xindizes(xindex).Intersect(coretubes.Yindizes(yindex).Intersect(coretubes.Zindizes(zindex))).ToList

            If result.Count = 1 Then
                ctindex = result(0)
            Else
                Debug.Print("Result count =/= 1")
            End If
        ElseIf attempt = 0 Then
            If coilalignment = "vertical" Then
                ctindex = GetTubeNoNew({origin(0) + 0.005, origin(1), origin(2)}, coretubes, coilalignment, 1)
            Else
                ctindex = GetTubeNoNew({origin(0), origin(1), origin(2) + 0.005}, coretubes, coilalignment, 1)
            End If
        ElseIf attempt = 1 Then
            If coilalignment = "vertical" Then
                ctindex = GetTubeNoNew({origin(0) - 0.005, origin(1), origin(2)}, coretubes, coilalignment, 2)
            Else
                ctindex = GetTubeNoNew({origin(0), origin(1), origin(2) - 0.005}, coretubes, coilalignment, 2)
            End If
        End If

        Return ctindex
    End Function

    Shared Function GetTubeNo(objoccs As SolidEdgeAssembly.Occurrences, origin() As Double, occshortname As String) As Integer
        Dim objocc As SolidEdgeAssembly.Occurrence
        Dim occorigin() As Double
        Dim i As Integer = 1
        Dim loopexit As Boolean = False
        Dim itemno As Integer = 0

        Do
            objocc = objoccs.Item(i)
            'only check the core tube dummies
            If objocc.Name.Contains(occshortname) Then
                Dim matchcount As Integer = 0

                occorigin = GetOccOrigin(objocc)

                'compare the origin of the occurance with the bow position
                For j As Integer = 0 To occorigin.Count - 1
                    If Math.Abs(occorigin(j) - origin(j)) < 0.1 Then
                        matchcount += 1
                    End If
                Next

                If matchcount = 3 Then
                    loopexit = True
                    itemno = i
                Else
                    loopexit = False
                End If
                If i = objoccs.Count - 1 Then
                    loopexit = True
                End If

            End If

            i += 1
        Loop Until loopexit Or i > objoccs.Count

        Return itemno
    End Function

    Shared Function GetOccOrigin(objocc As SolidEdgeAssembly.Occurrence) As Double()
        Dim origin(), x, y, z As Double
        Dim coordlist As New List(Of Double)

        objocc.GetOrigin(x, y, z)
        origin = {x, y, z}
        For Each value As Double In origin
            value *= 1000
            value = Math.Round(value, 4)
            coordlist.Add(value)
        Next

        origin = {coordlist(0), coordlist(1), coordlist(2)}

        Return origin
    End Function

    Shared Sub CreateDuplicates(asmdoc As SolidEdgeAssembly.AssemblyDocument, masteroccno As Integer, originoccno As Integer, copyoccnos As List(Of Integer), patternname As String)
        Dim asmoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim asmpatterns As SolidEdgeAssembly.AssemblyPatterns = asmdoc.AssemblyPatterns
        Dim objparts(), objorigin As SolidEdgeAssembly.Occurrence
        Dim dupeoccs As New List(Of SolidEdgeAssembly.Occurrence)

        'PartToPattern - the part(s) which will be duplicated
        objparts = {asmoccs.Item(masteroccno)}
        'FromOccurence - the reference part
        objorigin = asmoccs.Item(originoccno)

        'ToOccurence - the new reference parts for duplication
        For Each number As Integer In copyoccnos
            dupeoccs.Add(asmoccs.Item(number))
        Next

        Try
            asmpatterns.CreateDuplicate(patternname, 1, objparts, objorigin, copyoccnos.Count, dupeoccs.ToArray)
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub CreatePattern(asmdoc As SolidEdgeAssembly.AssemblyDocument, patternname As String, occnamelist As List(Of String), circuit As CircuitData, alignment As String, delta As Double)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim occlist As New List(Of SolidEdgeAssembly.Occurrence)
        Dim asmpattern As SolidEdgeAssembly.AssemblyPattern
        Dim objProfile As SolidEdgePart.Profile
        Dim objRefPattern As SolidEdgeFrameworkSupport.RectangularPattern2d

        Try

            'create sketch for pattern 
            objProfile = SEPart.CreatePatternSketch(asmdoc, circuit.Quantity, circuit.PitchX, circuit.PitchY, circuit.CircuitSize, alignment)

            objRefPattern = objProfile.RectangularPatterns2d.Item(1)

            'get the occurances for the pattern by their name
            For Each objocc As SolidEdgeAssembly.Occurrence In objoccs
                For Each occname In occnamelist
                    Dim addobj As Boolean = False
                    If objocc.Name.Contains(occname) Then
                        'check location
                        addobj = CheckLocation(objocc, delta, alignment, patternname)
                        If addobj Then
                            occlist.Add(objocc)
                        End If
                    End If
                Next
            Next

            asmpattern = asmdoc.AssemblyPatterns.Create(patternname, occlist.ToArray, objRefPattern, objProfile)

            If patternname.Contains("Back") Then
                For Each occname In occnamelist
                    For Each hp In circuit.Hairpins
                        If hp.PDMID <> "" And hp.PDMID <> "NULL" Then
                            If occname = hp.RefBow Then
                                'remove all occ of this bow in the new pattern
                                For n As Integer = 0 To asmpattern.Count - 1
                                    Dim patocc As SolidEdgeAssembly.AssemblyPatternOccurrence = asmpattern(n)
                                    Dim objarr(0) As Object
                                    patocc.GetOccurrences(objarr)
                                    For i As Integer = 0 To objarr.Count - 1
                                        Dim singleocc As SolidEdgeAssembly.Occurrence = objarr(1)
                                        If singleocc.Name.Contains(hp.RefBow) Then
                                            RemoveFromBOM(singleocc.Index, asmdoc.Occurrences)
                                        End If
                                    Next
                                Next
                            End If
                        End If
                    Next
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function CheckLocation(objocc As SolidEdgeAssembly.Occurrence, circoffset As Double, coilposition As String, patternname As String) As Boolean
        Dim x, y, z, relcoord As Double
        Dim addobject As Boolean = False
        objocc.GetOrigin(x, y, z)

        If coilposition = "horizontal" Then
            relcoord = Math.Round(x * 1000, 3)
            If relcoord > circoffset Then
                addobject = True
            End If
        Else
            relcoord = Math.Round(z * 1000, 3)
            If relcoord > circoffset Then
                addobject = True
            End If
        End If

        If patternname.Contains("Front") Then
            If y > 0 Then
                addobject = False
            End If
        Else
            If y < 0 Then
                addobject = False
            End If
        End If

        Return addobject
    End Function

    Shared Sub AddSubAssemblytoMaster(asmdoc As SolidEdgeAssembly.AssemblyDocument, subasmname As String, ByRef occlist As List(Of PartData))
        Dim no As Integer

        Try
            If subasmname <> "" Then
                no = AddOccurance(asmdoc, subasmname, General.currentjob.Workspace, "asm", occlist)
                If no > 0 Then
                    If Not subasmname.Contains("\Coil2.asm") Then
                        asmdoc.Relations3d.AddGround(asmdoc.Occurrences.Item(no))
                    End If
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub RepositionAssembly(asmdoc As SolidEdgeAssembly.AssemblyDocument, coil As CoilData)
        Dim offsets() As Double
        Dim firstcoil, secondcoil As SolidEdgeAssembly.Occurrence

        Try
            If coil.Alignment = "vertical" Then
                offsets = {0, coil.Gap, 0}
            Else
                offsets = {coil.Gap, 0, 0}
            End If
            firstcoil = asmdoc.Occurrences.Item("Coil1.asm:1")
            secondcoil = asmdoc.Occurrences.Item("Coil2.asm:1")

            For i As Integer = 1 To 3
                SetPlanarAsmByOcc(asmdoc, firstcoil, secondcoil, i, i, False, offsets(i - 1), True)
            Next

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceHeaders(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef consys As ConSysData)
        Dim offsets() As Double
        Dim normals() As Boolean
        Dim headerplanes(), finplanes() As Integer

        Try
            finplanes = {2, 3, 1}
            If consys.HeaderAlignment = "horizontal" Then
                headerplanes = {1, 3, 2}
                normals = {False, False, True}
            Else
                headerplanes = {2, 3, 1}
                normals = {False, False, False}
            End If

            'place headers
            If consys.InletHeaders.First.Tube.Quantity > 0 Then
                For i As Integer = 0 To consys.InletHeaders.Count - 1
                    Dim headerno As Integer = AddOccurance(asmdoc, consys.InletHeaders(i).Tube.FileName, General.currentjob.Workspace, "par", consys.Occlist, "inlet")
                    offsets = consys.InletHeaders(i).Origin
                    For j As Integer = 0 To 2
                        If j = 1 And offsets(j) < 0 And Not consys.InletHeaders(i).Tube.IsBrine And consys.InletHeaders(i).OddLocation = "back" Then
                            'offsets(j) = Math.Abs(offsets(j))
                        End If
                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(headerno), finplanes(j), headerplanes(j), normals(j), offsets(j), True)
                    Next
                Next
            End If

            'place headers
            For i As Integer = 0 To consys.OutletHeaders.Count - 1
                Dim headerno As Integer = AddOccurance(asmdoc, consys.OutletHeaders(i).Tube.FileName, General.currentjob.Workspace, "par", consys.Occlist, "outlet")
                offsets = consys.OutletHeaders(i).Origin
                For j As Integer = 0 To 2
                    If j = 1 And offsets(j) < 0 And Not consys.OutletHeaders(i).Tube.IsBrine And consys.OutletHeaders(i).OddLocation = "back" Then
                        'offsets(j) = Math.Abs(offsets(j))
                    End If
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(headerno), finplanes(j), headerplanes(j), normals(j), offsets(j), True)
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub PlaceStutzen(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim headercoord, offset As Double
        Dim coordlist As List(Of Double)
        Dim usednos As New List(Of Integer)
        Dim patterncount As Integer = 1
        Dim stutzenno, tubeno As Integer
        Dim side As String = "front"

        Try
            If consys.HeaderAlignment = "horizontal" Then
                headercoord = header.Origin(2)
                coordlist = header.Ylist
            Else
                headercoord = header.Origin(0)
                coordlist = header.Xlist
            End If

            If ((circuit.NoPasses = 1 Or circuit.NoPasses = 3) And header.OddLocation = "back") Or circuit.CircuitType.Contains("Defrost") Then
                side = "back"
            End If

            If header.StutzenDatalist.Count > 0 Then
                For i As Integer = 0 To header.StutzenDatalist.Count - 1
                    Dim skipstutzen As Boolean = False
                    If usednos.IndexOf(i) > -1 Then
                        skipstutzen = True
                    End If

                    If Not skipstutzen Then
                        Dim similarstutzen, ctlist As New List(Of Integer)
                        For j As Integer = i To header.StutzenDatalist.Count - 1
                            If header.StutzenDatalist(i).ID = header.StutzenDatalist(j).ID And header.StutzenDatalist(i).ABV = header.StutzenDatalist(j).ABV And j <> i Then
                                similarstutzen.Add(j)
                            End If
                        Next

                        stutzenno = AddOccurance(asmdoc, header.StutzenDatalist(i).ID, General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)

                        If stutzenno > 0 Then
                            Dim ctorigin() As Double = Calculation.GetPartOrigin(header.Xlist(i), header.Ylist(i), side, circuit.CoreTubeOverhang, 0, 0)
                            tubeno = GetTubeNoNew(ctorigin, consys.CoreTubes.First, consys.HeaderAlignment)

                            If tubeno > 0 Then
                                Dim stutzenalign As String = "normal"

                                With header.StutzenDatalist(i)
                                    If .Figure = 5 And General.currentunit.ModelRangeSuffix <> "CP" Then
                                        stutzenalign = SEPart.GetStutzenAlignment(asmdoc, stutzenno, .ABV, header, circuit.CoreTubeOverhang)
                                        If stutzenalign = "reverse" Then
                                            'set planar but with input of faces
                                            SetReversePlanar(asmdoc, tubeno, stutzenno, -5, True)
                                        End If
                                    End If
                                    offset = -5
                                    If .SpecialTag = "s1star*" Then
                                        offset = 15
                                    End If

                                    If stutzenalign = "normal" Then
                                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(stutzenno), 3, 3, False, offset, True)
                                    End If
                                    'get the face number for axial relation
                                    Dim faceno As Integer = GetFaceNo(asmdoc, stutzenno, circuit.FinType, stutzenalign, False)

                                    If faceno = 0 Then
                                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(stutzenno), 2, 2, False, 0, True)
                                    Else
                                        SetAxial(asmdoc, tubeno, stutzenno, faceno)
                                    End If
                                    Dim fixedplaneno As Integer = SEPart.GetPlaneNo(consys.HeaderAlignment)

                                    If General.currentunit.UnitDescription = "VShape" AndAlso .SpecialTag <> "" Then
                                        If .Figure = 4 Then
                                            Dim measures() As Boolean = Calculation.GetMeasures(side, "stutzen")
                                            Dim rotangle As Double
                                            Dim deltah As Double = Math.Abs(.ABV)
                                            Dim deltav As Double

                                            If deltah = 0 Then
                                                deltav = 25
                                                If header.Tube.HeaderType = "outlet" Then
                                                    deltav *= -1
                                                End If
                                            Else
                                                'deltav depending of the specialtag
                                                Select Case .SpecialTag
                                                    Case "sInT4r1"
                                                        deltav = 35
                                                    Case "sInT4r2"
                                                        deltav = 60.7
                                                    Case "sInT4r3"
                                                        deltav = 38.7
                                                    Case "sInT4r4"
                                                        deltav = 29.5
                                                    Case "sInT4r5"
                                                        deltav = 25
                                                    Case "sInT4r6"
                                                        deltav = 50
                                                    Case "sInT4r7"
                                                        deltav = 63.1
                                                    Case "sOutT4r1"
                                                        deltav = -50
                                                    Case "sOutT4r2"
                                                        deltav = -25
                                                End Select
                                                If (circuit.ConnectionSide = "right" And header.Tube.HeaderType = "inlet") Or (circuit.ConnectionSide = "left" And header.Tube.HeaderType = "outlet") Then
                                                    deltah *= -1
                                                End If
                                            End If
                                            If side = "back" Then
                                                measures(2) = True
                                            End If

                                            rotangle = Calculation.GetAngleRad(ctorigin(0), ctorigin(2), ctorigin(0) + deltah, ctorigin(2) + deltav, side)
                                            SetAngular(asmdoc, stutzenno, rotangle, measures, 1, 1, 3, 1)
                                        Else
                                            'examples needed
                                            Dim measures() As Boolean = Calculation.GetMeasures(side, "stutzen")
                                            Dim rotangle As Double = 0
                                            If header.Tube.HeaderType = "outlet" Then
                                                rotangle = Math.Round(Math.PI, 3)
                                            End If
                                            SetAngular(asmdoc, stutzenno, rotangle, measures, 1, 1, 3, 1)
                                        End If
                                    End If

                                    If (.SpecialTag.Contains("s1star") Or .SpecialTag = "s2star" Or .SpecialTag = "s6star" Or .SpecialTag.Contains("s4star")) AndAlso .Figure = 4 Then
                                        Dim measures() As Boolean = Calculation.GetMeasures(side, "stutzen")
                                        Dim rotangle As Double
                                        Dim dishor As Double = header.Displacehor
                                        Dim disver As Double = header.Displacever
                                        If .SpecialTag = "s2star" Then
                                            If circuit.ConnectionSide = "right" Then
                                                dishor = 25
                                            Else
                                                dishor = -25
                                            End If
                                            disver = -50
                                        ElseIf .SpecialTag = "s6star" Or .SpecialTag = "s4starE2" Then
                                            disver = 25
                                        ElseIf .SpecialTag = "s4starF1" Then
                                            disver = -25
                                        End If

                                        If header.Ylist(i) = header.Ylist.Min And .SpecialTag <> "s2star" Then
                                            rotangle = Calculation.GetAngleRad(ctorigin(0), ctorigin(2), ctorigin(0) + dishor, ctorigin(2) + 25, side)
                                        Else
                                            rotangle = Calculation.GetAngleRad(ctorigin(0), ctorigin(2), ctorigin(0) + dishor, ctorigin(2) + disver, side)
                                        End If

                                        SetAngular(asmdoc, stutzenno, rotangle, measures, 1, 1, 3, 1)
                                    ElseIf .SpecialTag = "" OrElse (General.currentunit.ModelRangeName = "GACV" And .Figure = 45) Then
                                        Dim normalsaligned As Boolean = GetNormals(.ABV, stutzenalign, side, consys.HeaderAlignment, .Figure, circuit.ConnectionSide)

                                        If side = "back" Then
                                            normalsaligned = Not normalsaligned
                                        End If

                                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(stutzenno), fixedplaneno, 1, Not normalsaligned, 0, True)
                                    End If

                                End With

                                For Each dupe In similarstutzen
                                    Dim origin() As Double = Calculation.GetPartOrigin(header.Xlist(dupe), header.Ylist(dupe), side, circuit.CoreTubeOverhang, 0, 0)
                                    Dim copyno As Integer = GetTubeNoNew(origin, consys.CoreTubes.First, consys.HeaderAlignment)
                                    If copyno > 0 Then
                                        ctlist.Add(copyno)
                                    End If
                                Next

                                If similarstutzen.Count > 0 Then
                                    Dim pattername As String = header.Tube.HeaderType + patterncount.ToString
                                    CreateDuplicates(asmdoc, stutzenno, tubeno, ctlist, pattername)

                                    patterncount += 1
                                    usednos.AddRange(similarstutzen.ToArray)
                                End If
                            End If
                        End If

                        usednos.Add(i)
                    End If

                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetReversePlanar(asmdoc As SolidEdgeAssembly.AssemblyDocument, fixedno As Integer, adjno As Integer, offset As Double, normalsaligned As Boolean)
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim fixedocc, adjocc As SolidEdgeAssembly.Occurrence
        Dim fixedpart, adjpart As SolidEdgePart.PartDocument
        Dim fixedrefplane, adjrefplane As SolidEdgeGeometry.Plane
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim fixedface, adjface As SolidEdgeGeometry.Face
        Dim fixedreference, adjreference, objfixedface, objadjface As Object
        Dim fixedpoint(), adjpoint(), minrange(2), maxrange(2) As Double

        Try
            objoccs = asmdoc.Occurrences
            fixedocc = objoccs.Item(fixedno)
            fixedpart = fixedocc.PartDocument

            adjocc = objoccs.Item(adjno)
            adjpart = adjocc.PartDocument

            objrelations = asmdoc.Relations3d

            'get the front face of both parts 
            objfixedface = SEPart.GetFace(fixedpart, "coretube")
            fixedface = TryCast(objfixedface, SolidEdgeGeometry.Face)

            objadjface = SEPart.GetFace(adjpart, "stutzen")
            adjface = TryCast(objadjface, SolidEdgeGeometry.Face)

            If fixedface IsNot Nothing And adjface IsNot Nothing Then
                'convert the face into a plane
                fixedrefplane = TryCast(fixedface.Geometry, SolidEdgeGeometry.Plane)
                adjrefplane = TryCast(adjface.Geometry, SolidEdgeGeometry.Plane)

                fixedface.GetRange(minrange, maxrange)
                fixedpoint = minrange
                adjpoint = {0, 0, 0}

                'create the reference in the assembly
                fixedrefplane.GetRootPoint(fixedpoint)
                fixedreference = asmdoc.CreateReference(fixedocc, fixedface)

                adjrefplane.GetRootPoint(adjpoint)
                adjreference = asmdoc.CreateReference(adjocc, adjface)

                'create the planar relation
                objplanar = objrelations.AddPlanar(fixedreference, adjreference, normalsaligned, fixedpoint, adjpoint)
                objplanar.Offset = offset / 1000
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetFaceNo(asmdoc As SolidEdgeAssembly.AssemblyDocument, occno As Integer, fin As String, orientation As String, onebranch As Boolean) As Integer
        Dim faceno As Integer = 0
        Dim objocc As SolidEdgeAssembly.Occurrence
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
            objocc = asmdoc.Occurrences.Item(occno)
            partdoc = objocc.PartDocument
            partmodels = partdoc.Models
            partmodel = partmodels.Item(1)
            objbody = partmodel.Body
            objloops = objbody.Loops

            If onebranch Then
                'check the diameter from variable table
                diameter = SEPart.GetSetVariableValue("Diameter", partdoc.Variables, "get")
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
            General.CreateLogEntry(ex.ToString)
        End Try

        Return faceno
    End Function

    Shared Function GetNormals(abv As Double, stutzenalign As String, side As String, headeralign As String, figure As Integer, conside As String) As Boolean
        Dim aligned As Boolean

        'relative position at first
        If abv > 0 Then
            'relative position = less (header higher / further right)
            aligned = False
        Else
            'relative position = more (header lower / further left)
            aligned = True
        End If
        If headeralign = "horizontal" Then
            aligned = Not aligned
        End If

        'switch if stutzen is fig 5 and reversed
        If stutzenalign = "reverse" Then
            aligned = Not aligned
        End If

        'switch if vertical unit with outlet on bending side
        If side.Contains("back") And headeralign = "vertical" Then
            If aligned Then
                aligned = False
            Else
                aligned = True
            End If
        End If

        If General.currentunit.ApplicationType = "Evaporator" And (figure = 4 Or figure = 45) And conside = "right" Then
            If abv <= 0 Then
                'aligned = True
            ElseIf figure = 45 Then
                aligned = True
            End If
        End If

        Return aligned
    End Function

    Shared Function GroupExists(asmdoc As SolidEdgeAssembly.AssemblyDocument, groupname As String, isexact As Boolean) As Boolean
        Dim gexists As Boolean = False

        Try
            For Each asmgroup As SolidEdgeAssembly.AssemblyGroup In asmdoc.AssemblyGroups
                If (isexact And asmgroup.Name = groupname) Or (Not isexact And asmgroup.Name.Contains(groupname)) Then
                    gexists = True
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try
        Return gexists
    End Function

    Shared Sub FillCoretubePositions(asmdoc As SolidEdgeAssembly.AssemblyDocument, ctname As String, ByRef ctdata As CoreTubeLocations)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences

        Try
            If ctdata.Shortname <> General.GetShortName(ctname) Then
                ctdata.Shortname = General.GetShortName(ctname)

                For Each occ As SolidEdgeAssembly.Occurrence In objoccs
                    If occ.Name.Contains(ctdata.Shortname) Then
                        Dim origin() As Double = GetOccOrigin(occ)

                        DoCTList(ctdata.Xlist, ctdata.Xindizes, origin(0), occ.Index)
                        DoCTList(ctdata.Ylist, ctdata.Yindizes, origin(1), occ.Index)
                        DoCTList(ctdata.Zlist, ctdata.Zindizes, origin(2), occ.Index)
                    End If
                Next

            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub DoCTList(ByRef coordlist As List(Of Double), ByRef indizes As List(Of List(Of Integer)), coord As Double, index As Integer)
        Dim cindex As Integer = coordlist.IndexOf(coord)

        If cindex = -1 Then
            coordlist.Add(coord)
            indizes.Add(New List(Of Integer) From {index})
        Else
            indizes(cindex).Add(index)
        End If
    End Sub

    Shared Sub AddPlates(asmdoc As SolidEdgeAssembly.AssemblyDocument, occlist As List(Of SolidEdgeAssembly.Occurrence), platelist As List(Of PlateData), stutzenlist As List(Of StutzenData),
                         workspace As String, materialcode As String, ByRef subocclist As List(Of PartData), coretubeoverhang As Double)
        Dim firstocc, copyocc As SolidEdgeAssembly.Occurrence
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim fixedfaces(), adjfaces() As SolidEdgeGeometry.Face
        Dim totalocclist As New List(Of List(Of SolidEdgeAssembly.Occurrence))
        Dim totalstutzenlist As New List(Of List(Of StutzenData))
        Dim currentplate As PlateData
        Dim todolists As New List(Of List(Of Integer))
        Dim plateno, startno, loopcount, no As Integer
        Dim copylist, removelist As New List(Of Integer)
        Dim firstoccid, pattername, alignment As String
        Dim normals As Boolean
        Dim offset As Double

        Try
            'split the stutzen and occlist, so each plate has its own lists
            Dim sortedplates = platelist.OrderBy(Function(c) c.InnerDiameter)
            startno = 0
            For i As Integer = 0 To sortedplates.Count - 1
                currentplate = sortedplates(i)
                Dim tempstutzen As New List(Of StutzenData)
                Dim tempoccs As New List(Of SolidEdgeAssembly.Occurrence)
                Dim temptodo As New List(Of Integer)
                For j As Integer = startno To currentplate.Quantity - 1 + startno
                    tempstutzen.Add(stutzenlist(j))
                    tempoccs.Add(occlist(j))
                    temptodo.Add(j - startno)
                Next
                totalstutzenlist.Add(tempstutzen)
                totalocclist.Add(tempoccs)
                todolists.Add(temptodo)
                startno += tempoccs.Count
            Next

            objoccs = asmdoc.Occurrences
            loopcount = 1

            For i As Integer = 0 To sortedplates.Count - 1
                currentplate = sortedplates(i)
                Do
                    'add the plate
                    plateno = AddOccurance(asmdoc, currentplate.ID, workspace, "par", subocclist, "inlet")

                    'assemble with first occlist, that is left in the todolist
                    firstocc = totalocclist(i)(todolists(i).Min)
                    alignment = GetStutzenOrientationforPlate(firstocc, coretubeoverhang)
                    fixedfaces = SEPart.GetFaces(firstocc, "Stutzen", "", 5, alignment)

                    If materialcode = "C" Then
                        adjfaces = SEPart.GetFaces(objoccs.Item(plateno), "CUplate", "", 0, "")
                        normals = False
                        offset = -5
                    Else
                        'figure is not needed correctly, as long as its not 3
                        adjfaces = SEPart.GetFaces(objoccs.Item(plateno), "plate", "", 0, "")
                        normals = True
                        offset = 0
                    End If
                    SetAxialFaces(asmdoc, firstocc, objoccs.Item(plateno), fixedfaces(1), adjfaces(1), normals, True)
                    SetPlanarFaces(asmdoc, firstocc, objoccs.Item(plateno), fixedfaces(0), adjfaces(0), True, offset)

                    firstoccid = totalstutzenlist(i)(todolists(i).Min).ID

                    'in case of only 1 left no need for a pattern
                    If todolists(i).Count > 1 Then
                        For n As Integer = 1 To todolists(i).Count - 1
                            no = todolists(i)(n)
                            If totalstutzenlist(i)(no).ID = firstoccid Then
                                copyocc = totalocclist(i)(no)
                                copylist.Add(copyocc.Index)
                                removelist.Add(no)
                            End If
                        Next
                        todolists(i).Remove(todolists(i).Min)
                        For Each remno In removelist
                            todolists(i).Remove(remno)
                        Next
                    Else
                        todolists(i).Remove(todolists(i).Min)
                    End If

                    If copylist.Count > 0 Then
                        pattername = "plate" + loopcount.ToString
                        CreateDuplicates(asmdoc, plateno, firstocc.Index, copylist, pattername)
                    End If
                    copylist.Clear()
                    removelist.Clear()
                    loopcount += 1
                    'if todolists(i) still has entries, then repeat with next ID
                Loop Until todolists(i).Count = 0
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetStutzenOrientationforPlate(stutzenocc As SolidEdgeAssembly.Occurrence, coretubeoverhang As Double) As String
        Dim x, y, z, y0 As Double
        Dim alignment As String

        y0 = coretubeoverhang - 5
        stutzenocc.GetOrigin(x, y, z)

        y = Math.Round(y * 1000, 3)

        'opposite of stutzen placement, open end needed
        If y = -y0 Then
            alignment = "reverse"
        Else
            alignment = "normal"
        End If

        Return alignment
    End Function

    Shared Sub SetPlanarFaces(asmdoc As SolidEdgeAssembly.AssemblyDocument, fixedocc As SolidEdgeAssembly.Occurrence, adjocc As SolidEdgeAssembly.Occurrence, fixedface As SolidEdgeGeometry.Face,
                             adjface As SolidEdgeGeometry.Face, normals As Boolean, offset As Double)
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim fixedrefplane, adjrefplane As SolidEdgeGeometry.Plane
        Dim firstplref, secondplref As Object
        Dim firstpoint(), secondpoint() As Double

        firstpoint = {0, 0, 0}
        secondpoint = {0, 0, 0}

        Try
            fixedrefplane = TryCast(fixedface.Geometry, SolidEdgeGeometry.Plane)
            fixedrefplane.GetRootPoint(firstpoint)
            adjrefplane = TryCast(adjface.Geometry, SolidEdgeGeometry.Plane)
            adjrefplane.GetRootPoint(secondpoint)

            firstplref = asmdoc.CreateReference(fixedocc, fixedface)
            secondplref = asmdoc.CreateReference(adjocc, adjface)

            objrelations = asmdoc.Relations3d
            objplanar = objrelations.AddPlanar(firstplref, secondplref, normals, firstpoint, secondpoint)

            objplanar.Offset = offset / 1000

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AddNippleTube(asmdoc As SolidEdgeAssembly.AssemblyDocument, nippletube As TubeData, position As Double, ByRef occlist As List(Of PartData), header As HeaderData, circuit As CircuitData, headeralignment As String)
        Dim headerocc As SolidEdgeAssembly.Occurrence
        Dim nippleno As Integer
        Dim normals() As Boolean
        Dim offsets() As Double = {position, 0, 0}
        Dim headerplanes() As Integer = {1, 2, 3}
        Dim nippleplanes() As Integer

        Try
            nippleno = AddOccurance(asmdoc, nippletube.FileName, General.currentjob.Workspace, "par", occlist, nippletube.HeaderType)
            headerocc = asmdoc.Occurrences.Item(header.Tube.FileName + ":1")

            If nippletube.TubeType = "adapter" Then
                'adapter needs the opposite angle - planeno relation than nipple
                nippleplanes = GetNipplePlanes(Math.Abs(nippletube.Angle) - 90)
                'normals are shifted by 90°
                normals = GetNormalsByAngle(nippletube.Angle + 90)
                If nippletube.Angle = 0 Or nippletube.Angle = 180 Then
                    offsets = {position, 0, GNData.GetAdapterOffset(nippletube.BottomCapID, "adapter")}
                Else
                    offsets = {position, GNData.GetAdapterOffset(nippletube.BottomCapID, "adapter"), 0}
                End If

                If nippletube.Angle = -90 Then
                    offsets(1) = -offsets(1)
                ElseIf nippletube.Angle = 180 Then
                    offsets(2) = -offsets(2)
                End If
            Else
                nippleplanes = GetNipplePlanes(nippletube.Angle)
                normals = GetNormalsByAngle(nippletube.Angle)
            End If

            For i As Integer = 1 To 3
                SetPlanar(asmdoc, headerocc, asmdoc.Occurrences.Item(nippleno), i, nippleplanes(i - 1), normals(i - 1), offsets(i - 1), True)
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetNipplePlanes(angle As Integer) As Integer()
        Dim nippleplanes() As Integer

        If angle = 0 Or angle = 180 Then
            nippleplanes = {1, 2, 3}
        Else
            nippleplanes = {1, 3, 2}
        End If

        Return nippleplanes
    End Function

    Shared Function GetNormalsByAngle(angle As Integer) As Boolean()
        Dim normals() As Boolean

        Select Case angle
            Case -90
                normals = {False, True, False}
            Case 0
                normals = {False, False, False}
            Case 90
                normals = {False, False, True}
            Case 180
                normals = {False, True, True}
            Case Else
                normals = {False, True, False}
        End Select
        Return normals
    End Function

    Shared Sub AssembleAdapterNipple(asmdoc As SolidEdgeAssembly.AssemblyDocument, nippletube As TubeData, ByRef occlist As List(Of PartData))
        Dim nippleno, adapterno As Integer
        Dim nippleoffset As Double

        Try
            nippleno = AddOccurance(asmdoc, nippletube.FileName, General.currentjob.Workspace, "par", occlist, nippletube.HeaderType)
            adapterno = nippleno - 1

            nippleoffset = GNData.GetAdapterOffset(GNData.GetAdapterID(nippletube.Diameter), "nipple")

            SetPlanar(asmdoc, asmdoc.Occurrences.Item(adapterno), asmdoc.Occurrences.Item(nippleno), 1, 1, normals:=False, offset:=0, fixed:=True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(adapterno), asmdoc.Occurrences.Item(nippleno), 2, 3, normals:=True, offset:=nippleoffset, fixed:=True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(adapterno), asmdoc.Occurrences.Item(nippleno), 3, 2, normals:=False, offset:=0, fixed:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AddTubeCap(asmdoc As SolidEdgeAssembly.AssemblyDocument, tube As TubeData, capID As String, ByRef consys As ConSysData, pressure As Integer, location As String, nippleno As Integer)
        Dim tubeocc As SolidEdgeAssembly.Occurrence
        Dim capaxface, headeraxface, headerplanface As SolidEdgeGeometry.Face
        Dim objplanface As Object
        Dim capno As Integer
        Dim tubeoccname As String

        Try
            If Not IO.File.Exists(General.GetFullFilename(General.currentjob.Workspace, capID, "par")) Then
                WSM.CheckoutPart(capID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                General.WaitForFile(General.currentjob.Workspace, capID, "par", 100)
            End If

            capno = AddOccurance(asmdoc, capID, General.currentjob.Workspace, "par", consys.Occlist, tube.HeaderType)

            If capno > 0 Then
                If tube.TubeType = "nipple" Then
                    tubeoccname = General.GetShortName(tube.FileName) + ":" + nippleno.ToString
                Else
                    tubeoccname = General.GetShortName(tube.FileName) + ":1"
                End If
                tubeocc = asmdoc.Occurrences.Item(tubeoccname)

                'get faces
                capaxface = SEPart.GetCapAxFace(asmdoc, tube.Diameter)
                headeraxface = SEPart.GetHeaderFace(tubeocc, "axial")
                If capaxface IsNot Nothing Then
                    'set axial relation
                    SetCapAxial(asmdoc, tubeocc.Index, capno, capaxface, headeraxface)
                End If

                'get header face for planar relation
                headerplanface = SEPart.GetHeaderFace(tubeocc, location)

                'get cap face, depending of cap type either face or refplane
                objplanface = SEPart.GetCapPlanFace(asmdoc.Occurrences.Item(capno).PartDocument, tube.Diameter, pressure, tube.Materialcodeletter)

                If objplanface IsNot Nothing Then
                    'set planar relation
                    SetCapPlanar(asmdoc, tubeocc.Index, capno, headerplanface, objplanface, 0)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetCapAxial(asmdoc As SolidEdgeAssembly.AssemblyDocument, headerno As Integer, capno As Integer, capface As SolidEdgeGeometry.Face, headerface As SolidEdgeGeometry.Face)
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim headerocc, capocc As SolidEdgeAssembly.Occurrence
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim objheaderref, objcapref As Object


        Try
            objrelations = asmdoc.Relations3d
            objoccs = asmdoc.Occurrences

            headerocc = objoccs.Item(headerno)
            capocc = objoccs.Item(capno)

            objheaderref = asmdoc.CreateReference(headerocc, headerface)
            objcapref = asmdoc.CreateReference(capocc, capface)

            objaxial = objrelations.AddAxial(objheaderref, objcapref, False)
            objaxial.FixedRotate = True

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub SetCapPlanar(asmdoc As SolidEdgeAssembly.AssemblyDocument, headerno As Integer, capno As Integer, headerface As SolidEdgeGeometry.Face, capface As Object, offset As Double)
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim headerocc, capocc As SolidEdgeAssembly.Occurrence
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim headerplane, capplane As SolidEdgeGeometry.Plane
        Dim caprefplane As SolidEdgePart.RefPlane
        Dim caprefface As SolidEdgeGeometry.Face
        Dim objheaderref, objcapref As Object
        Dim minrange(2), maxrange(2), headerpoint(), cappoint() As Double
        Dim normalsaligned As Boolean

        Try
            objcapref = Nothing
            objrelations = asmdoc.Relations3d
            objoccs = asmdoc.Occurrences

            headerocc = objoccs.Item(headerno)
            capocc = objoccs.Item(capno)

            'get root point for relation
            headerplane = TryCast(headerface.Geometry, SolidEdgeGeometry.Plane)
            headerface.GetRange(minrange, maxrange)
            headerpoint = minrange
            headerplane.GetRootPoint(headerpoint)
            objheaderref = asmdoc.CreateReference(headerocc, headerface)

            caprefplane = TryCast(capface, SolidEdgePart.RefPlane)
            If caprefplane IsNot Nothing Then
                normalsaligned = False
                cappoint = {0, 0, 0}
                caprefplane.GetRootPoint(cappoint)
                objcapref = asmdoc.CreateReference(capocc, caprefplane)
            Else
                'short break for another trycast
                Threading.Thread.Sleep(500)

                caprefface = TryCast(capface, SolidEdgeGeometry.Face)
                If caprefface IsNot Nothing Then
                    'same procedure as headerface
                    capplane = TryCast(caprefface.Geometry, SolidEdgeGeometry.Plane)
                    normalsaligned = True
                    caprefface.GetRange(minrange, maxrange)
                    cappoint = minrange
                    capplane.GetRootPoint(cappoint)
                    objcapref = asmdoc.CreateReference(capocc, caprefface)
                End If
            End If

            If objcapref IsNot Nothing Then
                objplanar = objrelations.AddPlanar(objheaderref, objcapref, normalsaligned, headerpoint, cappoint)
                If offset <> 0 Then
                    objplanar.Offset = offset / 1000
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AddCapillaryTube(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef consys As ConSysData, headertube As TubeData, ctmaterial As String)
        Dim tubeID As String
        Dim tubeno As Integer

        Try
            If ctmaterial = "C" Then
                tubeID = "0000903626"
            Else
                tubeID = "0000896642"
            End If
            WSM.CheckoutPart(tubeID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            Dim headername As String = headertube.FileName + ":1"
            tubeno = AddOccurance(asmdoc, tubeID, General.currentjob.Workspace, "par", consys.Occlist, "outlet")
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(headername), asmdoc.Occurrences.Item(tubeno), 1, 1, False, headertube.Length / 2, True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(headername), asmdoc.Occurrences.Item(tubeno), 2, 2, False, 0, True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(headername), asmdoc.Occurrences.Item(tubeno), 3, 3, False, headertube.Diameter / 2, True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AddSV(asmdoc As SolidEdgeAssembly.AssemblyDocument, tube As TubeData, SVID As String, ByRef consys As ConSysData, conside As String, Optional nippleno As Integer = 1)
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim tubeocc As SolidEdgeAssembly.Occurrence
        Dim svno, mp As Integer
        Dim svfeature As SolidEdgePart.ExtrudedCutout
        Dim tubefeature As SolidEdgePart.ExtrudedProtrusion
        Dim objexpro As SolidEdgePart.ExtrudedProtrusion
        Dim tubedoc, svdoc As SolidEdgePart.PartDocument
        Dim fixedfaces, svfaces As SolidEdgeGeometry.Faces
        Dim fixedface, svaxface, svplface As SolidEdgeGeometry.Face
        Dim fixedaxref, svaxref, fixedplref, svplref As Object
        Dim objparentplane As SolidEdgePart.RefPlane
        Dim objparentprofil As SolidEdgePart.Profile
        Dim planarnormalsaligned As Boolean = False
        Dim axialnormalsaligned As Boolean = False
        Dim planaroffset, fixedpoint(), svpoint() As Double


        Try
            svno = AddOccurance(asmdoc, SVID, General.currentjob.Workspace, "par", consys.Occlist, tube.HeaderType)
            Try
                tubeocc = asmdoc.Occurrences.Item(tube.TubeFile.Shortname + ":" + nippleno.ToString)
            Catch ex As Exception
                tubeocc = asmdoc.Occurrences.Item(tube.TubeFile.Shortname.Replace("Nipple1", "Nipple" + nippleno.ToString) + ":1")
            End Try
            tubedoc = tubeocc.PartDocument

            svfeature = SEPart.GetFeature(tubedoc, "svhole", "cutout")
            If svfeature.Suppress OrElse svfeature.Status = SolidEdgePart.FeatureStatusConstants.igFeatureFailed Then
                Dim featurename As String
                If tube.TubeType = "header" Then
                    featurename = "Header"
                Else
                    featurename = "Nippletube"
                End If
                tubefeature = SEPart.GetFeature(tubedoc, featurename, "extrusion")
                fixedfaces = tubefeature.SideFaces
                fixedface = fixedfaces.Item(1)
                objparentplane = tubedoc.RefPlanes.Item(3)
            Else
                fixedfaces = svfeature.SideFaces
                fixedface = fixedfaces.Item(1)
                objparentprofil = svfeature.Profile
                objparentplane = objparentprofil.Plane
            End If

            fixedaxref = asmdoc.CreateReference(tubeocc, fixedface)

            fixedplref = asmdoc.CreateReference(tubeocc, objparentplane)


            If tube.SVPosition(1) = "axial" Then
                If tube.SVPosition(0) = "header" Then
                    'use nipple cutout for axial
                    If (conside = "left" And Not tube.IsBrine) Or (conside = "right" And tube.IsBrine) Then
                        planaroffset = -tube.Length + 4.5
                        planarnormalsaligned = True
                    Else
                        planaroffset = -4.5
                    End If
                ElseIf tube.SVPosition(0) = "flange" Then
                    planaroffset = -tube.Length
                    planarnormalsaligned = True
                    axialnormalsaligned = True
                Else
                    'position at the end of the nipple tube
                    planaroffset = tube.Length - 5
                    If Not tube.IsBrine Then
                        planaroffset = -planaroffset
                    End If
                    planarnormalsaligned = True
                End If
            Else
                mp = 1
                If consys.HeaderAlignment = "vertical" Then
                    mp = -1
                    planarnormalsaligned = True
                End If
                planaroffset = (tube.Diameter / 2 - 5) * mp
            End If

            svdoc = asmdoc.Occurrences.Item(svno).PartDocument
            objexpro = svdoc.Models.Item(1).ExtrudedProtrusions.Item(7)
            svfaces = objexpro.SideFaces
            svaxface = svfaces.Item(1)

            svaxref = asmdoc.CreateReference(asmdoc.Occurrences.Item(svno), svaxface)

            objrelations = asmdoc.Relations3d
            objaxial = objrelations.AddAxial(fixedaxref, svaxref, axialnormalsaligned)
            objaxial.FixedRotate = True

            svplface = objexpro.BottomCap

            svplref = asmdoc.CreateReference(asmdoc.Occurrences.Item(svno), svplface)

            fixedpoint = {0, 0, 0}
            svpoint = {0, 0, 0}

            objplanar = objrelations.AddPlanar(fixedplref, svplref, planarnormalsaligned, fixedpoint, svpoint)
            objplanar.Offset = -planaroffset / 1000

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AddSV2Cap(asmdoc As SolidEdgeAssembly.AssemblyDocument, svID As String, ByRef consys As ConSysData, headertype As String)
        Dim capocc, adjocc As SolidEdgeAssembly.Occurrence
        Dim fixedfaces(), adjfaces() As SolidEdgeGeometry.Face

        Try
            capocc = asmdoc.Occurrences.Item(asmdoc.Occurrences.Count)
            Dim svno As Integer = AddOccurance(asmdoc, svID, General.currentjob.Workspace, "par", consys.Occlist, headertype)
            adjocc = asmdoc.Occurrences.Item(svno)

            fixedfaces = SEPart.GetFaces(capocc, "cap", "fixed", 5, "normal")
            adjfaces = SEPart.GetFaces(adjocc, "SV", "adj", 3, "normal")

            'offset = -5
            SetAxialFaces(asmdoc, capocc, adjocc, fixedfaces(1), adjfaces(1), False, True)
            SetPlanarFaces(asmdoc, capocc, adjocc, fixedfaces(0), adjfaces(0), True, -5)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceVentsConsys(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim partno(header.VentIDs.Count - 1) As Integer
        Dim i, j, no, planenos() As Integer
        Dim counterpart As String
        Dim offset, ventoffset(), ventlength As Double
        Dim ventnormals() As Boolean

        Try
            i = 2
            j = 0

            'add muffe and sealing, don't add stopfen
            Do
                partno(j) = AddOccurance(asmdoc, header.VentIDs(i), General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)
                j += 1
                i -= 1
            Loop Until i < 1

            RemoveFromBOM(partno(1), asmdoc.Occurrences)

            'only CU checked so far
            Select Case header.Ventsize
                Case "G1/8"
                    offset = 2
                Case "G3/8"
                    offset = 2
                Case "G1/2"
                    If header.Tube.Materialcodeletter = "C" Then
                        offset = 22
                    Else
                        offset = 1
                    End If
                Case "G3/4"
                    offset = 15
                Case Else
                    offset = 20
            End Select

            If header.Tube.Materialcodeletter = "C" Then
                'assemble muffe and sealing
                AssembleMuffeSealing1(asmdoc, partno(0), partno(1), header.Ventsize, offset, "protrusion", 1, "C")
            Else
                'assemble muffe and sealing
                If header.Ventsize = "G1/4" Or header.Ventsize = "G1" Then
                    AssembleMuffeSealing2(asmdoc, partno(0), partno(1), 2, 1, 0, "C")
                Else
                    AssembleMuffeSealing2(asmdoc, partno(0), partno(1), 3, 1, 0, "C", ventsize:=header.Ventsize)
                End If
            End If

            If header.VentIDs.Count = 4 Then
                'only for certain GACV
                partno(2) = AddOccurance(asmdoc, header.VentIDs(3), General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)
                If header.VentIDs(3).Contains("898859") Then
                    counterpart = "header"
                    no = 2
                Else
                    counterpart = "fin"        '→ find coretube!
                    no = 1
                End If
                If header.Tube.Materialcodeletter = "C" Then
                    'assemble tube and muffe
                    AssembleMuffeSealing2(asmdoc, partno(0), partno(2), 1, no, -50, "C")
                    'assemble tube and header 
                    If counterpart = "header" Then
                        'outlet
                        AssembleTubeHeader(asmdoc, asmdoc.Occurrences.Item(General.GetShortName(header.Tube.FileName) + ":1").Index, partno(2))
                    Else
                        'inlet   
                        ventoffset = GNData.GetVentoffset(circuit.FinType, circuit.ConnectionSide, header.Tube.Materialcodeletter, circuit.CoreTubeOverhang)
                        ventnormals = GNData.GetVentnormals(circuit.ConnectionSide)
                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=1, secondplane:=1, normals:=False, offset:=ventoffset(2), True)
                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=2, secondplane:=3, normals:=ventnormals(1), offset:=ventoffset(0), True)
                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=3, secondplane:=2, normals:=ventnormals(2), offset:=ventoffset(1), True)
                    End If
                Else
                    ventoffset = GNData.GetVentoffset(circuit.FinType, circuit.ConnectionSide, header.Tube.Materialcodeletter, circuit.CoreTubeOverhang)
                    ventnormals = GNData.GetVentnormals(circuit.ConnectionSide)

                    If circuit.FinType = "F" And header.Tube.HeaderType = "inlet" And header.Tube.Materialcodeletter = "V" Then
                        'different connection piece (T-shape)
                        AssembleMuffeSealing2(asmdoc, partno(0), partno(2), 3, no, 26, header.Tube.Materialcodeletter, secondfeature:="cutout")
                        If circuit.ConnectionSide = "left" Then
                            ventoffset(1) += 5
                        Else
                            ventoffset(1) -= 25
                        End If
                    Else
                        'only for inlet, outlet always consists out of 3 parts 
                        'assemble tube and muffe // for VA no = 1 !
                        AssembleMuffeSealing2(asmdoc, partno(0), partno(2), 3, no, 43, header.Tube.Materialcodeletter)
                    End If

                    'assemble tube and fin
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=1, secondplane:=1, normals:=False, offset:=ventoffset(2), True)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=2, secondplane:=3, normals:=ventnormals(1), offset:=ventoffset(0), True)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(1), asmdoc.Occurrences.Item(partno(2)), firstplane:=3, secondplane:=2, normals:=ventnormals(2), offset:=ventoffset(1), True)
                End If
            ElseIf header.Tube.IsBrine Then
                Dim ismirrored As Boolean = False
                Dim mp As Double = 1
                If circuit.ConnectionSide = "right" Then
                    ismirrored = True
                    mp = -1
                End If
                ventlength = GNData.GetVentLength(header.Ventsize, header.Tube.Materialcodeletter)
                'assemble muffe and brine header
                If consys.HeaderAlignment = "vertical" Then
                    If header.Tube.Materialcodeletter = "V" Then
                        planenos = {1, 2, 3}
                        ventnormals = {False, True, True}
                        ventoffset = {0, header.Tube.Diameter / 2 + ventlength, header.Ventposition}
                    Else
                        If header.Tube.HeaderType = "inlet" Then
                            planenos = {2, 3, 1}
                            ventnormals = {False, True, True}
                            ventoffset = {header.Ventposition, 0, -ventlength * mp}
                        Else
                            planenos = {2, 1, 3}
                            ventnormals = {False, Not ismirrored, ismirrored}
                            ventoffset = {header.Ventposition, -ventlength * mp, 0}
                        End If
                    End If
                Else
                    If header.Tube.Materialcodeletter = "C" Then
                        planenos = {3, 2, 1}
                        ventoffset = {28, 0, -32}
                        ventnormals = {False, False, True}
                        ventoffset(0) = header.Ventposition
                    Else
                        planenos = {2, 1, 3}
                        ventnormals = {False, true, false}
                        ventoffset = {28, 0, -ventlength * 1.35}
                        ventoffset(0) = header.Ventposition
                    End If
                End If

                For k As Integer = 1 To 3
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(General.GetShortName(header.Tube.FileName) + ":1"), asmdoc.Occurrences.Item(partno(0)), k, planenos(k - 1), ventnormals(k - 1), ventoffset(k - 1), True)
                Next
            ElseIf circuit.IsOnebranchEvap Then

            Else
                'normal Condenser & Evaporators
                'assemble muffe and header
                AssembleVent4(asmdoc, asmdoc.Occurrences.Item(General.GetShortName(header.Tube.FileName) + ":1"), partno(0), header, circuit.FinType, circuit.ConnectionSide)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub RemoveFromBOM(partno As Integer, objoccs As SolidEdgeAssembly.Occurrences)
        Dim objocc As SolidEdgeAssembly.Occurrence = objoccs.Item(partno)
        objocc.IncludeInBom = False
    End Sub

    Shared Sub AssembleMuffeSealing1(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstno As Integer, secondno As Integer, ventsize As String, offset As Double, secondtype As String, planeno As Integer, material As String)
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim firstocc, secondocc As SolidEdgeAssembly.Occurrence
        Dim firstaxis, secondaxis As SolidEdgePart.RefAxis
        Dim firstaxobj, secondaxobj As Object

        Try
            objoccs = asmdoc.Occurrences
            objrelations = asmdoc.Relations3d

            firstocc = objoccs.Item(firstno)
            secondocc = objoccs.Item(secondno)

            firstaxis = SEPart.GetRotationAxis(firstocc.PartDocument, "cutout")
            secondaxis = SEPart.GetRotationAxis(secondocc.PartDocument, secondtype)

            If firstaxis Is Nothing Then
                firstaxis = SEPart.GetRotationAxis(firstocc.PartDocument, "revpro")
            End If

            If secondaxis Is Nothing Then
                secondaxis = SEPart.GetRotationAxis(secondocc.PartDocument, "revpro")
            End If

            If firstaxis IsNot Nothing And secondaxis IsNot Nothing Then
                firstaxobj = asmdoc.CreateReference(firstocc, firstaxis)
                secondaxobj = asmdoc.CreateReference(secondocc, secondaxis)

                objaxial = objrelations.AddAxial(firstaxobj, secondaxobj, False)
                objaxial.FixedRotate = True
            End If

            If ventsize = "G1" Then
                If General.currentunit.ApplicationType = "Condenser" Then
                    SetPlanar(asmdoc, firstocc, secondocc, 1, 1, False, offset, True)
                Else
                    If material = "C" Then
                        SetPlanar(asmdoc, firstocc, secondocc, 2, 1, False, offset, True)
                    Else
                        SetPlanar(asmdoc, firstocc, secondocc, 2, 2, False, offset, True)
                    End If
                End If
            ElseIf ventsize = "G1/4" Then
                SetPlanar(asmdoc, firstocc, secondocc, 2, 2, False, offset, True)
            Else
                SetPlanar(asmdoc, firstocc, secondocc, planeno, planeno, False, offset, True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AssembleMuffeSealing2(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstno As Integer, secondno As Integer, firstplno As Integer, no As Integer, offset As Double, headermat As String,
                                     Optional secondplno As Integer = 3, Optional secondfeature As String = "protrusion", Optional ventsize As String = "")
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim firstocc, secondocc As SolidEdgeAssembly.Occurrence
        Dim firstaxis, secondaxis As SolidEdgePart.RefAxis
        Dim secondface As SolidEdgeGeometry.Face
        Dim secondplane As SolidEdgeGeometry.Plane
        Dim firstplane As SolidEdgePart.RefPlane
        Dim firstaxobj, secondaxobj, objface, firstplref, secondplref As Object
        Dim firstpoint(), secondpoint() As Double

        Try
            objoccs = asmdoc.Occurrences
            objrelations = asmdoc.Relations3d

            firstocc = objoccs.Item(firstno)
            secondocc = objoccs.Item(secondno)

            firstaxis = SEPart.GetRotationAxis(firstocc.PartDocument, "cutout")
            If firstaxis Is Nothing Then
                firstaxis = SEPart.GetRotationAxis(firstocc.PartDocument, "protrusion")
            End If
            secondface = SEPart.GetSideFace(secondocc.PartDocument, secondfeature, no, 1)
            If secondface Is Nothing Then
                secondaxis = SEPart.GetRotationAxis(secondocc.PartDocument, "protrusion")
            End If

            If firstaxis IsNot Nothing And (secondface IsNot Nothing Or secondaxis IsNot Nothing) Then
                firstaxobj = asmdoc.CreateReference(firstocc, firstaxis)
                If secondface Is Nothing Then
                    secondaxobj = asmdoc.CreateReference(secondocc, secondaxis)
                Else
                    secondaxobj = asmdoc.CreateReference(secondocc, secondface)
                End If

                objaxial = objrelations.AddAxial(firstaxobj, secondaxobj, False)
                objaxial.FixedRotate = True
            End If

            If no = 1 Then
                If ventsize = "G2" Then
                    SetPlanar(asmdoc, firstocc, secondocc, 2, 2, True, offset, True)
                ElseIf headermat = "C" Then
                    SetPlanar(asmdoc, firstocc, secondocc, firstplno, secondplno, False, offset, True)
                Else
                    SetPlanar(asmdoc, firstocc, secondocc, firstplno, 3, True, offset, True)
                End If
            Else
                'use reverse planar method
                objface = SEPart.GetFace(secondocc.PartDocument, "stutzen")
                If objface IsNot Nothing Then
                    secondface = TryCast(objface, SolidEdgeGeometry.Face)
                    If secondface IsNot Nothing Then
                        'create the reference etc
                        secondplane = TryCast(secondface.Geometry, SolidEdgeGeometry.Plane)
                        secondpoint = {0, 0, 0}

                        secondplane.GetRootPoint(secondpoint)
                        secondplref = asmdoc.CreateReference(secondocc, secondface)

                        firstplane = SEPart.GetRefPlane(firstocc.PartDocument, planenumber:=1)

                        firstpoint = {0, 0, 0}
                        firstplane.GetRootPoint(firstpoint)

                        firstplref = asmdoc.CreateReference(firstocc, firstplane)

                        objplanar = objrelations.AddPlanar(firstplref, secondplref, False, firstpoint, secondpoint)
                        objplanar.Offset = -16 / 1000

                    End If
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AssembleTubeHeader(asmdoc As SolidEdgeAssembly.AssemblyDocument, firstno As Integer, secondno As Integer)
        Dim firstdoc As SolidEdgePart.PartDocument
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim firstocc, secondocc As SolidEdgeAssembly.Occurrence
        Dim ventcutout As SolidEdgePart.ExtrudedCutout
        Dim objfaces As SolidEdgeGeometry.Faces
        Dim firstface, secondface As SolidEdgeGeometry.Face
        Dim firstplane, secondplane As SolidEdgePart.RefPlane
        Dim firstaxobj, secondaxobj, firstplref, secondplref As Object
        Dim firstpoint(2), secondpoint(2), headerdiameter As Double


        'for outlet 
        Try
            objoccs = asmdoc.Occurrences
            objrelations = asmdoc.Relations3d

            firstocc = objoccs.Item(firstno)
            secondocc = objoccs.Item(secondno)

            ventcutout = SEPart.GetFeature(firstocc.PartDocument, "vent", "cutout")

            If ventcutout IsNot Nothing Then
                objfaces = ventcutout.SideFaces
                firstface = objfaces.Item(1)
                secondface = SEPart.GetSideFace(secondocc.PartDocument, "protrusion", 1, 1)

                firstaxobj = asmdoc.CreateReference(firstocc, firstface)
                secondaxobj = asmdoc.CreateReference(secondocc, secondface)

                objaxial = objrelations.AddAxial(firstaxobj, secondaxobj, False)

                'locking the rotation
                SetPlanar(asmdoc, firstocc, secondocc, firstplane:=1, secondplane:=2, normals:=True, offset:=0, fixed:=False)

                firstplane = SEPart.GetRefPlane(firstocc.PartDocument, planename:="Ventplane")
                secondplane = SEPart.GetRefPlane(secondocc.PartDocument, planenumber:=3)

                firstpoint = {0, 0, 0}
                secondpoint = {0, 0, 0}

                firstplref = asmdoc.CreateReference(firstocc, firstplane)
                secondplref = asmdoc.CreateReference(secondocc, secondplane)

                firstdoc = firstocc.PartDocument
                headerdiameter = SEPart.GetSetVariableValue("HeaderDiameter", firstdoc.Variables, "Get")

                objplanar = objrelations.AddPlanar(firstplref, secondplref, True, firstpoint, secondpoint)
                objplanar.Offset = -(headerdiameter / 2 - 7) / 1000

            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub AssembleVent4(asmdoc As SolidEdgeAssembly.AssemblyDocument, headerocc As SolidEdgeAssembly.Occurrence, ventno As Integer, header As HeaderData, fintype As String, conside As String)
        Dim headeraxref, headerplref, ventaxref As Object
        Dim ventocc As SolidEdgeAssembly.Occurrence
        Dim ventdoc, headerdoc As SolidEdgePart.PartDocument
        Dim objrelations As SolidEdgeAssembly.Relations3d
        Dim objaxial As SolidEdgeAssembly.AxialRelation3d
        Dim objplanar As SolidEdgeAssembly.PlanarRelation3d
        Dim ventfeatureplane As SolidEdgeGeometry.Plane
        Dim ventexpros As SolidEdgePart.ExtrudedProtrusions
        Dim ventexpro As SolidEdgePart.ExtrudedProtrusion
        Dim ventrevpro As SolidEdgePart.RevolvedProtrusion
        Dim ventaxis As SolidEdgePart.RefAxis
        Dim ventcutout As SolidEdgePart.ExtrudedCutout
        Dim ventface, ventsideface As SolidEdgeGeometry.Face
        Dim ventsidefaces As SolidEdgeGeometry.Faces
        Dim ventprofile As SolidEdgePart.Profile
        Dim ventplane, ventrefplane As SolidEdgePart.RefPlane
        Dim ventplref As Object
        Dim headerpoint(), ventpoint(), minrange(2), maxrange(2), offset As Double
        Dim normals As Boolean = False


        headerdoc = headerocc.PartDocument

        ventcutout = SEPart.GetFeature(headerdoc, "Vent", "cutout")

        If ventcutout IsNot Nothing Then
            'get the sideface as axial ref
            ventface = ventcutout.SideFaces(0)

            'get the parent plane as planar ref
            ventprofile = ventcutout.Profile
            ventplane = ventprofile.Plane

            'get the vent part
            ventocc = asmdoc.Occurrences.Item(ventno)
            ventdoc = ventocc.PartDocument

            'get the axis for axial ref
            ventaxis = SEPart.GetRotationAxis(ventdoc, "cutout")
            If ventaxis Is Nothing Then
                ventaxis = SEPart.GetRotationAxis(ventdoc, "revpro")
            End If
            offset = (header.Tube.Diameter - 7) / 2000

            'depending of size get object for planar ref
            If header.Ventsize = "G1" Or (header.Ventsize = "G1/2" And header.Tube.Materialcodeletter <> "C") Or (header.Ventsize = "G3/4" And General.currentunit.ApplicationType = "Condenser") Or header.Ventsize = "G2" Then
                If header.Ventsize = "G1/2" And header.Tube.Materialcodeletter <> "C" Then
                    ventrefplane = ventdoc.RefPlanes.Item(3)
                    offset = 0.06
                    normals = False
                Else
                    If General.currentunit.ApplicationType = "Condenser" Then
                        If header.Ventsize = "G2" Then
                            offset += header.Tube.Diameter / 2000
                            ventrefplane = ventdoc.RefPlanes.Item(2)
                        Else
                            offset = (header.Tube.Diameter + 7) / 2000
                            ventrefplane = ventdoc.RefPlanes.Item(1)
                        End If
                    Else
                        ventrefplane = ventdoc.RefPlanes.Item(2)
                    End If
                    normals = True
                End If
                ventpoint = {0, 0, 0}
                ventrefplane.GetRootPoint(ventpoint)
                ventplref = asmdoc.CreateReference(ventocc, ventrefplane)
            ElseIf header.Ventsize = "G1/4" And header.Tube.Materialcodeletter = "V" Then
                ventrevpro = ventdoc.Models.Item(1).RevolvedProtrusions.Item(1)
                offset += 5 / 1000
                ventaxis = ventrevpro.Axis
                ventsidefaces = ventrevpro.SideFaces
                ventsideface = ventsidefaces.Item(11)
                ventsideface.GetRange(minrange, maxrange)
                ventpoint = minrange
                ventfeatureplane = TryCast(ventsideface.Geometry, SolidEdgeGeometry.Plane)
                ventfeatureplane.GetRootPoint(ventpoint)
                ventplref = asmdoc.CreateReference(ventocc, ventsideface)
            Else
                'works for 3/8
                If header.Ventsize = "G1/4" Or ((header.Ventsize = "G1/2" Or (header.Ventsize = "G3/8" And header.Tube.Materialcodeletter = "C")) And General.currentunit.ApplicationType = "Condenser") Then
                    ventrevpro = ventdoc.Models.Item(1).RevolvedProtrusions.Item(1)
                    ventsidefaces = ventrevpro.SideFaces
                    ventsideface = ventsidefaces.Item(7)
                    offset -= 5 / 1000
                Else
                    If header.Ventsize = "G1/2" Then
                        ventrevpro = ventdoc.Models.Item(1).RevolvedProtrusions.Item(1)
                        ventsidefaces = ventrevpro.SideFaces
                        ventsideface = ventsidefaces.Item(7)
                        normals = True
                        offset += 22
                    Else
                        ventexpros = ventdoc.Models.Item(1).ExtrudedProtrusions
                        ventexpro = ventexpros.Item(2)
                        ventsideface = ventexpro.BottomCap
                    End If
                End If
                ventsideface.GetRange(minrange, maxrange)
                ventpoint = minrange
                ventfeatureplane = TryCast(ventsideface.Geometry, SolidEdgeGeometry.Plane)
                ventfeatureplane.GetRootPoint(ventpoint)
                ventplref = asmdoc.CreateReference(ventocc, ventsideface)
            End If

            'create ax reference for both parts
            headeraxref = asmdoc.CreateReference(headerocc, ventface)
            ventaxref = asmdoc.CreateReference(ventocc, ventaxis)

            'create pl reference for both parts
            headerpoint = {0, 0, 0}
            ventplane.GetRootPoint(headerpoint)
            headerplref = asmdoc.CreateReference(headerocc, ventplane)

            objrelations = asmdoc.Relations3d

            If header.Tube.HeaderType = "inlet" Then
                objaxial = objrelations.AddAxial(headeraxref, ventaxref, normals)
            Else
                objaxial = objrelations.AddAxial(headeraxref, ventaxref, True)
            End If

            normals = Not normals
            objaxial.FixedRotate = True

            objplanar = objrelations.AddPlanar(headerplref, ventplref, normals, headerpoint, ventpoint)
            objplanar.Offset = offset

        Else
            General.CreateLogEntry("Vent")
        End If

    End Sub

    Shared Sub PlaceThreads(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef consys As ConSysData, headertype As String)
        Dim objoccs As SolidEdgeAssembly.Occurrences = asmdoc.Occurrences
        Dim nippleocc, threadocc As SolidEdgeAssembly.Occurrence
        Dim threadno As Integer
        Dim values() As Double
        Dim erpcode As String = Database.GetValue("CSG.DB_Flanges", "ERPCode", "Article_Number", consys.FlangeID)

        Try
            nippleocc = objoccs.Item(objoccs.Count)
            threadno = AddOccurance(asmdoc, consys.FlangeID, General.currentjob.Workspace, "par", consys.Occlist, headertype)
            threadocc = objoccs.Item(threadno)

            'set axial
            Dim tubefeature As SolidEdgePart.ExtrudedProtrusion = SEPart.GetFeature(nippleocc.PartDocument, "ExtrudedProtrusion_1", "extrusion")
            Dim tubesidefaces As SolidEdgeGeometry.Faces = tubefeature.SideFaces
            Dim fixedaxface As SolidEdgeGeometry.Face = tubesidefaces.Item(1)

            Dim theaddoc As SolidEdgePart.PartDocument = threadocc.PartDocument
            Dim threadfeature As SolidEdgePart.ExtrudedProtrusion = theaddoc.Models.Item(1).ExtrudedProtrusions.Item(1)
            Dim threadsidefaces As SolidEdgeGeometry.Faces = threadfeature.SideFaces
            Dim adjaxface As SolidEdgeGeometry.Face = threadsidefaces.Item(1)

            SetAxialFaces(asmdoc, nippleocc, threadocc, fixedaxface, adjaxface, True, True)

            'set planar
            Dim fixedplface As SolidEdgeGeometry.Face = tubefeature.BottomCap
            Dim adjplface As SolidEdgeGeometry.Face

            If erpcode = "239" Then
                adjplface = threadfeature.BottomCap
            Else
                If erpcode = "" Then
                    erpcode = "236"
                End If
                adjplface = threadfeature.TopCap
            End If

            values = GNData.ThreadData(erpcode)

            SetPlanarFaces(asmdoc, nippleocc, threadocc, fixedplface, adjplface, True, -values(0))

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceFlanges(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef consys As ConSysData, nippletube As TubeData)
        Dim nippleocc As SolidEdgeAssembly.Occurrence
        Dim flangeno As Integer
        Dim offsets(), offset As Double

        Try

            nippleocc = asmdoc.Occurrences.Item(asmdoc.Occurrences.Count)

            'add flange to asm
            flangeno = AddOccurance(asmdoc, consys.FlangeID, General.currentjob.Workspace, "par", consys.Occlist, nippletube.HeaderType)

            If flangeno > 0 Then
                '3x planar relation to the nipple tube's ref planes, last relation uses offset 
                If consys.ConType = 2 Then
                    offset = nippletube.Length + 3 - (consys.FlangeDims.HF + consys.FlangeDims.SB + 2) - GNData.GetTubeOffset(nippletube.Diameter, nippletube.Materialcodeletter, General.currentjob.Plant) '+ consys.InletHeaders.First.Tube.Diameter / 2
                Else
                    'shell as adapter, has cutout 
                    If nippletube.Diameter < 50 Then
                        offset = nippletube.Length - 8
                    ElseIf nippletube.Diameter < 88 Then
                        offset = nippletube.Length - 10
                    Else
                        offset = nippletube.Length - 12
                    End If
                End If

                offsets = {0, 0, offset}

                For i As Integer = 1 To 3
                    SetPlanar(asmdoc, nippleocc, asmdoc.Occurrences.Item(flangeno), i, i, False, offsets(i - 1), True)
                Next
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleBranch(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim partno, tubeno, figure, faceno As Integer
        Dim abv As Double
        Dim stutzenalign As String
        Dim ismirrored As Boolean = False

        Try
            If circuit.ConnectionSide = "left" Then
                ismirrored = True
            End If
            partno = AddOccurance(asmdoc, header.StutzenDatalist.First.ID, General.currentjob.Workspace, "par", consys.Occlist)

            If partno > 0 Then
                Dim ctorigin() As Double = Calculation.GetPartOrigin(header.Xlist.First, header.Ylist.First, "front", circuit.CoreTubeOverhang, 0, 0)
                tubeno = GetTubeNoNew(ctorigin, consys.CoreTubes.First, "vertical")

                If tubeno > 0 Then
                    'fixed value for abv to avoid figure = 8, loopcount used for figure, result is either 5 or 4. 5 would be correct, so if it's something else, than must be figure 3
                    figure = SEPart.GetFigure(asmdoc, partno, 1)
                    If header.Tube.HeaderType = "outlet" And (circuit.FinType = "N" Or circuit.FinType = "M") Then
                        figure = 5
                    End If
                    If figure <> 5 Or header.StutzenDatalist.First.ID = "935732" Then
                        SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 3, 3, False, -5, True)
                        stutzenalign = "normal"
                        If header.Tube.HeaderType = "outlet" And consys.VType = "P" Then
                            If header.StutzenDatalist.First.ID = "935732" Then
                                SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 1, 1, True, offset:=0, True)
                            Else
                                SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 2, 2, True, offset:=0, True)
                            End If
                        Else
                            SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 2, 2, False, offset:=0, True)
                        End If
                    Else
                        abv = Math.Round(consys.InletHeaders.First.Dim_a - circuit.CoreTubeOverhang + 5)
                        stutzenalign = SEPart.GetStutzenAlignment(asmdoc, partno, abv, header, circuit.CoreTubeOverhang)
                        If stutzenalign = "reverse" Then
                            SetReversePlanar(asmdoc, tubeno, partno, -5, True)
                            SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 1, 1, Not ismirrored, offset:=0, True)
                        Else
                            SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 3, 3, False, -5, True)
                            SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 2, 2, ismirrored, offset:=0, True)
                        End If
                    End If

                    faceno = GetFaceNo(asmdoc, partno, circuit.FinType, orientation:=stutzenalign, onebranch:=True)
                    SetAxial(asmdoc, tubeno, partno, faceno)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleParts(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData, coil As CoilData)
        Dim vtype, SVID, madapterID, nadapterID, capID, partname, alignment As String
        Dim partorder As New List(Of String) From {"Stutzen"}
        Dim idorder As New List(Of String) From {header.StutzenDatalist.First.ID}
        Dim occlist As New List(Of SolidEdgeAssembly.Occurrence)
        Dim fixedno, adjno, figure As Integer
        Dim occorigin(2) As Double
        Dim offset, m, L3, L1, FPoffset() As Double
        Dim FPnormals() As Boolean
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim stutzenocc, adjocc As SolidEdgeAssembly.Occurrence
        Dim fixedfaces(), adjfaces() As SolidEdgeGeometry.Face
        Dim stutzendoc As SolidEdgePart.PartDocument

        Try
            If consys.VType = "X" Then
                vtype = "DX"
                If header.Tube.HeaderType = "inlet" Then
                    figure = 3
                Else
                    figure = 5
                End If
            Else
                vtype = "XP"
                If header.Tube.HeaderType = "outlet" And (circuit.FinType = "N" Or circuit.FinType = "M") Then
                    figure = 5
                Else
                    figure = 3
                End If
            End If

            'SV is always needed (currently)
            SVID = GNData.GetSVID(circuit.Pressure)
            WSM.CheckoutPart(SVID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            If consys.HeaderMaterial = "V" Or consys.HeaderMaterial = "W" Then
                If circuit.FinType = "E" Then
                    'get n adapter
                    nadapterID = "879661"
                    WSM.CheckoutPart(nadapterID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                    partorder.Add("nadapter")
                    idorder.Add(nadapterID)
                End If

                'get m adapter - not considering N fin
                If vtype = "DX" And header.Tube.HeaderType = "inlet" Then
                    If circuit.FinType = "N" Then
                        madapterID = "940933"
                    Else
                        madapterID = "738274"
                    End If
                Else
                    'm=f(finneddepth)
                    'different m for N FP
                    If circuit.FinType = "N" Then
                        If vtype = "DX" Then
                            madapterID = "940932"
                        Else
                            madapterID = "922264"
                        End If
                    Else
                        m = GNData.GetDimM(consys.VType, coil.FinnedHeight)
                        madapterID = GNData.GetIDM(m, circuit.Pressure)
                    End If
                End If
                If circuit.FinType = "N" Then
                    capID = "006198"
                Else
                    capID = "883642"
                End If
                WSM.CheckoutPart(capID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                WSM.CheckoutPart(madapterID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                partorder.AddRange({"nadapter", "cap", "SV"})
                idorder.AddRange({madapterID, capID, SVID})
            Else
                partorder.Add("SV")
                idorder.Add(SVID)
            End If

            objoccs = asmdoc.Occurrences

            'get stutzenno
            partname = header.StutzenDatalist.First.ID + "0001-.par:1"
            Try
                stutzenocc = objoccs.Item(partname)
            Catch ex As Exception
                'partname could be 2-.par
                partname = header.StutzenDatalist.First.ID + "0002-.par:1"
                stutzenocc = objoccs.Item(partname)
            End Try
            occlist.Add(stutzenocc)
            fixedno = stutzenocc.Index

            For i As Integer = 1 To idorder.Count - 1
                adjno = AddOccurance(asmdoc, idorder(i), General.currentjob.Workspace, "par", consys.Occlist)
                adjocc = objoccs.Item(adjno)
                occlist.Add(adjocc)

                If partorder(1) = "SV" And vtype = "XP" And consys.HeaderMaterial = "C" And circuit.FinType <> "N" Then
                    'different handling, using refplanes of both parts
                    'distance from tube sheet
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(fixedno), asmdoc.Occurrences.Item(adjno), 3, 3, False, 90, True)

                    stutzendoc = stutzenocc.PartDocument
                    L3 = SEPart.GetSetVariableValue("L3", stutzendoc.Variables, "get")
                    L1 = SEPart.GetSetVariableValue("L1", stutzendoc.Variables, "get")

                    offset = Math.Round(L3 + L1 * Math.Cos(45 * Math.PI / 180) - 17, 1)

                    If header.Tube.HeaderType = "inlet" Then
                        FPoffset = {-27, -offset}
                        FPnormals = {False, True}
                    Else
                        FPoffset = {-89, offset}
                        FPnormals = {True, False}
                    End If

                    'vertical distance
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(fixedno), asmdoc.Occurrences.Item(adjno), 1, 2, FPnormals(0), FPoffset(0), True)

                    If circuit.ConnectionSide = "left" Then
                        FPoffset(1) = -FPoffset(1)
                    End If
                    'distance from tube end, f(L3)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(fixedno), asmdoc.Occurrences.Item(adjno), 2, 1, FPnormals(1), FPoffset(1), True)
                Else
                    'get the face for axial method → input: partorder + function + relationtype (e.g. madapter + fixed + axial)
                    'because of N DX check the origin of outlet → if it is in the coil, then reverse
                    alignment = "normal"
                    If header.Tube.HeaderType = "outlet" And circuit.FinType = "N" And partorder(i - 1) = "Stutzen" Then
                        occlist(i - 1).GetOrigin(occorigin(0), occorigin(1), occorigin(2))
                        For n As Integer = 0 To occorigin.Count - 1
                            occorigin(n) = Math.Round(occorigin(n) * 1000)
                        Next
                        If (occorigin(0) > 0 And circuit.ConnectionSide = "right") Or (occorigin(0) < coil.FinnedDepth And circuit.ConnectionSide = "left") Then
                            alignment = "reverse"
                        End If
                    End If
                    'faces of already fixed part
                    fixedfaces = SEPart.GetFaces(occlist(i - 1), partorder(i - 1), "fixed", figure, alignment)

                    'faces of part to be fixed
                    adjfaces = SEPart.GetFaces(occlist(i), partorder(i), "adj", 3, "normal")

                    offset = GetPlOffset(partorder(i - 1) + "-" + partorder(i), circuit.FinType)

                    SetAxialFaces(asmdoc, occlist(i - 1), occlist(i), fixedfaces(1), adjfaces(1), False, True)

                    SetPlanarFaces(asmdoc, occlist(i - 1), occlist(i), fixedfaces(0), adjfaces(0), True, offset)
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetPlOffset(pairname As String, fintype As String) As Double
        Dim offset As Double

        If pairname.Contains("adapter") Then
            If pairname.Contains("cap") Then
                offset = 0
            Else
                offset = -6
            End If
        Else
            If pairname.Contains("cap") Then
                offset = -5
            Else
                If fintype = "E" Then
                    offset = -5
                Else
                    offset = -10
                End If
            End If
        End If

        Return offset
    End Function

    Shared Sub PlaceSingleVentsCU(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim partname, sutztenoccname, ventoccname As String
        Dim offsets(), offset As Double
        Dim normals(), ismirrored As Boolean
        Dim partno(header.VentIDs.Count - 1) As Integer
        Dim i, j As Integer

        Try
            i = 2
            j = 0

            'add muffe and sealing, don't add stopfen
            Do
                partno(j) = AddOccurance(asmdoc, header.VentIDs(i), General.currentjob.Workspace, "par", consys.Occlist)
                j += 1
                i -= 1
            Loop Until i < 1

            RemoveFromBOM(partno(1), asmdoc.Occurrences)

            'only CU checked so far
            Select Case header.Ventsize
                Case "G3/8"
                    offset = 2
                Case "G1/2"
                    offset = 1
                Case Else
                    offset = 20
            End Select

            If header.Tube.Materialcodeletter = "C" Then
                'assemble muffe and sealing
                AssembleMuffeSealing1(asmdoc, partno(0), partno(1), header.Ventsize, offset, "protrusion", 1, "C")
            Else
                'assemble muffe and sealing
                If header.Ventsize = "G1/4" Or header.Ventsize = "G1" Then
                    AssembleMuffeSealing2(asmdoc, partno(0), partno(1), 2, 1, 0, "C")
                Else
                    AssembleMuffeSealing2(asmdoc, partno(0), partno(1), 3, 1, 0, "C")
                End If
            End If

            If circuit.ConnectionSide = "left" Then
                ismirrored = True
            Else
                ismirrored = False
            End If

            partname = General.GetFullFilename(General.currentjob.Workspace, header.StutzenDatalist.First.ID, "par")
            partname = General.GetShortName(partname)

            sutztenoccname = partname + ":1"

            partname = General.GetFullFilename(General.currentjob.Workspace, header.VentIDs(2), "par")
            partname = General.GetShortName(partname)

            If header.Tube.HeaderType = "inlet" Then
                offsets = {0, 30.6, 50}
                normals = {True, ismirrored, ismirrored}
                ventoccname = partname + ":1"
            Else
                ventoccname = partname + ":2"
                offsets = {0, -30.6, 50}
                normals = {True, Not ismirrored, Not ismirrored}
            End If

            If circuit.FinType = "N" Or circuit.FinType = "M" Then
                offsets(2) -= 12
                If header.Tube.HeaderType = "outlet" And Not ismirrored Then
                    offsets(1) = -offsets(1)
                    normals(1) = Not normals(1)
                    normals(2) = Not normals(2)
                End If
            End If

            If circuit.ConnectionSide = "left" Then
                offsets(1) = -offsets(1)
            End If

            SetPlanar(asmdoc, asmdoc.Occurrences.Item(sutztenoccname), asmdoc.Occurrences.Item(ventoccname), 1, 2, normals(0), offsets(0), True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(sutztenoccname), asmdoc.Occurrences.Item(ventoccname), 2, 1, normals(1), offsets(1), True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(sutztenoccname), asmdoc.Occurrences.Item(ventoccname), 3, 3, normals(2), offsets(2), True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AssembleVent5(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, connectionside As String, fintype As String)
        Dim ventno As Integer
        Dim offset As Double
        Dim partname, stutzenoccname As String
        Dim naligned As Boolean

        Try
            ventno = AddOccurance(asmdoc, header.VentIDs.First, General.currentjob.Workspace, "par", consys.Occlist)

            If connectionside = "left" Then
                offset = -25
                naligned = True
            Else
                offset = 25
                naligned = False
            End If
            If header.Tube.HeaderType = "outlet" And (connectionside = "left" Or fintype = "F") Then
                offset = -offset
                naligned = Not naligned
            End If

            partname = General.GetFullFilename(General.currentjob.Workspace, header.StutzenDatalist.First.ID, "par")
            partname = General.GetShortName(partname)

            stutzenoccname = partname + ":1"

            SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenoccname), asmdoc.Occurrences.Item(ventno), 1, 2, False, 0, True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenoccname), asmdoc.Occurrences.Item(ventno), 2, 3, naligned, offset, True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenoccname), asmdoc.Occurrences.Item(ventno), 3, 1, naligned, 38, True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleVentsVA(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim adjocc, fixedocc As SolidEdgeAssembly.Occurrence
        Dim adjdoc As SolidEdgePart.PartDocument
        Dim adjface, fixedface As SolidEdgeGeometry.Face
        Dim adjmodel As SolidEdgePart.Model
        Dim partname, occname As String
        Dim offsets() As Double
        Dim partno(header.VentIDs.Count - 1) As Integer
        Dim i, j, stutzenno As Integer
        Dim normals() As Boolean
        Dim ismirrored As Boolean = False
        'even for CU tubes, an adapter pipe should always be necessary, smallest muffe is Ø14.5mm, stutzen is Ø15mm
        'vent parts always VA

        Try
            If circuit.ConnectionSide = "left" Then
                ismirrored = True
            End If
            i = 2
            j = 0

            'add muffe and sealing, don't add stopfen
            Do
                partno(j) = AddOccurance(asmdoc, header.VentIDs(i), General.currentjob.Workspace, "par", consys.Occlist)
                j += 1
                i -= 1
            Loop Until i < 1

            RemoveFromBOM(partno(1), asmdoc.Occurrences)

            'assemble muffe and sealing
            AssembleMuffeSealing2(asmdoc, partno(0), partno(1), 3, 1, 0, "C")

            'assemble muffe and stopfen
            'AssembleVent1(partno(0), partno(2), "G1/8", -4, "cutout", 3, "V")

            If consys.InletHeaders.First.StutzenDatalist.First.ID = consys.OutletHeaders.First.StutzenDatalist.First.ID And header.Tube.HeaderType = "outlet" Then
                occname = header.StutzenDatalist.First.ID + "0001-.par:2"
            Else
                occname = header.StutzenDatalist.First.ID + "0001-.par:1"
            End If

            'add pipe

            partname = General.GetFullFilename(General.currentjob.Workspace, header.VentIDs(3), "par")
            partno(3) = AddOccurance(asmdoc, header.VentIDs(3), General.currentjob.Workspace, "par", consys.Occlist)

            'assemble pipe and muffe 
            SetReversePlanar(asmdoc, partno(0), partno(3), offset:=-20, False)

            stutzenno = asmdoc.Occurrences.Item(occname).Index

            'get the faces for axial relation
            fixedocc = asmdoc.Occurrences.Item(partno(0))
            fixedface = SEPart.GetSideFace(fixedocc.PartDocument, "cutout", 1, 1)

            adjocc = asmdoc.Occurrences.Item(partno(3))
            adjdoc = adjocc.PartDocument
            adjmodel = adjdoc.Models.Item(1)
            adjface = SEPart.GetFaceFromLoop(adjmodel, "pipe", "axial", "5", "normal")

            SetAxialFaces(asmdoc, fixedocc, adjocc, fixedface, adjface, True, False)

            If header.Tube.HeaderType = "inlet" Or ((circuit.FinType = "N" Or circuit.FinType = "M") And circuit.ConnectionSide = "right") Then    'Or partname.Contains("928192") 
                offsets = {0, 3.6, 38}
                normals = {ismirrored, ismirrored, True}
            Else
                offsets = {0, -3.6, 38}
                Dim swmirror As Boolean = Not ismirrored
                normals = {swmirror, swmirror, True}
            End If

            If ismirrored Then
                offsets(1) = -offsets(1)
            End If

            'assemble pipe and stutzen, using 3 planar relations - all ref planes
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(occname), asmdoc.Occurrences.Item(partno(3)), 1, 1, normals(0), offsets(0), True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(occname), asmdoc.Occurrences.Item(partno(3)), 2, 3, normals(1), offsets(1), True)
            SetPlanar(asmdoc, asmdoc.Occurrences.Item(occname), asmdoc.Occurrences.Item(partno(3)), 3, 2, normals(2), offsets(2), True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleBrineStutzen(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData, circuit As CircuitData)
        Dim partno, tubeno, faceno As Integer
        Dim origin(), offset As Double
        Dim ismirrored As Boolean = False

        Try
            If circuit.ConnectionSide = "right" Then
                ismirrored = True
            End If

            partno = AddOccurance(asmdoc, header.StutzenDatalist.First.ID, General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)

            If partno > 0 Then
                origin = Calculation.GetPartOrigin(header.Xlist.First, header.Ylist.First, "back", circuit.CoreTubeOverhang, 0, 0)

                tubeno = GetTubeNoNew(origin, consys.CoreTubes.First, consys.HeaderAlignment)

                If tubeno > 0 Then
                    If consys.HeaderMaterial = "C" Then
                        offset = 0
                    Else
                        offset = -5
                    End If
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 3, 3, False, offset, True)
                    SetPlanar(asmdoc, asmdoc.Occurrences.Item(tubeno), asmdoc.Occurrences.Item(partno), 1, 1, ismirrored, 0, True)

                    faceno = GetFaceNo(asmdoc, partno, "F", "normal", True)
                    SetAxial(asmdoc, tubeno, partno, faceno)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleBrineParts(asmdoc As SolidEdgeAssembly.AssemblyDocument, ByRef header As HeaderData, ByRef consys As ConSysData)
        Dim madapter, capID, SVID As String
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim fixocc, svocc As SolidEdgeAssembly.Occurrence
        Dim stutzendoc As SolidEdgePart.PartDocument
        Dim stutzenmodel As SolidEdgePart.Model
        Dim fixedfaces(), adjfaces(), stutzenplface, stutzenaxface As SolidEdgeGeometry.Face
        Dim occlist As New List(Of SolidEdgeAssembly.Occurrence)
        Dim partorder As New List(Of String) From {"Stutzen"}
        Dim idorder As New List(Of String) From {header.StutzenDatalist.First.ID}
        Dim adjno, svno As Integer
        Dim offset As Double

        Try
            objoccs = asmdoc.Occurrences
            SVID = "107560"

            If header.Tube.HeaderType = "inlet" Then
                fixocc = objoccs.Item(header.StutzenDatalist.First.ID + "0001-.par:1")
            Else
                fixocc = objoccs.Item(header.StutzenDatalist.First.ID + "0001-.par:2")
            End If

            If consys.HeaderMaterial = "V" Then
                madapter = "0000970787"
                WSM.CheckoutPart(madapter, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                capID = "0000883642"
                WSM.CheckoutPart(capID, General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)
                idorder.AddRange({madapter, capID, SVID})
                partorder.AddRange({"nadapter", "cap", "SV"})
                occlist.Add(fixocc)

                For i As Integer = 1 To idorder.Count - 1
                    adjno = AddOccurance(asmdoc, idorder(i), General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)
                    fixocc = objoccs.Item(adjno)
                    occlist.Add(fixocc)
                Next

                For i As Integer = 1 To partorder.Count - 1
                    'faces of already fixed part
                    fixedfaces = SEPart.GetFaces(occlist(i - 1), partorder(i - 1), "fixed", 5, "reverse")

                    'faces of part to be fixed
                    adjfaces = SEPart.GetFaces(occlist(i), partorder(i), "adj", 3, "normal")

                    offset = GetPlOffset(partorder(i - 1) + "-" + partorder(i), "E")

                    SetAxialFaces(asmdoc, occlist(i - 1), occlist(i), fixedfaces(1), adjfaces(1), False, True)

                    SetPlanarFaces(asmdoc, occlist(i - 1), occlist(i), fixedfaces(0), adjfaces(0), True, offset)
                Next
            Else
                svno = AddOccurance(asmdoc, "107560", General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)
                svocc = objoccs.Item(svno)
                stutzendoc = fixocc.PartDocument
                stutzenmodel = stutzendoc.Models.Item(1)
                stutzenplface = SEPart.GetFaceFromLoop(stutzenmodel, "Stutzen", "planar", 5, "reverse")
                adjfaces = SEPart.GetFaces(svocc, "SV", "", 5, "normal")

                SetPlanarFaces(asmdoc, fixocc, svocc, stutzenplface, adjfaces(0), True, -5)

                stutzenaxface = SEPart.GetSideFace(stutzendoc, "cutout", 1, 1)

                SetAxialFaces(asmdoc, fixocc, svocc, stutzenaxface, adjfaces(1), False, True)
            End If
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub PlaceSingleBrineVents(asmdoc As SolidEdgeAssembly.AssemblyDocument, header As HeaderData, ByRef consys As ConSysData)
        Dim objoccs As SolidEdgeAssembly.Occurrences
        Dim adjocc, fixedocc As SolidEdgeAssembly.Occurrence
        Dim adjdoc As SolidEdgePart.PartDocument
        Dim adjface, fixedface As SolidEdgeGeometry.Face
        Dim adjmodel As SolidEdgePart.Model
        Dim partname, occname As String
        Dim offsets() As Double
        Dim i, j, partno(header.VentIDs.Count - 2), stutzenno, secondplno As Integer
        Dim normals() As Boolean


        Try
            i = 2
            j = 0

            Do
                'first 3 ventids are the vent parts, potential 4th is the pipe
                partname = General.GetFullFilename(General.currentjob.Workspace, header.VentIDs(i), "par")
                partname = General.GetShortName(partname)

                partno(j) = AddOccurance(asmdoc, partname, General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)

                j += 1
                i -= 1
            Loop Until i < 1

            RemoveFromBOM(partno(1), asmdoc.Occurrences)

            'assemble muffe and sealing
            AssembleMuffeSealing2(asmdoc, partno(0), partno(1), secondplno, 1, 0, "C", secondplno)

            objoccs = asmdoc.Occurrences
            If header.Tube.HeaderType = "inlet" Then
                occname = header.StutzenDatalist.First.ID + "0001-.par:1"
            Else
                occname = header.StutzenDatalist.First.ID + "0001-.par:2"
            End If
            stutzenno = objoccs.Item(occname).Index

            'add pipe
            If header.Ventsize = "G1/8" Then
                partname = General.GetFullFilename(General.currentjob.Workspace, header.VentIDs(3), "par")
                partname = General.GetShortName(partname)
                partno(2) = AddOccurance(asmdoc, partname, General.currentjob.Workspace, "par", consys.Occlist, header.Tube.HeaderType)

                'assemble pipe and muffe 
                SetReversePlanar(asmdoc, partno(0), partno(2), offset:=-20, False)

                'get the faces for axial relation
                fixedocc = asmdoc.Occurrences.Item(partno(0))
                fixedface = SEPart.GetSideFace(fixedocc.PartDocument, "cutout", 1, 1)

                adjocc = asmdoc.Occurrences.Item(partno.Last)
                adjdoc = adjocc.PartDocument
                adjmodel = adjdoc.Models.Item(1)
                adjface = SEPart.GetFaceFromLoop(adjmodel, "pipe", "axial", "5", "normal")

                SetAxialFaces(asmdoc, fixedocc, adjocc, fixedface, adjface, True, False)

                offsets = {0, -50, 80}
                normals = {False, False, False}
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=1, secondplane:=1, normals:=normals(0), offset:=offsets(0), True)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=2, secondplane:=2, normals:=normals(1), offset:=offsets(1), True)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=3, secondplane:=3, normals:=normals(2), offset:=offsets(2), True)
            Else
                offsets = {0, -50, 116}
                normals = {True, False, False}
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=1, secondplane:=3, normals:=normals(0), offset:=offsets(0), True)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=2, secondplane:=2, normals:=normals(1), offset:=offsets(1), True)
                SetPlanar(asmdoc, asmdoc.Occurrences.Item(stutzenno), asmdoc.Occurrences.Item(partno.Last), firstplane:=3, secondplane:=1, normals:=normals(2), offset:=offsets(2), True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetPSMfromASM(asmname As String) As String
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim occ1range(), occ2range() As Double
        Dim filename As String = ""

        Try
            General.seapp.Documents.Open(asmname)
            General.seapp.DoIdle()

            asmdoc = General.seapp.ActiveDocument

            occ1range = GetRangefromOcc(asmdoc.Occurrences.Item(1))
            occ2range = GetRangefromOcc(asmdoc.Occurrences.Item(2))

            If occ1range(1) > occ2range(1) Then
                filename = asmdoc.Occurrences.Item(1).OccurrenceFileName
            Else
                filename = asmdoc.Occurrences.Item(2).OccurrenceFileName
            End If
            General.seapp.Documents.CloseDocument(asmname, SaveChanges:=False, DoIdle:=True)

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return filename
    End Function

    Shared Function GetRangefromOcc(occurance As SolidEdgeAssembly.Occurrence) As Double()
        Dim minrange(2), maxrange(2), range(2) As Double

        occurance.Range(minrange(0), minrange(1), minrange(2), maxrange(0), maxrange(1), maxrange(2))

        For i As Integer = 0 To 2
            range(i) = Math.Round(Math.Abs((Math.Abs(maxrange(i)) - Math.Abs(minrange(i)))) * 1000, 1)
        Next

        Return range
    End Function

    Shared Sub RemovePatternfromBOM(asmdoc As SolidEdgeAssembly.AssemblyDocument, bowid As String)
        Dim asmpats As SolidEdgeAssembly.AssemblyPatterns = asmdoc.AssemblyPatterns

        Try
            'only patterns are relevent, single occurences are done already
            For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmpats
                If asmpat.Name.Contains("_Backbow") Then
                    Dim occarray(0) As Object
                    asmpat.GetOccurrences(occarray)
                    For i As Integer = 0 To occarray.Count - 1
                        Dim patocc As SolidEdgeAssembly.Occurrence = CType(occarray(i), SolidEdgeAssembly.Occurrence)
                        If patocc.Name.Contains(bowid) Then
                            patocc.IncludeInBom = False
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function CreateBlankView(asmdoc As SolidEdgeAssembly.AssemblyDocument, reftube As String, Optional justupdate As Boolean = False) As String
        Dim activeselectset As SolidEdgeFramework.SelectSet = General.seapp.ActiveSelectSet
        Dim viewconfig As SolidEdgeAssembly.Configuration
        Dim tubename As String = General.GetShortName(reftube)
        Dim configname As String = "BlankC"
        Dim showcomponent1 As Boolean = True
        Dim showcomponent2 As Boolean = True

        Try
            'show only coretubes and fins
            'hide everything
            General.seapp.StartCommand(33072)
            General.seapp.DoIdle()
            Threading.Thread.Sleep(1000)

            If reftube.Contains("Support") Then
                configname = "BlankS"
                If asmdoc.Occurrences.Item(1).Name.Contains("Coretube") Then
                    showcomponent1 = False
                End If
                If asmdoc.Occurrences.Item(2).Name.Contains("Coretube") Then
                    showcomponent2 = False
                End If
            End If

            asmdoc.Occurrences.Item(1).Visible = showcomponent1
            asmdoc.Occurrences.Item(2).Visible = showcomponent2
            asmdoc.Occurrences.Item(tubename + ".par:1").Visible = True

            If reftube.Contains("Coil") Then
                asmdoc.Occurrences.Item(tubename + ".par:2").Visible = True
            End If

            For Each asmgr As SolidEdgeAssembly.AssemblyGroup In asmdoc.AssemblyGroups
                If asmgr.Name.Contains(tubename.Substring(6, 4)) Then
                    activeselectset.Add(asmgr)
                    General.seapp.StartCommand(33093)
                End If
            Next

            If justupdate Then
                asmdoc.Configurations.Item(configname).Update()
            Else
                viewconfig = asmdoc.Configurations.Add(configname)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            'show everything
            General.seapp.StartCommand(33071)
        End Try

        Return configname
    End Function

    Shared Sub CreateBowLevelViews(asmdoc As SolidEdgeAssembly.AssemblyDocument, frontbows As List(Of String), backbows As List(Of String), frontlevel As List(Of Double), backlevel As List(Of Double), occlist As List(Of PartData), refconfig As String)
        Dim uniquefront, uniqueback As List(Of String)

        Try
            uniquefront = General.GetUniqueStrings(frontbows)
            uniqueback = General.GetUniqueStrings(backbows)
            For i As Integer = 1 To Math.Max(frontlevel.Max, backlevel.Max)
                asmdoc.Configurations.Item(refconfig).Apply()
                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)
                Dim indexlist As New List(Of Integer)
                Dim doneID As New List(Of String)
                Dim plist As New List(Of SolidEdgeAssembly.AssemblyPattern)
                Dim activeselectset As SolidEdgeFramework.SelectSet = General.seapp.ActiveSelectSet

                If uniquefront.Count > 0 Then
                    For j As Integer = 0 To uniquefront.Count - 1
                        If frontlevel(frontbows.IndexOf(uniquefront(j))) = i Then
                            If doneID.IndexOf(uniquefront(j)) = -1 Then
                                'find all items in occlist with this ID
                                Dim templist = From p In occlist Where p.Occname.Contains(uniquefront(j))
                                Debug.Print("frontlevel - " + i.ToString + " // " + templist.Count.ToString + " Items")
                                For Each e In templist.ToList
                                    indexlist.Add(e.Occindex)
                                Next

                                doneID.Add(uniquefront(j))
                            End If
                        End If
                    Next
                End If

                If uniqueback.Count > 0 Then
                    For j As Integer = 0 To uniqueback.Count - 1
                        If backlevel(backbows.IndexOf(uniqueback(j))) = i Then
                            If doneID.IndexOf(backbows(j)) = -1 Then
                                Dim templist = From p In occlist Where p.Occname.Contains(uniqueback(j))

                                Debug.Print("backlevel - " + i.ToString + " // " + templist.Count.ToString + " Items")
                                For Each e In templist.ToList
                                    indexlist.Add(e.Occindex)
                                Next

                                doneID.Add(uniqueback(j))
                            End If
                        End If
                    Next
                End If

                'find all patterns using the doneIDs
                For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmdoc.AssemblyPatterns
                    If asmpat.Name.Contains("bows") Then
                        Dim occarray(0) As Object
                        asmpat.GetOccurrences(occarray)
                        Dim patocc As SolidEdgeAssembly.Occurrence = CType(occarray(0), SolidEdgeAssembly.Occurrence)
                        For Each d In doneID
                            If patocc.Name.Contains(d) Then
                                plist.Add(asmpat)
                            End If
                        Next
                    End If
                Next

                For k As Integer = 0 To indexlist.Count - 1
                    asmdoc.Occurrences.Item(indexlist(k)).Visible = True
                Next
                For Each p In plist
                    activeselectset.Add(p)
                    General.seapp.StartCommand(33093)
                Next

                If doneID.Count > 0 Then
                    Try
                        asmdoc.Configurations.Add("Level" + i.ToString + refconfig.Substring(5, 1))
                    Catch ex As Exception
                        asmdoc.Configurations.Item("Level" + i.ToString + refconfig.Substring(5, 1)).Update()
                    End Try
                    General.seapp.DoIdle()
                End If

                activeselectset.RemoveAll()
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        Finally
            'show everything
            General.seapp.StartCommand(33071)
            Try
                asmdoc.Configurations.Add("Complete")
            Catch ex As Exception
                asmdoc.Configurations.Item("Complete").Update()
            End Try
            General.seapp.DoIdle()
        End Try

    End Sub

    Shared Sub CreateConsysViews(asmdoc As SolidEdgeAssembly.AssemblyDocument, consys As ConSysData, circuit As CircuitData, refconfig As String)
        Dim activeselectset As SolidEdgeFramework.SelectSet = General.seapp.ActiveSelectSet

        Try
            'hide everything
            asmdoc.Configurations.Item(refconfig).Apply()
            'If circuit.CircuitType = "Defrost" Then
            'Else
            '    asmdoc.Configurations.Item("BlankC").Apply()
            'End If
            'General.seapp.StartCommand(33072)
            Threading.Thread.Sleep(1000)

            Dim tempin = From plist In consys.Occlist Where plist.Configref = "inlet"

            If tempin.Count > 0 Then
                For Each occ In tempin.ToList
                    asmdoc.Occurrences.Item(occ.Occindex).Visible = True
                Next

                For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmdoc.AssemblyPatterns
                    If asmpat.Name.Contains("inlet") Then
                        Threading.Thread.Sleep(100)
                        activeselectset.Add(asmpat)
                        Threading.Thread.Sleep(100)
                        General.seapp.StartCommand(33093)
                    End If
                Next

                If General.currentunit.ApplicationType = "Condenser" Then
                    If consys.InletHeaders.First.OddLocation = "back" Then
                        asmdoc.Occurrences.Item("ConsysFin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par:2").Visible = True
                    Else
                        asmdoc.Occurrences.Item("ConsysFin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par:1").Visible = True
                    End If
                End If

                asmdoc.Configurations.Add("Inlet")

                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)

                If General.currentunit.ApplicationType = "Condenser" Then
                    If consys.InletHeaders.First.Tube.TopCapID <> "" Then
                        For Each occ In tempin.ToList
                            If occ.Occname.Contains(consys.InletHeaders.First.Tube.TopCapID) Then
                                'check if it assembled with header tube
                                Dim partnerlist As List(Of String) = PartnerOcc(asmdoc.Occurrences.Item(occ.Occindex))
                                For Each p In partnerlist
                                    If p.Contains("InletHeader") Then
                                        asmdoc.Occurrences.Item(occ.Occindex).Visible = False
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                    If consys.InletHeaders.First.Tube.BottomCapID <> "" Then
                        For Each occ In tempin.ToList
                            If occ.Occname.Contains(consys.InletHeaders.First.Tube.BottomCapID) Then
                                'check if it assembled with header tube
                                Dim partnerlist As List(Of String) = PartnerOcc(asmdoc.Occurrences.Item(occ.Occindex))
                                For Each p In partnerlist
                                    If p.Contains("InletHeader") Then
                                        asmdoc.Occurrences.Item(occ.Occindex).Visible = False
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                asmdoc.Configurations.Add("SideInlet")
                General.seapp.DoIdle()

                Threading.Thread.Sleep(1000)

                activeselectset.RemoveAll()
            End If

            'hide everything
            'General.seapp.StartCommand(33072)
            asmdoc.Configurations.Item(refconfig).Apply()
            'If circuit.CircuitType = "Defrost" Then
            'Else
            '    asmdoc.Configurations.Item("BlankC").Apply()
            'End If
            Threading.Thread.Sleep(1000)

            Dim tempout = From plist In consys.Occlist Where plist.Configref = "outlet"

            If tempout.Count > 0 Then
                For Each occ In tempout.ToList
                    asmdoc.Occurrences.Item(occ.Occindex).Visible = True
                Next

                For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmdoc.AssemblyPatterns
                    If asmpat.Name.Contains("outlet") Then
                        Threading.Thread.Sleep(100)
                        activeselectset.Add(asmpat)
                        Threading.Thread.Sleep(100)
                        General.seapp.StartCommand(33093)
                    End If
                Next

                If General.currentunit.ApplicationType = "Condenser" Then
                    If consys.OutletHeaders.First.OddLocation = "back" Then
                        asmdoc.Occurrences.Item("ConsysFin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par:2").Visible = True
                    Else
                        asmdoc.Occurrences.Item("ConsysFin" + circuit.Coilnumber.ToString + circuit.CircuitNumber.ToString + ".par:1").Visible = True
                    End If
                End If

                asmdoc.Configurations.Add("Outlet")

                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)

                If General.currentunit.ApplicationType = "Condenser" Then
                    If consys.OutletHeaders.First.Tube.TopCapID <> "" Then
                        For Each occ In tempout.ToList
                            If occ.Occname.Contains(consys.OutletHeaders.First.Tube.TopCapID) Then
                                'check if it assembled with header tube
                                Dim partnerlist As List(Of String) = PartnerOcc(asmdoc.Occurrences.Item(occ.Occindex))
                                For Each p In partnerlist
                                    If p.Contains("OutletHeader") Then
                                        asmdoc.Occurrences.Item(occ.Occindex).Visible = False
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                    If consys.OutletHeaders.First.Tube.BottomCapID <> "" Then
                        For Each occ In tempout.ToList
                            If occ.Occname.Contains(consys.OutletHeaders.First.Tube.BottomCapID) Then
                                'check if it assembled with header tube
                                Dim partnerlist As List(Of String) = PartnerOcc(asmdoc.Occurrences.Item(occ.Occindex))
                                For Each p In partnerlist
                                    If p.Contains("OutletHeader") Then
                                        asmdoc.Occurrences.Item(occ.Occindex).Visible = False
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                asmdoc.Configurations.Add("SideOutlet")
                General.seapp.DoIdle()

                Threading.Thread.Sleep(1000)
                activeselectset.RemoveAll()
            End If

            If consys.HasHotgas And consys.HeaderAlignment = "horizontal" Then
                'hide everything
                General.seapp.StartCommand(33072)
                Threading.Thread.Sleep(1000)

                If consys.HotGasData.Headertype = "inlet" Then
                    asmdoc.Occurrences.Item("InletHeader1_1_1.par:1").Visible = True
                Else
                    asmdoc.Occurrences.Item("OutletHeader1_1_1.par:1").Visible = True
                End If
                asmdoc.Configurations.Add("HGHeader")
                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)
            End If

            'show everything
            General.seapp.StartCommand(33071)
            asmdoc.Configurations.Add("Complete")
            General.seapp.DoIdle()

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Sub ReplaceOcc(asmdoc As SolidEdgeAssembly.AssemblyDocument, ToReplace As SolidEdgeAssembly.Occurrence, NewPart As String)

        Try
            asmdoc.Occurrences.Item(ToReplace.Index).Replace(NewPart, True)
            General.seapp.DoIdle()
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Function SingleBowViews(ByRef coil As CoilData) As Boolean
        Dim asmdoc As SolidEdgeAssembly.AssemblyDocument
        Dim bowids As New List(Of String)
        Dim x0, y0, z0 As Double
        Dim success As Boolean = False

        Try
            asmdoc = General.seapp.Documents.Open(coil.CoilFile.Fullfilename)
            General.seapp.DoIdle()

            'hide everything
            General.seapp.StartCommand(33072)
            General.seapp.DoIdle()
            Threading.Thread.Sleep(1000)

            asmdoc.Occurrences.Item("CoilFin" + coil.Number.ToString + "1.par:1").Visible = True
            asmdoc.Occurrences.Item("CoilFin" + coil.Number.ToString + "1.par:2").Visible = True

            Try
                asmdoc.Configurations.Add("Blank")
            Catch ex As Exception
                asmdoc.Configurations.Item("Blank").Update()
            End Try

            'front
            For i As Integer = 0 To coil.Circuits.Count - 1
                Dim templist As List(Of String) = General.GetUniqueStrings(coil.Circuits(i).Frontbowids)
                bowids = General.GetUniqueStrings(templist, bowids)
            Next
            coil.Frontbowids.AddRange(bowids.ToArray)

            For j As Integer = 0 To bowids.Count - 1
                'check added occurences
                asmdoc.Configurations.Item("Blank").Apply()
                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)

                Dim indexlist As New List(Of Integer)
                Dim plist As New List(Of SolidEdgeAssembly.AssemblyPattern)
                Dim activeselectset As SolidEdgeFramework.SelectSet = General.seapp.ActiveSelectSet

                For i As Integer = 0 To coil.Occlist.Count - 1
                    If coil.Occlist(i).Occname.Contains(bowids(j)) Then
                        Dim additem As Boolean = True
                        If bowids(j).Contains(Library.TemplateParts.BOW1) Then
                            If Not coil.Occlist(i).Occname.Contains(bowids(j) + ".par") Then
                                additem = False
                            End If
                        End If
                        'check origin (y)
                        asmdoc.Occurrences.Item(coil.Occlist(i).Occindex).GetOrigin(x0, y0, z0)
                        If y0 < 0 And additem Then
                            indexlist.Add(coil.Occlist(i).Occindex)
                        End If
                    End If
                Next
                'check patterns
                For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmdoc.AssemblyPatterns
                    If asmpat.Name.ToLower.Contains("front") Then
                        Dim occarray(0) As Object
                        asmpat.GetOccurrences(occarray)
                        Dim patocc As SolidEdgeAssembly.Occurrence = CType(occarray(0), SolidEdgeAssembly.Occurrence)
                        If patocc.Name.Contains(bowids(j)) Then
                            Dim additem As Boolean = True
                            If bowids(j).Contains(Library.TemplateParts.BOW1) Then
                                If Not patocc.Name.Contains(bowids(j) + ".par") Then
                                    additem = False
                                End If
                            End If
                            If additem Then
                                plist.Add(asmpat)
                            End If
                        End If
                    End If
                Next

                For k As Integer = 0 To indexlist.Count - 1
                    General.seapp.DoIdle()
                    asmdoc.Occurrences.Item(indexlist(k)).Visible = True
                    Threading.Thread.Sleep(100)
                    General.seapp.DoIdle()
                Next
                For Each p In plist
                    activeselectset.Add(p)
                    General.seapp.DoIdle()
                    General.seapp.StartCommand(33093)
                    Threading.Thread.Sleep(1000)
                    General.seapp.DoIdle()
                Next

                Try
                    asmdoc.Configurations.Add(bowids(j) + "_f")
                    General.seapp.DoIdle()
                Catch ex As Exception
                    asmdoc.Configurations.Item(bowids(j) + "_f").Update()
                    General.seapp.DoIdle()
                End Try

                activeselectset.RemoveAll()
                General.seapp.DoIdle()
            Next

            bowids.Clear()

            'back
            For i As Integer = 0 To coil.Circuits.Count - 1
                Dim templist As List(Of String) = General.GetUniqueStrings(coil.Circuits(i).Backbowids)
                bowids = General.GetUniqueStrings(templist, bowids)
            Next
            coil.Backbowids.AddRange(bowids.ToArray)

            For j As Integer = 0 To bowids.Count - 1
                'check added occurences
                asmdoc.Configurations.Item("Blank").Apply()
                General.seapp.DoIdle()
                Threading.Thread.Sleep(1000)

                Dim indexlist As New List(Of Integer)
                Dim plist As New List(Of SolidEdgeAssembly.AssemblyPattern)
                Dim activeselectset As SolidEdgeFramework.SelectSet = General.seapp.ActiveSelectSet

                For i As Integer = 0 To coil.Occlist.Count - 1
                    If coil.Occlist(i).Occname.Contains(bowids(j)) Then
                        Dim additem As Boolean = True
                        Dim ishp As Boolean = False
                        If bowids(j).Contains(Library.TemplateParts.BOW1) Then
                            If Not coil.Occlist(i).Occname.Contains(bowids(j) + ".par") Then
                                additem = False
                            End If
                        End If
                        'check origin (y)
                        asmdoc.Occurrences.Item(coil.Occlist(i).Occindex).GetOrigin(x0, y0, z0)
                        'check if hairpin
                        For k As Integer = 0 To coil.Circuits.Count - 1
                            For Each hp In coil.Circuits(k).Hairpins
                                If hp.PDMID.Contains(bowids(j)) Then
                                    ishp = True
                                End If
                            Next
                        Next
                        If (y0 > 0 Or ishp) And additem Then
                            indexlist.Add(coil.Occlist(i).Occindex)
                        End If
                    End If
                Next
                'check patterns
                For Each asmpat As SolidEdgeAssembly.AssemblyPattern In asmdoc.AssemblyPatterns
                    If asmpat.Name.ToLower.Contains("back") Then
                        Dim occarray(0) As Object
                        asmpat.GetOccurrences(occarray)
                        Dim patocc As SolidEdgeAssembly.Occurrence = CType(occarray(0), SolidEdgeAssembly.Occurrence)
                        If patocc.Name.Contains(bowids(j)) Then
                            Dim additem As Boolean = True
                            If bowids(j).Contains(Library.TemplateParts.BOW1) Then
                                If Not patocc.Name.Contains(bowids(j) + ".par") Then
                                    additem = False
                                End If
                            End If
                            If additem Then
                                plist.Add(asmpat)
                            End If
                        End If
                    End If
                Next

                For k As Integer = 0 To indexlist.Count - 1
                    General.seapp.DoIdle()
                    asmdoc.Occurrences.Item(indexlist(k)).Visible = True
                    General.seapp.DoIdle()
                Next
                For Each p In plist
                    activeselectset.Add(p)
                    General.seapp.DoIdle()
                    General.seapp.StartCommand(33093)
                    General.seapp.DoIdle()
                Next

                Try
                    asmdoc.Configurations.Add(bowids(j) + "_b")
                Catch ex As Exception
                    asmdoc.Configurations.Item(bowids(j) + "_b").Update()
                End Try

                activeselectset.RemoveAll()
            Next

            'show everything
            General.seapp.StartCommand(33071)
            General.seapp.DoIdle()

            General.seapp.Documents.CloseDocument(asmdoc.FullName, SaveChanges:=True, DoIdle:=True)

            success = True
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return success
    End Function

    Shared Function PartnerOcc(mainocc As SolidEdgeAssembly.Occurrence) As List(Of String)
        Dim objrels As SolidEdgeAssembly.Relations3d = mainocc.Relations3d
        Dim relname As String
        Dim partnerlist As New List(Of String)

        Try
            For i As Integer = 1 To objrels.Count
                Dim obj As Object = objrels.Item(i)
                relname = TypeName(obj)
                If relname.Contains("Axial") Then
                    Dim axrel As SolidEdgeAssembly.AxialRelation3d = TryCast(obj, SolidEdgeAssembly.AxialRelation3d)
                    If axrel.Occurrence1.Name = mainocc.Name Then
                        partnerlist.Add(axrel.Occurrence2.Name)
                    Else
                        partnerlist.Add(axrel.Occurrence1.Name)
                    End If
                ElseIf relname.Contains("Planar") Then
                    Dim plrel As SolidEdgeAssembly.PlanarRelation3d = DirectCast(obj, SolidEdgeAssembly.PlanarRelation3d)
                    If plrel.Occurrence1.Name = mainocc.Name Then
                        partnerlist.Add(plrel.Occurrence2.Name)
                    Else
                        partnerlist.Add(plrel.Occurrence1.Name)
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return partnerlist
    End Function

    Shared Sub PostSaveConsys(ByRef consys As ConSysData, ByRef circuit As CircuitData, ByRef coil As CoilData)
        Dim stindex As Integer
        Dim newstname As String = ""
        Try
            Dim asmdoc As SolidEdgeAssembly.AssemblyDocument = Nothing
            Do
                Try
                    asmdoc = General.seapp.ActiveDocument
                Catch ex As Exception

                End Try
            Loop Until asmdoc IsNot Nothing
            Dim newname As String = asmdoc.FullName

            consys.ConSysFile.Fullfilename = newname

            'replace coretube and fin par name in the coil.asm
            Dim ctocc As SolidEdgeAssembly.Occurrence = asmdoc.Occurrences.Item(2)
            Dim filename As String = ctocc.Name
            circuit.CoreTube.FileName = filename.Replace(":1", "")
            circuit.CoreTube.TubeFile.Fullfilename = ctocc.PartFileName
            circuit.CoreTube.TubeFile.Shortname = General.GetShortName(circuit.CoreTube.TubeFile.Fullfilename)

            circuit.FinName = asmdoc.Occurrences.Item(1).Name.Replace(":1", "")

            If circuit.CircuitType = "Defrost" Then
                'find support tube
                stindex = General.FindOccInList(consys.Occlist, "Supporttube")
                If stindex > -1 Then
                    newstname = asmdoc.Occurrences.Item(stindex).Name.Replace(":1", "")
                End If
            End If

            'open coilfile 
            General.seapp.Documents.CloseDocument(asmdoc.FullName, False, DoIdle:=True)

            General.seapp.Documents.Open(coil.CoilFile.Fullfilename)
            General.seapp.DoIdle()

            asmdoc = General.seapp.ActiveDocument

            If asmdoc.Occurrences.Item(1).Name.Contains("Fin+" + coil.Number.ToString + ".par:1") Then
                ReplaceOcc(asmdoc, asmdoc.Occurrences.Item(1), General.currentjob.Workspace + "\" + circuit.FinName)
            End If

            If asmdoc.Occurrences.Item(3).Name.Contains("Coretube" + coil.Number.ToString + ".par:1") Then
                ReplaceOcc(asmdoc, asmdoc.Occurrences.Item(3), General.currentjob.Workspace + "\" + circuit.CoreTube.TubeFile.Fullfilename)
            End If

            If circuit.CircuitType = "Defrost" AndAlso newstname <> "" Then
                'find occ in coil asm
                Dim coilstindex As Integer = General.FindOccInList(coil.Occlist, "Supporttube")
                If coilstindex > -1 Then
                    ReplaceOcc(asmdoc, asmdoc.Occurrences.Item(coilstindex), General.currentjob.Workspace + "\" + newstname)
                End If
            End If

            General.seapp.Documents.CloseDocument(asmdoc.FullName, True, DoIdle:=True)

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub WriteCustomProps(asmdoc As SolidEdgeAssembly.AssemblyDocument, f As FileData)
        SEPart.GetSetCustomProp(asmdoc, "CSG", "1", "write")
        SEPart.GetSetCustomProp(asmdoc, "Auftragsnummer", f.Orderno, "write")
        SEPart.GetSetCustomProp(asmdoc, "Position", f.Orderpos, "write")
        SEPart.GetSetCustomProp(asmdoc, "Order_Projekt", f.Projectno, "write")
        SEPart.GetSetCustomProp(asmdoc, "AGP_Nummer", f.AGPno + ".", "write")
        SEPart.GetSetCustomProp(asmdoc, "CDB_Benennung_de", f.CDB_de, "write")
        SEPart.GetSetCustomProp(asmdoc, "CDB_Benennung_en", f.CDB_en, "write")
        SEPart.GetSetCustomProp(asmdoc, "CDB_Zusatzbenennung", f.CDB_Zusatzbenennung, "write")
        SEPart.GetSetCustomProp(asmdoc, "CDB_z_bemerkung", f.CDB_z_Bemerkung, "write")
        SEPart.GetSetCustomProp(asmdoc, "GUE_Item", f.LNCode, "write")
        SEPart.GetSetCustomProp(asmdoc, "AGP_Plant", f.Plant, "write")
    End Sub

    Shared Function CheckOccExists(objoccs As SolidEdgeAssembly.Occurrences, occname As String) As SolidEdgeAssembly.Occurrence
        Dim occ As SolidEdgeAssembly.Occurrence = Nothing
        Try
            occ = objoccs.Item(occname)
        Catch
            Debug.Print("Didn't find occ " + occname)
        End Try

        Return occ
    End Function

End Class
