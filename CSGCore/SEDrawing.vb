Public Class SEDrawing

    Shared Function CreateDrawingCoil(coil As CoilData) As Boolean
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        Dim success As Boolean = False
        Try
            'checkout template drawing
            WSM.CheckoutCircs("767687", General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            If General.WaitForFile(General.currentjob.Workspace, "767687", ".dft", 100) Then

                'copy the drawing
                IO.File.Copy(General.GetFullFilename(General.currentjob.Workspace, "767687", ".dft"), coil.CoilFile.Fullfilename.Replace("asm", "dft"))
                Threading.Thread.Sleep(1000)

                General.seapp.Documents.Open(coil.CoilFile.Fullfilename.Replace("asm", "dft"))
                Threading.Thread.Sleep(1000)

                General.seapp.DoIdle()
                dftdoc = General.seapp.ActiveDocument

                If General.currentunit.ModelRangeName = "GACV" Then
                    GACVDrawings.MainCoil(dftdoc, coil)
                Else
                    CondenserDrawings.MainCoil(dftdoc, coil)
                End If

                'write properties
                WriteCostumProps(dftdoc, coil.CoilFile)
                SEPart.GetSetCustomProp(dftdoc, "GUE_Block", "1", "write")
                SEPart.GetSetCustomProp(dftdoc, "Z_Kategorie", "2D-Baugruppenzeichnung", "write")
                General.seapp.DoIdle()
                SEDraft.FitWindow()
                General.seapp.DoIdle()
                General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)
                success = True
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return success
    End Function

    Shared Sub CreateDrawingConSys(consys As ConSysData, coil As CoilData, circuit As CircuitData)
        Dim dftdoc As SolidEdgeDraft.DraftDocument

        Try
            'checkout template drawing
            WSM.CheckoutCircs("767687", General.currentjob.OrderDir, General.currentjob.Workspace, General.batfile)

            If General.WaitForFile(General.currentjob.Workspace, "767687", ".dft", 100) Then

                'copy the drawing
                IO.File.Copy(General.GetFullFilename(General.currentjob.Workspace, "767687", ".dft"), consys.ConSysFile.Fullfilename.Replace("asm", "dft"))
                Threading.Thread.Sleep(1000)

                General.seapp.Documents.Open(consys.ConSysFile.Fullfilename.Replace("asm", "dft"))
                Threading.Thread.Sleep(1000)

                General.seapp.DoIdle()
                dftdoc = General.seapp.ActiveDocument

                If General.currentunit.ModelRangeName = "GACV" Then
                    If circuit.CircuitType = "Defrost" Then
                        GACVDrawings.MainConsysDefrost(consys, coil, circuit, dftdoc)
                    Else
                        GACVDrawings.MainConsys(consys, coil, circuit, dftdoc)
                    End If
                ElseIf General.currentunit.ApplicationType = "Condenser" Then
                    CondenserDrawings.MainConsys(dftdoc, coil, circuit, consys)
                End If

                'write properties
                WriteCostumProps(dftdoc, consys.ConSysFile)
                SEPart.GetSetCustomProp(dftdoc, "GUE_Block", "0", "write")
                SEPart.GetSetCustomProp(dftdoc, "Z_Kategorie", "2D-Baugruppenzeichnung", "write")
                SEDraft.FitWindow()
                General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub ChangeModellink(dftdoc As SolidEdgeDraft.DraftDocument, workspace As String, newdoc As String)
        Dim mlinks As SolidEdgeDraft.ModelLinks = dftdoc.ModelLinks
        Dim mlink As SolidEdgeDraft.ModelLink = mlinks.Item(1)
        Dim objDVs As SolidEdgeDraft.DrawingViews = dftdoc.ActiveSheet.DrawingViews
        Dim objPLs As SolidEdgeDraft.PartsLists = dftdoc.PartsLists


        Try
            mlink.ChangeSource(workspace + "\" + newdoc)

            For Each DV As SolidEdgeDraft.DrawingView In objDVs
                DV.Update()
            Next

            For Each PL As SolidEdgeDraft.PartsList In objPLs
                PL.Update()
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function ChangeCaptionForBowViews(ByRef coil As CoilData, asmdoc As SolidEdgeAssembly.AssemblyDocument, coildft As String) As Boolean
        Dim occindex As Integer
        Dim bowoccnos As New List(Of Integer)
        Dim newbowocc As SolidEdgeAssembly.Occurrence
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        Dim bowdicc As New Dictionary(Of String, String)
        Dim success As Boolean
        Dim reflist As New List(Of String)

        Try

            For i As Integer = 0 To coil.Frontbowids.Count - 1
                If coil.Frontbowids(i).Contains(Library.TemplateParts.BOW1) OrElse coil.Frontbowids(i).Contains(Library.TemplateParts.BOW9) Then
                    'get occ index from occlist of coil
                    occindex = General.FindOccInList(coil.Occlist, coil.Frontbowids(i))
                    newbowocc = asmdoc.Occurrences.Item(occindex)
                    If bowoccnos.IndexOf(occindex) = -1 Then
                        bowdicc.Add(coil.Frontbowids(i), newbowocc.Name.Substring(0, 10))
                        bowoccnos.Add(occindex)
                        If coil.Occlist(occindex).Configref.Substring(0, 1) = "c" Then
                            reflist.Add("cooling")
                        Else
                            reflist.Add("brine")
                        End If
                    End If
                    coil.Frontbowids(i) = newbowocc.Name.Substring(0, 10)
                End If
            Next

            For i As Integer = 0 To coil.Backbowids.Count - 1
                If coil.Backbowids(i).Contains(Library.TemplateParts.BOW1) OrElse coil.Backbowids(i).Contains(Library.TemplateParts.BOW9) Then
                    'get occ index from occlist of coil
                    occindex = General.FindOccInList(coil.Occlist, coil.Backbowids(i))
                    newbowocc = asmdoc.Occurrences.Item(occindex)
                    If bowoccnos.IndexOf(occindex) = -1 Then
                        bowdicc.Add(coil.Backbowids(i), newbowocc.Name.Substring(0, 10))
                        bowoccnos.Add(occindex)
                        If coil.Occlist(occindex).Configref.Substring(0, 1) = "c" Then
                            reflist.Add("cooling")
                        Else
                            reflist.Add("brine")
                        End If
                    End If
                    coil.Backbowids(i) = newbowocc.Name.Substring(0, 10)
                End If
            Next

            dftdoc = General.seapp.Documents.Open(coildft)
            General.seapp.DoIdle()

            For i As Integer = 0 To bowdicc.Count - 1
                Dim newfrontcaption As String = "Front - " + bowdicc.Values(i)
                Dim newbackcaption As String = "Back - " + bowdicc.Values(i)
                If reflist(i) = "brine" Then
                    newfrontcaption = "Brine Front - " + bowdicc.Values(i)
                    newbackcaption = "Brine Back - " + bowdicc.Values(i)
                End If

                ChangeCaptions(dftdoc, "Coil" + coil.Number.ToString, "Front - " + bowdicc.Keys(i), newfrontcaption)
                ChangeCaptions(dftdoc, "Coil" + coil.Number.ToString, "Back - " + bowdicc.Keys(i), newbackcaption)
                ChangeCaptions(dftdoc, "Front", bowdicc.Keys(i), newfrontcaption)
                ChangeCaptions(dftdoc, "Back", bowdicc.Keys(i), newbackcaption)
            Next

            General.seapp.Documents.CloseDocument(dftdoc.FullName, SaveChanges:=True, DoIdle:=True)
            success = True
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
            success = False
        End Try

        Return success
    End Function

    Shared Sub ChangeCaptions(dftdoc As SolidEdgeDraft.DraftDocument, sheetname As String, oldpartialname As String, newname As String)

        Try
            For Each objSheet As SolidEdgeDraft.Sheet In dftdoc.Sheets
                If objSheet.Name = sheetname Then
                    For Each objDV As SolidEdgeDraft.DrawingView In objSheet.DrawingViews
                        If objDV.CaptionDefinitionTextPrimary.Contains(oldpartialname) Then
                            objDV.CaptionDefinitionTextPrimary = newname
                            If General.currentunit.ModelRangeName = "GACV" And newname.Length > 15 Then
                                objDV.CaptionDefinitionTextPrimary = objDV.CaptionDefinitionTextPrimary.Replace(" - ", vbNewLine)
                            End If
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub WriteCostumProps(dftdoc As SolidEdgeDraft.DraftDocument, f As FileData)
        SEPart.GetSetCustomProp(dftdoc, "CSG", "1", "write")
        SEPart.GetSetCustomProp(dftdoc, "Auftragsnummer", f.Orderno, "write")
        SEPart.GetSetCustomProp(dftdoc, "Position", f.Orderpos, "write")
        SEPart.GetSetCustomProp(dftdoc, "Order_Projekt", f.Projectno, "write")
        SEPart.GetSetCustomProp(dftdoc, "AGP_Nummer", f.AGPno + ".", "write")
        SEPart.GetSetCustomProp(dftdoc, "CDB_Benennung_de", f.CDB_de, "write")
        SEPart.GetSetCustomProp(dftdoc, "CDB_Benennung_en", f.CDB_en, "write")
        SEPart.GetSetCustomProp(dftdoc, "CDB_Zusatzbenennung", f.CDB_Zusatzbenennung, "write")
        SEPart.GetSetCustomProp(dftdoc, "CDB_z_bemerkung", f.CDB_z_Bemerkung, "write")
        SEPart.GetSetCustomProp(dftdoc, "GUE_Item", f.LNCode, "write")
        SEPart.GetSetCustomProp(dftdoc, "AGP_Plant", f.Plant, "write")
    End Sub

    Shared Function AddBowViewtoSheet(objsheet As SolidEdgeDraft.Sheet, asmlink As SolidEdgeDraft.ModelLink, scalefactor As Double, bowid As String, side As String, coil As CoilData) As SolidEdgeDraft.DrawingView
        Dim objDV As SolidEdgeDraft.DrawingView = Nothing
        Dim bowERP, hpERP, configref As String
        Dim igView As Integer


        Try
            If side = "Front" Then
                igView = 4
                hpERP = ""
            Else
                igView = 6
                hpERP = Database.GetValue("CSG.Hairpins", "ERPCode", "Article_Number", bowid)
                If hpERP = "" Or hpERP = "NULL" Then
                    'check if bowid is replacement of hairpin 
                    For Each circ In coil.Circuits
                        For Each hp In circ.Hairpins
                            If hp.RefBow = bowid Then
                                hpERP = hp.ERPCode
                            End If
                        Next
                    Next
                End If
            End If
            'viewconfig name = bowid_[s]
            objDV = objsheet.DrawingViews.AddAssemblyView(asmlink, igView, scalefactor, -1, 0, 0, bowid + "_" + side.Substring(0, 1).ToLower)

            'get ERPCode for bow view
            bowERP = Database.GetValue("CSG.DB_Bows", "ERPCode", "Article_Number", bowid.Substring(0, 10))
            configref = General.GetConfigname(coil.Occlist, bowid)
            If hpERP <> "" Then
                objDV.CaptionDefinitionTextPrimary = side + " - " + hpERP
            ElseIf bowERP <> "" Then
                If configref.Substring(0, 1) = "c" Then
                    objDV.CaptionDefinitionTextPrimary = side + " - " + bowERP
                Else
                    objDV.CaptionDefinitionTextPrimary = "Brine " + side + " - " + bowERP
                End If
            Else
                If configref.Substring(0, 1) = "c" Then
                    objDV.CaptionDefinitionTextPrimary = side + " - " + bowid
                Else
                    objDV.CaptionDefinitionTextPrimary = "Brine " + side + " - " + bowid
                End If
            End If
            objDV.CaptionLocation = SolidEdgeFrameworkSupport.DimViewCaptionLocationConstants.igDimViewCaptionLocationTop
            objDV.DisplayCaption = True

            'handle visibility
            'hide hidden lines
            For Each DVLine As SolidEdgeDraft.DVLine2d In objDV.DVLines2d
                DVLine.ModelMember.ShowEdgesHiddenByOtherParts = False
                DVLine.ModelMember.ShowHiddenEdges = False
            Next

            'hide hidden circles
            For Each DVCircle As SolidEdgeDraft.DVCircle2d In objDV.DVCircles2d
                DVCircle.ModelMember.ShowEdgesHiddenByOtherParts = False
                DVCircle.ModelMember.ShowHiddenEdges = False
            Next

            objDV.SetOrigin(0, 0)
        Catch ex As Exception

        End Try
        Return objDV
    End Function

    Shared Sub AddSheet(mainsheet As SolidEdgeDraft.Sheet)
        Dim dftdoc As SolidEdgeDraft.DraftDocument = mainsheet.Parent
        Dim bgsheet As SolidEdgeDraft.Sheet = mainsheet.Background
        Dim secondsheet As SolidEdgeDraft.Sheet

        Try
            For Each objsheet As SolidEdgeDraft.Sheet In dftdoc.Sheets
                If objsheet.Name = "Bows" Then
                    Exit Sub
                End If
            Next
            secondsheet = dftdoc.Sheets.AddSheet("Bows", 0)
            secondsheet.Background = bgsheet
            If bgsheet.Name.Contains("A2") Then
                secondsheet.SheetSetup.SheetSizeOption = 33
            End If
            mainsheet.Activate()
        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub AddADBlock(dftdoc As SolidEdgeDraft.DraftDocument, objDV As SolidEdgeDraft.DrawingView, position As String)
        Dim objBlock As SolidEdgeDraft.Block
        Dim objBOcc As SolidEdgeDraft.BlockOccurrence
        Dim x, y, x0, y0, mp, xmin, ymin, xmax, ymax As Double

        Try

            objDV.GetCaptionPosition(x, y)
            objDV.DisplayCaption = False
            objDV.GetOrigin(x0, y0)
            objDV.Range(xmin, ymin, xmax, ymax)

            If General.currentunit.ModelRangeName = "GACV" Then
                'position is left or right
                If position = "right" Then
                    objBlock = dftdoc.Blocks.Item(0)
                    mp = 1
                Else
                    objBlock = dftdoc.Blocks.Item(1)
                    mp = -1
                End If
                objBOcc = dftdoc.ActiveSheet.BlockOccurrences.Add(objBlock.Name, x0 + mp * (xmax - xmin) / 2, y0)
                If position = "right" Then
                    objBOcc.SetOrigin(xmax, y0)
                Else
                    objBOcc.SetOrigin(xmin, y0)
                End If
                If objDV.CaptionDefinitionTextPrimary.Length > 15 Then
                    objDV.CaptionDefinitionTextPrimary = objDV.CaptionDefinitionTextPrimary.Replace(" - ", vbNewLine)
                End If
            Else
                objBlock = dftdoc.Blocks.Item(1)
                'position is always at the bottom
                'objBOcc = dftdoc.ActiveSheet.BlockOccurrences.Add(objBlock.Name, x0 + (xmax - xmin) / 4, y0 - (ymax - ymin) / 2 + 0.5 / objDV.ScaleFactor * 0.001, Rotation:=Math.PI / 2)
                objBOcc = dftdoc.ActiveSheet.BlockOccurrences.Add(objBlock.Name, x0 + (xmax - xmin) / 4, ymin, Rotation:=Math.PI / 2)
            End If

            objBOcc.Block.Name = "Bow" + position
            objDV.DisplayCaption = True

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

    End Sub

    Shared Function GetLinesFromOcc(objDV As SolidEdgeDraft.DrawingView, partID As String, conside As String, Optional yref As Double = 100) As List(Of SolidEdgeDraft.DVLine2d)
        Dim objLines As New List(Of SolidEdgeDraft.DVLine2d)
        Dim objDVLines As SolidEdgeDraft.DVLines2d
        Dim xs, ys, xe, ye, xref As Double

        Try
            objDVLines = objDV.DVLines2d
            For Each objDVLine As SolidEdgeDraft.DVLine2d In objDVLines
                If objDVLine.ModelMember.FileName.Contains(partID) Then
                    If yref < 100 Then
                        objDVLine.GetStartPoint(xs, ys)
                        objDVLine.GetEndPoint(xe, ye)
                        'check if vertical
                        If Math.Abs(xe - xs) < 0.001 Then
                            'check the y condition
                            If Math.Round(Math.Max(ys, ye), 6) >= yref And Math.Round(Math.Min(ys, ye), 6) <= yref Then
                                xref = Math.Round(Math.Max(xe, xref), 6)
                                If (Math.Round(xe, 6) >= xref And conside = "left") Or (Math.Round(xe, 6) <= xref And conside = "right") Then
                                    objLines.Add(objDVLine)
                                End If
                            End If
                        End If
                    Else
                        objLines.Add(objDVLine)
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objLines
    End Function

    Shared Function GetCirclesFromOcc(objDV As SolidEdgeDraft.DrawingView, partIDList As List(Of String)) As List(Of SolidEdgeDraft.DVCircle2d)
        Dim objCircles As New List(Of SolidEdgeDraft.DVCircle2d)

        Try
            For Each objDVCircle As SolidEdgeDraft.DVCircle2d In objDV.DVCircles2d
                For Each pID In partIDList
                    If objDVCircle.ModelMember.FileName.Contains(pID) Then
                        objCircles.Add(objDVCircle)
                    End If
                Next
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objCircles
    End Function

    Shared Function GetArcsFromOcc(objDV As SolidEdgeDraft.DrawingView, partID As String, diameter As Double) As List(Of SolidEdgeDraft.DVArc2d)
        Dim objArcs As New List(Of SolidEdgeDraft.DVArc2d)

        Try
            For Each objDVArc As SolidEdgeDraft.DVArc2d In objDV.DVArcs2d
                If objDVArc.ModelMember.FileName.Contains(partID) Then
                    If diameter > 0 Then
                        Dim xs, xe, ys, ye As Double
                        objDVArc.GetStartPoint(xs, ys)
                        objDVArc.GetEndPoint(xe, ye)
                        If Math.Abs(Math.Round((ye - ys) * 1000, 3)) = diameter Then
                            objArcs.Add(objDVArc)
                        End If
                    Else
                        objArcs.Add(objDVArc)
                    End If
                End If
            Next

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return objArcs
    End Function

    Shared Function GetHorHeaderDimLines(headerlines As List(Of SolidEdgeDraft.DVLine2d), diameter As Double) As SolidEdgeDraft.DVLine2d()
        Dim xs, ys, xe, ye As Double
        Dim xlist As New List(Of Double)
        Dim plist As New List(Of SolidEdgeDraft.DVLine2d)
        Dim leftline, rightline As SolidEdgeDraft.DVLine2d

        Try
            For Each l In headerlines
                l.GetStartPoint(xs, ys)
                l.GetEndPoint(xe, ye)
                If Math.Round(xe, 6) = Math.Round(xs, 6) AndAlso Math.Round(l.Length * 1000, 3) = diameter Then
                    xlist.Add(Math.Round(xe, 6))
                    plist.Add(l)
                End If
            Next

            leftline = plist(xlist.IndexOf(xlist.Min))
            rightline = plist(xlist.IndexOf(xlist.Max))

        Catch ex As Exception

        End Try

#Disable Warning BC42104 ' Die rightline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
#Disable Warning BC42104 ' Die leftline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
        Return {leftline, rightline}
#Enable Warning BC42104 ' Die leftline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
#Enable Warning BC42104 ' Die rightline-Variable wird verwendet, bevor ihr ein Wert zugewiesen wird. Zur Laufzeit kann eine Nullverweisausnahme auftreten.
    End Function

    Shared Function GetHeaderFrame(headerdimlines() As SolidEdgeDraft.DVLine2d) As Double()
        Dim headerframe(3), xs, xe, ys, ye As Double

        headerdimlines(0).GetStartPoint(xs, ys)
        headerdimlines(0).GetEndPoint(xe, ye)

        headerframe(0) = Math.Round(xs, 6)
        headerframe(1) = Math.Round(Math.Min(ys, ye), 6)

        headerdimlines(1).GetStartPoint(xs, ys)
        headerdimlines(1).GetEndPoint(xe, ye)

        headerframe(2) = Math.Round(xs, 6)
        headerframe(3) = Math.Round(Math.Max(ys, ye), 6)

        Return headerframe
    End Function

    Shared Sub ControlDimDistance(objDim As SolidEdgeFrameworkSupport.Dimension, controloptions() As String, Optional newtrackdistance As Double = 0, Optional switchaxisdir As Boolean = False,
                                   Optional breakpos As Integer = 2)
        Dim xmin1, ymin1, xmax1, ymax1, xmin2, ymin2, xmax2, ymax2 As Double

        Try
            If newtrackdistance <> 0 Then
                objDim.TrackDistance = newtrackdistance
            End If
            If switchaxisdir Then
                objDim.MeasurementAxisDirection = Not objDim.MeasurementAxisDirection
            End If
            objDim.Range(xmin1, ymin1, xmax1, ymax1)
            If breakpos <> 2 Then
                objDim.BreakPosition = breakpos
            Else
                objDim.TrackDistance *= -1
            End If
            objDim.Range(xmin2, ymin2, xmax2, ymax2)

            If breakpos = 2 Then
                If controloptions(0) = "x" Then
                    If controloptions(1) = "bigger" Then
                        If xmin1 >= xmin2 Then
                            objDim.TrackDistance *= -1
                        End If
                    Else
                        If xmin1 <= xmin2 Then
                            objDim.TrackDistance *= -1
                        End If
                    End If
                Else
                    If controloptions(1) = "bigger" Then
                        If ymin1 >= ymin2 Then
                            objDim.TrackDistance *= -1
                        End If
                    Else
                        If ymin1 <= ymin2 Then
                            objDim.TrackDistance *= -1
                        End If
                    End If
                End If
            Else
                If controloptions(0) = "x" Then
                    If controloptions(1) = "bigger" Then
                        If xmin1 >= xmin2 Then
                            objDim.BreakPosition = 4 - breakpos
                        End If
                    Else
                        If xmin1 <= xmin2 Then
                            objDim.BreakPosition = 4 - breakpos
                        End If
                    End If
                Else
                    If controloptions(1) = "bigger" Then
                        If ymin1 >= ymin2 Then
                            objDim.BreakPosition = 4 - breakpos
                        End If
                    Else
                        If ymin1 <= ymin2 Then
                            objDim.BreakPosition = 4 - breakpos
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub CreateBoundary(objsheet As SolidEdgeDraft.Sheet, x As Double, y As Double, diameter As Double)
        Dim objboundaries As SolidEdgeFrameworkSupport.Boundaries2d = objsheet.Boundaries2d
        Dim newboundary As SolidEdgeFrameworkSupport.Boundary2d = Nothing
        Dim objcirclist As New List(Of SolidEdgeFrameworkSupport.Circle2d)

        Try

            Dim newcircle As SolidEdgeFrameworkSupport.Circle2d = objsheet.Circles2d.AddByCenterRadius(x, y, diameter / 2000)
            newcircle.Style.Width = 0.00025

            objcirclist.Add(newcircle)

            Try
                newboundary = objboundaries.AddByObjects(1, CType(objcirclist.ToArray, Array), x, y)
            Catch ex As Exception
                General.CreateLogEntry(ex.ToString)
            End Try

            If newboundary Is Nothing Then
                newboundary = objboundaries.Item(objboundaries.Count)
            End If

            newboundary.Style.FillName = "Normal"
            newboundary.Style.FillColor = 255
            newboundary.Style.LinearColor = 255

        Catch ex As Exception

        End Try
    End Sub

    Shared Sub UpdateDrawingAfterSave(filename As String)
        Dim dftdoc As SolidEdgeDraft.DraftDocument
        'open the dft and update the partlist, then save and close

        If IO.File.Exists(filename) Then
            dftdoc = SEDraft.OpenDFT(filename)

            Try
                For Each objpartlist As SolidEdgeDraft.PartsList In dftdoc.PartsLists
                    objpartlist.Update()
                Next

                SEDraft.FitWindow()

                For Each objsheet As SolidEdgeDraft.Sheet In dftdoc.Sheets
                    If objsheet.SectionType = SolidEdgeDraft.SheetSectionTypeConstants.igWorkingSection Then
                        objsheet.Activate()
                        General.seapp.DoIdle()
                        For Each objDV As SolidEdgeDraft.DrawingView In objsheet.DrawingViews
                            objDV.Update()
                            General.seapp.DoIdle()
                        Next
                    End If
                Next

            Catch ex As Exception

            Finally
                General.seapp.Documents.CloseDocument(filename, SaveChanges:=True, DoIdle:=True)
            End Try
        End If

    End Sub

End Class
