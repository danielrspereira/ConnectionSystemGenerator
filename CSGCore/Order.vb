Imports System.IO
Imports Newtonsoft.Json
Imports System.Net
Imports System.Text
Imports System.Xml

Public Class Order

    Shared Function GetEUPlant() As String
        Dim localtime As Date = My.Computer.Clock.LocalTime
        Dim mytimezone As TimeZone = TimeZone.CurrentTimeZone
        Dim deltatime As TimeSpan = TimeZone.CurrentTimeZone.GetUtcOffset(localtime)
        Dim deltahours As Integer
        Dim plantname As String = "Tata"

        Try
            If mytimezone.IsDaylightSavingTime(localtime) Then
                deltahours = deltatime.Hours - 1
            Else
                deltahours = deltatime.Hours
            End If
            If deltahours = 2 Then
                plantname = "Sibiu"
            End If
        Catch ex As Exception

        End Try

        Return plantname
    End Function

    Shared Sub CreateOrderDir(ByRef currentjob As JobInfo)
        currentjob.OrderDir = General.csgdir + "\" + currentjob.OrderNumber + "_" + currentjob.OrderPosition

        If Directory.Exists(currentjob.OrderDir) = False Then
            My.Computer.FileSystem.CreateDirectory(currentjob.OrderDir)
        Else
            For Each f As String In Directory.GetFiles(currentjob.OrderDir)
                Try
                    If f.Contains(".json") And f.Contains("cond_") Then
                        File.Delete(f)
                        Threading.Thread.Sleep(100)
                    End If
                Catch ex As Exception

                End Try
            Next
        End If
        General.errorlogfile = currentjob.OrderDir + "\" + currentjob.OrderNumber + "_" + currentjob.OrderPosition + "_error.txt"
        If File.Exists(General.errorlogfile) Then
            File.Delete(General.errorlogfile)
        End If
        General.actionlogfile = currentjob.OrderDir + "\" + currentjob.OrderNumber + "_" + currentjob.OrderPosition + "_action.txt"
        If File.Exists(General.actionlogfile) Then
            File.Delete(General.actionlogfile)
        End If
    End Sub

    Shared Function GetOrderkey(orderno As String, orderpos As String) As String
        Dim orderkey As String
        Dim lengthpos As Integer

        lengthpos = orderpos.Length
        Do While lengthpos < 4
            orderpos = "0" + orderpos
            lengthpos = orderpos.Length
        Loop

        orderkey = orderno + "_" + orderpos + "_00"

        Return orderkey
    End Function

    Shared Sub SearchXML(ByRef currentjob As JobInfo)
        Dim pcfname, fullfilename, fileending As String

        fileending = General.GetPCFFileEnding()
        pcfname = GetOrderkey(currentjob.OrderNumber, currentjob.OrderPosition) + fileending

        General.SetGPCDir(currentjob.OrderNumber)

        If File.Exists(currentjob.OrderDir + "\" + pcfname) Then
            File.Delete(currentjob.OrderDir + "\" + pcfname)
        End If

        If Directory.Exists(General.gpcdir) Then
            fullfilename = FindPCFFile(pcfname, Date.Now.Year)
            Debug.Print(fullfilename)
            If fullfilename.Contains(fileending) Then
                If File.Exists(fullfilename) Then
                    currentjob.Path = fullfilename.Substring(0, fullfilename.LastIndexOf("\"))
                    My.Computer.FileSystem.CopyFile(fullfilename, currentjob.OrderDir + "\" + pcfname)
                    General.localpcffile = currentjob.OrderDir + "\" + pcfname
                    General.WaitForFile(currentjob.OrderDir, pcfname, "pcf", 100)
                End If
            End If
        End If

    End Sub

    Shared Function FindPCFFile(pcfname As String, searchyear As Integer) As String
        Dim maindir As String = General.gpcdir + "\" + searchyear.ToString
        Dim currentweek As Integer
        Dim datetoday As Date = Date.Now
        Dim filefound As Boolean = False
        Dim loopexit As Boolean = False
        Dim searchdir, searchweek, fullfilename As String
        Dim i, loopcount As Integer

        fullfilename = ""

        If Directory.Exists(maindir) Then
            Debug.Print("Starting: " + Date.Now.ToLongTimeString)
            currentweek = DatePart(DateInterval.WeekOfYear, datetoday, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            i = currentweek
            loopcount = 0

            Do
                If i < 10 Then
                    searchweek = "0" + i.ToString
                Else
                    searchweek = i.ToString
                End If
                searchdir = maindir + "\" + searchweek

                If Directory.Exists(searchdir) Then
                    fullfilename = searchdir + "\" + pcfname
                    If File.Exists(fullfilename) Then
                        filefound = True
                    End If
                End If

                loopcount += +1
                i -= 1
                If loopcount > 53 Then
                    loopexit = True
                End If
                If i = 0 And loopcount < 53 And filefound = False Then
                    i = 53
                    maindir = maindir.Replace(searchyear.ToString, (searchyear - 1).ToString)
                End If
            Loop Until filefound Or loopexit
        Else
            fullfilename = "Directory not found"
        End If
        If filefound = False And searchyear > 2020 Then
            fullfilename = FindPCFFile(pcfname, searchyear - 1)
        End If
        Debug.Print("Finished: " + Date.Now.ToLongTimeString)
        Return fullfilename
    End Function

    Shared Sub ConstructDataModel()
        Dim firstcoil, secondcoil_mc, secondcoil, secondcoil_vs As CoilData
        Dim firstcircuit, secondcircuit_mc, secondcircuit, secondcircuit_vs, firstsubcircuit, firstsubcircuit_vs, secondsubcircuit_mc, secondsubcircuit_vs, brinecircuit As CircuitData
        Dim firstconsys, secondconsys_mc, secondconsys, secondconsys_vs, firstsubconsys, firstsubconsys_vs, secondsubconsys_mc, secondsubconsys_vs, brineconsys As ConSysData
        Dim circuitcounter As Integer = 1
        Dim coilcounter As Integer = 1

        General.coillist.Clear()
        General.circuitlist.Clear()
        General.consyslist.Clear()

        Try
            PCFData.GatherDatafromPCF(General.localpcffile)

            PCFData.FillUnitDataFromXML(General.currentunit)

            firstcoil = PCFData.FillCoilDataFromXML(1)
            firstcircuit = PCFData.FillCircuitDataFromXML("First", 1, firstcoil, circuitcounter)
            firstconsys = PCFData.FillConSysDataFromXML(firstcircuit, 1, "First")
            General.coillist.Add(firstcoil)
            General.circuitlist.Add(firstcircuit)
            General.consyslist.Add(firstconsys)

            If General.currentunit.HasIntegratedSubcooler Then
                circuitcounter += 1
                firstsubcircuit = PCFData.FillCircuitDataFromXML("Subcooler", 1, firstcoil, circuitcounter)
                General.circuitlist.Add(firstsubcircuit)
                firstsubconsys = PCFData.FillConSysDataFromXML(firstsubcircuit, 1, "Subcooler")
                General.consyslist.Add(firstsubconsys)
            End If

            If General.currentunit.MultiCircuitDesign = "1" Then
                circuitcounter += 1
                secondcircuit_mc = PCFData.FillCircuitDataFromXML("Second", 2, firstcoil, circuitcounter)
                General.circuitlist.Add(secondcircuit_mc)
                secondconsys_mc = PCFData.FillConSysDataFromXML(secondcircuit_mc, 2, "Second")
                General.consyslist.Add(secondconsys_mc)
                If General.currentunit.HasIntegratedSubcooler Then
                    circuitcounter += 1
                    secondsubcircuit_mc = PCFData.FillCircuitDataFromXML("Subcooler", 2, firstcoil, circuitcounter)
                    General.circuitlist.Add(secondsubcircuit_mc)
                    secondsubconsys_mc = PCFData.FillConSysDataFromXML(secondsubcircuit_mc, 2, "Subcooler")
                    General.consyslist.Add(secondsubconsys_mc)
                End If
            ElseIf General.currentunit.MultiCircuitDesign = "2" Then        'Sandwich Design
                circuitcounter += 1
                secondcircuit_mc = PCFData.FillCircuitDataFromXML("Sandwich", 2, firstcoil, circuitcounter)
                secondconsys_mc = PCFData.FillConSysDataFromXML(secondcircuit_mc, 2, "Main")
                General.circuitlist.Add(secondcircuit_mc)
                General.consyslist.Add(secondconsys_mc)
                If General.currentunit.HasIntegratedSubcooler Then
                    circuitcounter += 1
                    secondsubcircuit_mc = PCFData.FillCircuitDataFromXML("Subcooler", 2, firstcoil, circuitcounter)
                    General.circuitlist.Add(secondsubcircuit_mc)
                    secondsubconsys_mc = PCFData.FillConSysDataFromXML(secondsubcircuit_mc, 2, "Subcooler")
                    General.consyslist.Add(secondsubconsys_mc)
                End If
            End If

            If General.currentunit.UnitDescription = "VShape" Then
                coilcounter += 1
                circuitcounter += 1
                secondcoil = PCFData.FillCoilDataFromXML(1)
                secondcoil.Number = coilcounter        'number = 2 for anything but sandwich // number = 3 for sandwich
                General.coillist.Add(secondcoil)
                secondcircuit = PCFData.FillCircuitDataFromXML("First", 1, secondcoil, 1)
                secondcircuit.Coilnumber = coilcounter
                firstcircuit.ConnectionSide = "left"
                secondcircuit.ConnectionSide = "right"
                General.circuitlist.Add(secondcircuit)
                secondconsys = PCFData.FillConSysDataFromXML(secondcircuit, 1, "")
                General.consyslist.Add(secondconsys)

                If General.currentunit.HasIntegratedSubcooler Then
                    circuitcounter += 1
                    firstsubcircuit_vs = PCFData.FillCircuitDataFromXML("Subcooler", 1, secondcoil, circuitcounter)
                    firstsubcircuit_vs.Coilnumber = coilcounter
                    firstsubcircuit.ConnectionSide = "left"
                    firstsubcircuit_vs.ConnectionSide = "right"
                    General.circuitlist.Add(firstsubcircuit_vs)
                    firstsubconsys_vs = PCFData.FillConSysDataFromXML(firstsubcircuit_vs, 1, "Subcooler")
                    firstsubconsys_vs.Circnumber = circuitcounter
                    General.consyslist.Add(firstsubconsys_vs)
                End If
                If General.currentunit.MultiCircuitDesign = "2" Then
                    coilcounter += 1
                    secondcoil_vs = PCFData.FillCoilDataFromXML(2)
                    secondcoil_vs.Number = coilcounter
                    General.coillist.Add(secondcoil_vs)        'number = 4, each coil has a sandwich partner
                    secondcircuit_vs = PCFData.FillCircuitDataFromXML("First", 2, secondcoil_vs, circuitcounter)
                    secondcircuit_vs.Coilnumber = coilcounter
                    secondcircuit_mc.ConnectionSide = "left"
                    secondcircuit_vs.ConnectionSide = "right"
                    General.circuitlist.Add(secondcircuit_vs)
                    secondconsys_vs = PCFData.FillConSysDataFromXML(secondcircuit_vs, 2, "")
                    secondconsys_vs.Circnumber = circuitcounter
                    General.consyslist.Add(secondconsys_vs)
                    If General.currentunit.HasIntegratedSubcooler Then
                        circuitcounter += 1
                        secondsubcircuit_vs = PCFData.FillCircuitDataFromXML("Subcooler", 2, secondcoil_vs, circuitcounter)
                        secondsubcircuit_vs.Coilnumber = coilcounter
                        secondsubcircuit_mc.ConnectionSide = "left"
                        secondsubcircuit_vs.ConnectionSide = "right"
                        General.circuitlist.Add(secondsubcircuit_vs)
                        secondsubconsys_vs = PCFData.FillConSysDataFromXML(secondsubcircuit_vs, 2, "")
                        secondsubconsys_vs.Circnumber = circuitcounter
                        General.consyslist.Add(secondsubconsys_vs)
                    End If
                End If
            End If

            If General.currentunit.HasBrineDefrost Then
                brinecircuit = PCFData.CreateBrineCircuit(firstcircuit)
                brinecircuit.CircuitNumber = 2
                General.circuitlist.Add(brinecircuit)
                brineconsys = PCFData.FillConSysDataFromXML(firstcircuit, 1, "brine")
                brineconsys.Circnumber = 2
                General.consyslist.Add(brineconsys)
            End If

            For i As Integer = 1 To General.coillist.Count
                For j As Integer = 1 To General.circuitlist.Count
                    If General.circuitlist(j - 1).Coilnumber = i Then
                        General.coillist(i - 1).Circuits.Add(General.circuitlist(j - 1))
                        General.coillist(i - 1).ConSyss.Add(General.consyslist(j - 1))
                    End If
                Next
            Next

            General.currentunit.Coillist = General.coillist

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try
    End Sub

    Shared Sub CreateWS(ByRef job As JobInfo, username As String, Optional isCDBtest As Boolean = False)
        job.Workspace = WSM.StartWSM(CreateWSName, username, isCDBtest)
        WSM.fullpartids.Clear()
    End Sub

    Shared Function CreateWSName() As String
        Dim currenthour, currentmin, currentsec As Integer
        Dim today, wsmfolder As String
        Dim currentdate As Date = Date.Today

        currenthour = Date.Now.Hour
        currentmin = Date.Now.Minute
        currentsec = Date.Now.Second
        today = currentdate.ToShortDateString

        today = today.Replace(".", "_")
        today = today.Replace("/", "_")
        today = today.Replace(" ", "")
        wsmfolder = "CSG_" + today + "_" + currenthour.ToString + "_" + currentmin.ToString + "_" + currentsec.ToString

        Return wsmfolder
    End Function

    Shared Sub AddOrderDatatoCustomProps(obj As Object, filetype As String, pdmdesignation_de As String, pdmdesignation_en As String, adddesignation As String)
        Dim objpsets As SolidEdgeFramework.PropertySets
        Dim objgsets As SolidEdgeFramework.Properties
        Dim propindex As Integer

        Try
            Select Case filetype
                Case "par"
                    Dim partdoc As SolidEdgePart.PartDocument = CType(obj, SolidEdgePart.PartDocument)
                    objpsets = partdoc.Properties
                    objgsets = objpsets.Item(4)

                    SEUtils.ChangeProp(objgsets, 0, "Auftragsnummer", General.currentjob.OrderNumber)
                    SEUtils.ChangeProp(objgsets, 0, "Position", General.currentjob.OrderPosition)
                    SEUtils.ChangeProp(objgsets, 0, "CSG", "1")
                    SEUtils.ChangeProp(objgsets, 0, "Order_Projekt", General.currentjob.ProjectNumber)

                    If pdmdesignation_de <> "" Then
                        propindex = SEUtils.FindProp(objgsets, "CDB_Benennung_de")
                        SEUtils.ChangeProp(objgsets, propindex, "CDB_Benennung_de", pdmdesignation_de)
                    End If

                    If pdmdesignation_en <> "" Then
                        propindex = SEUtils.FindProp(objgsets, "CDB_Benennung_en")
                        SEUtils.ChangeProp(objgsets, propindex, "CDB_Benennung_en", pdmdesignation_en)
                    End If

                    If adddesignation <> "" Then
                        propindex = SEUtils.FindProp(objgsets, "CDB_Zusatzbenennung")
                        SEUtils.ChangeProp(objgsets, propindex, "CDB_Zusatzbenennung", adddesignation)
                    End If
            End Select
        Catch ex As Exception

        End Try


    End Sub

    Shared Function LNProdOrder(username As String, password As String, orderno As String, pos As String, plant As String) As String
        Dim dataStream As Stream
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim website, responseFromServer, payload As String
        Dim byteArray As Byte()

        responseFromServer = ""

        Try
            If General.currentjob.IsERPTest Then
                website = "https://deffbswui01-tst.europe.guentner-corp.com/c4ws/services/BillOfMaterial_Export/LN_Europe_TEST?wsdl"
            Else
                website = "https://deffbswui01.europe.guentner-corp.com/c4ws/services/BillOfMaterial_Export/LN_Europe_LIVE?wsdl"
            End If

            If plant = "Beji" Then
                website = "https://idpasswui01.asia.guentner-corp.com/c4ws/services/BillOfMaterial_Export/LN_APO_LIVE?wsdl"
            End If

            request = WebRequest.Create(website)
            request.ContentType = "text/xml;charset=utf-8"
            request.Method = "POST"

            If plant = "Beji" Then
                payload = GeneratePayloadBOMAPO(orderno, pos)
            Else
                payload = GeneratePayloadBOMEU(orderno, pos)
            End If

            byteArray = Encoding.UTF8.GetBytes(payload)
            request.ContentLength = byteArray.Length

            Dim header As String = CreateAuth(username, password)
            request.Headers.Add(HttpRequestHeader.Authorization, header)
            request.Headers.Add("SOAP:Action")

            dataStream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()

            response = request.GetResponse()
            Using dataStream1 As Stream = response.GetResponseStream()
                Dim reader As New StreamReader(dataStream1)
                responseFromServer = reader.ReadToEnd()
            End Using

            response.Close()

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return responseFromServer
    End Function

    Shared Function GeneratePayloadBOMAPO(orderno As String, orderpos As String)
        Dim soap As New XElement("soapenv__Envelope",
                                 New XAttribute("xmlns__soapenv", "http://schemas.xmlsoap.org/soap/envelope/"),
                                 New XAttribute("xmlns__bil", "http://www.infor.com/businessinterface/BillOfMaterial_Export"),
                                            New XElement("soapenv__Header",
                                                    New XElement("bil__Activation",
                                                         New XElement("username", "web_bom"),
                                                         New XElement("password", "dxv?0cc5xot7jcz-1gAQ"),
                                                         New XElement("company", "7000"))),
                                            New XElement("soapenv__Body",
                                                         New XElement("bil__ShowBOM",
                                                                      New XElement("ShowBOMRequest",
                                                                                   New XElement("ControlArea",
                                                                                                New XElement("processingScope", "request"),
                                                                                                New XElement("SuppressNillable", "?")),
                                                                                   New XElement("DataArea",
                                                                                                New XElement("BillOfMaterial_Export",
                                                                                                             New XElement("OrderNumber", orderno),
                                                                                                             New XElement("OrderLine", orderpos)))))))
        Dim ssoap As String = soap.ToString
        ssoap = ssoap.Replace("__", ":")

        Return ssoap.ToString
    End Function

    Shared Function GeneratePayloadBOMEU(orderno As String, orderpos As String)
        Dim soap As New XElement("soapenv__Envelope",
                                            New XAttribute("xmlns__soapenv", "http://schemas.xmlsoap.org/soap/envelope/"),
                                            New XAttribute("xmlns__bom", "http://www.infor.com/businessinterface/BillOfMaterial_Export"),
                                            New XElement("soapenv__Body",
                                                         New XElement("bom__ShowBOM",
                                                                      New XElement("ShowBOMRequest",
                                                                                   New XElement("ControlArea",
                                                                                                New XElement("processingScope", "request"),
                                                                                                New XElement("SuppressNillable", "?")),
                                                                                   New XElement("DataArea",
                                                                                                New XElement("BillOfMaterial_Export",
                                                                                                             New XElement("OrderNumber", orderno),
                                                                                                             New XElement("OrderLine", orderpos)))))))
        Dim ssoap As String = soap.ToString
        ssoap = ssoap.Replace("__", ":")

        Return ssoap.ToString
    End Function

    Shared Function CreateAuth(login As String, password As String) As String
        Dim auth As String = String.Format("{0}:{1}", login, password)
        Dim encoded As String = Convert.ToBase64String(Encoding.UTF8.GetBytes(auth))
        Return String.Format("{0},{1}", "Basic", encoded)
    End Function

    Shared Sub CheckForPlant(BOMList As List(Of BOMItem))
        For i As Integer = 0 To BOMList.Count - 1
            If BOMList(i).Warehouse <> "" AndAlso BOMList(i).Item.Contains(General.currentjob.ProjectNumber) Then
                If BOMList(i).Warehouse.Substring(0, 1) = "5" Then
                    General.currentjob.Plant = "Sibiu"
                    Exit For
                End If
            End If
        Next
    End Sub

    Shared Function ConstructBOM(response As String) As List(Of BOMItem)
        Dim xmldoc As New XmlDocument
        Dim artnolist, descriplist, itemnolist, parentitemnolist, poslist, quantitylist, warehouselist, itemgrouplist, optionlist As List(Of String)
        Dim BOMList As New List(Of BOMItem)
        Dim nodelist As XmlNodeList
        Dim itemcount As Integer

        Try
            xmldoc.LoadXml(response)
            nodelist = xmldoc.GetElementsByTagName("NumberOfRecords")
            If nodelist.Count = 1 Then
                itemcount = nodelist.Item(0).InnerText
            End If

            artnolist = XMLDataList(xmldoc.GetElementsByTagName("Item"), itemcount)
            descriplist = XMLDataList(xmldoc.GetElementsByTagName("ItemDescription"), itemcount)
            itemnolist = XMLDataList(xmldoc.GetElementsByTagName("ItemNumber"), itemcount)
            parentitemnolist = XMLDataList(xmldoc.GetElementsByTagName("ParentItemNumber"), itemcount)
            poslist = XMLDataList(xmldoc.GetElementsByTagName("Position"), itemcount)
            quantitylist = XMLDataList(xmldoc.GetElementsByTagName("Quantity"), itemcount, "double")
            warehouselist = XMLDataList(xmldoc.GetElementsByTagName("Warehouse"), itemcount)
            itemgrouplist = XMLDataList(xmldoc.GetElementsByTagName("ItemGroup"), itemcount)
            optionlist = XMLDataList(xmldoc.GetElementsByTagName("Option"), itemcount)


            For i As Integer = 0 To itemcount - 1
                BOMList.Add(New BOMItem With {.Item = artnolist(i), .ItemDescription = descriplist(i), .ItemNumber = itemnolist(i), .ParentItemNumber = parentitemnolist(i),
                            .Position = poslist(i), .Quantity = quantitylist(i), .Warehouse = warehouselist(i), .ItemGroup = itemgrouplist(i), .OptionTag = optionlist(i)})
            Next

            If General.username = "mlewin" Then
                CheckForPlant(BOMList)
            End If

            If response <> "" AndAlso General.currentjob.ProjectNumber = "" Then
                General.currentjob.ProjectNumber = BOMList(0).Item.Substring(0, 9)
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return BOMList
    End Function

    Shared Function XMLDataList(nodelist As XmlNodeList, itemcount As Integer, Optional datatype As String = "string") As List(Of String)
        Dim valuelist As New List(Of String)

        For i As Integer = 0 To itemcount - 1
            Dim value As String
            If nodelist.Count = 0 Then
                value = 0
            Else
                Dim node As XmlNode = nodelist.Item(i)
                If datatype = "double" Then
                    value = General.TextToDouble(node.InnerText.Trim)
                Else
                    value = node.InnerText.Trim
                End If
            End If
            valuelist.Add(value)
        Next

        Return valuelist
    End Function

    Shared Function GetBOMItemsByQuery(BOMList As List(Of BOMItem), attrs As String(), values As String()) As List(Of BOMItem)
        Dim templist, plist As New List(Of BOMItem)

        Try
            Select Case attrs(0)
                Case "parent"
                    Dim firstlist = From p In BOMList Where p.ParentItemNumber = values(0)

                    templist = firstlist.ToList
                Case "item"
                    Dim firstlist = From p In BOMList Where p.Item.Contains(values(0))

                    templist = firstlist.ToList
                Case "description"
                    Dim firstlist = From p In BOMList Where p.ItemDescription.Contains(values(0))

                    templist = firstlist.ToList
                Case "position"
                    Dim firstlist = From p In BOMList Where p.Position = values(0)

                    templist = firstlist.ToList
                Case "tag"
                    Dim firstlist = From p In BOMList Where p.OptionTag.Contains(values(0))

                    templist = firstlist.ToList
            End Select

            If attrs.Count > 1 Then
                For i As Integer = 1 To attrs.Count - 1
                    templist = GetBOMItemByQuery(templist, attrs(i), values(i))
                Next
            End If

            plist = templist

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return plist
    End Function

    Shared Function GetBOMItemByQuery(BOMList As List(Of BOMItem), attr As String, value As String) As List(Of BOMItem)
        Dim plist As New List(Of BOMItem)

        Select Case attr
            Case "position"
                Dim temp = From p In BOMList Where p.Position = value

                plist = temp.ToList
            Case "parent"
                Dim temp = From p In BOMList Where p.ParentItemNumber = value

                plist = temp.ToList
            Case "description"
                Dim temp = From p In BOMList Where p.ItemDescription.Contains(value)

                plist = temp.ToList
            Case "item"
                Dim temp = From p In BOMList Where p.Item.Contains(value)

                plist = temp.ToList
            Case "tag"
                Dim firstlist = From p In BOMList Where p.OptionTag.Contains(value)

                plist = firstlist.ToList
        End Select

        Return plist
    End Function

    Shared Sub GetCoilBOMItem(ByRef coil As CoilData, circuit As CircuitData)

        Try
            Dim tlist As List(Of BOMItem) = GetBOMItemsByQuery(General.currentunit.BOMList, {"description", "tag"}, {circuit.FinType + "/" + coil.NoRows.ToString + "/", coil.Number})
            If tlist.Count > 0 Then
                coil.BOMItem = tlist.First
            ElseIf coil.Number = 1 Then
                tlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"description", "parent", "item"}, {circuit.FinType + "/" + coil.NoRows.ToString + "/", 1, General.currentjob.ProjectNumber})
                If tlist.Count > 0 Then
                    coil.BOMItem = tlist.First
                End If
            End If

            If tlist.Count = 0 Then
                If coil.Number = 2 Then
                    'calculate gap
                    coil.Gap = Calculation.CoilGap(General.currentunit.Coillist)
                    If General.currentunit.UnitDescription = "VShape" Then
                        Dim templist As List(Of BOMItem) = GetBOMItemsByQuery(General.currentunit.BOMList, {"description", "position"}, {circuit.FinType + "/" + coil.NoRows.ToString + "/", (coil.Number * 10).ToString})
                        If templist.Count > 0 Then
                            coil.BOMItem = templist.First
                        End If
                    End If
                Else
                    If General.currentjob.Plant = "Sibiu" Then
                        coil.BOMItem = General.currentunit.BOMList.First
                    ElseIf General.currentjob.Plant = "Beji" Then
                        Dim templist As List(Of BOMItem) = GetBOMItemByQuery(General.currentunit.BOMList, "position", "1500")
                        If templist.Count > 0 Then
                            coil.BOMItem = templist.First
                        Else
                            coil.BOMItem = New BOMItem With {.ParentItemNumber = -1}
                        End If
                    Else
                        If General.currentunit.UnitDescription = "VShape" Then
                            Dim templist As List(Of BOMItem) = GetBOMItemsByQuery(General.currentunit.BOMList, {"description", "position"}, {circuit.FinType + "/" + coil.NoRows.ToString + "/", (coil.Number * 10).ToString})
                            If templist.Count > 0 Then
                                coil.BOMItem = templist.First
                            End If
                        Else
                            Dim templist As List(Of BOMItem) = GetBOMItemByQuery(General.currentunit.BOMList, "position", "1410")
                            If templist.Count > 0 Then
                                coil.BOMItem = templist.First
                            Else
                                coil.BOMItem = New BOMItem With {.ParentItemNumber = -1}
                            End If
                        End If
                    End If
                End If
            End If

            If coil.BOMItem.Item = "" AndAlso General.currentunit.ApplicationType = "Evaporator" Then
                Dim templist As List(Of BOMItem) = GetBOMItemsByQuery(General.currentunit.BOMList, {"tag"}, {"0"})
                If templist.Count = 1 Then
                    coil.BOMItem = templist.First
                End If
            End If

        Catch ex As Exception
            Debug.Print("No BOM item for coil found")
        End Try

    End Sub

    Shared Function GetConsysBOMItem(consys As ConSysData, coil As CoilData, circtype As String) As BOMItem
        Dim bomlist As List(Of BOMItem)
        Dim bomitem As New BOMItem
        Dim circtag As String

        Try

            bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "item"}, {coil.BOMItem.ItemNumber, ".A." + General.currentjob.OrderPosition + "." + consys.Circnumber.ToString})
            If circtype = "Subcooler" Then
                circtag = "SC1"
            ElseIf circtype = "Defrost" Then
                circtag = "BDC"
            Else
                circtag = "FC" + consys.Circnumber.ToString
            End If
            If bomlist.Count > 0 Then
                bomitem = bomlist(0)
            Else
                bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "tag"}, {coil.BOMItem.ItemNumber, circtag})
                If bomlist.Count > 0 Then
                    bomitem = bomlist(0)
                Else
                    If General.currentunit.UnitDescription = "VShape" Then
                        bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"item", "parent"}, {".A." + General.currentjob.OrderPosition, coil.BOMItem.ItemNumber})
                        If bomlist.Count > 0 Then
                            bomitem = bomlist(0)
                        End If
                    Else
                        'GCDC example different item → search for ANSCHLUSS_PRODAVIT, Anschluss-system, Conenction system
                        bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "description"}, {coil.BOMItem.ItemNumber, "Connection"})
                        If bomlist.Count = 0 Then
                            bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "description"}, {coil.BOMItem.ItemNumber, "Anschluss"})
                            If bomlist.Count = 0 Then
                                bomlist = GetBOMItemsByQuery(General.currentunit.BOMList, {"parent", "description"}, {coil.BOMItem.ItemNumber, "ANSCHLUSS"})
                                If bomlist.Count = 0 Then
                                    bomitem = New BOMItem
                                Else
                                    bomitem = bomlist(0)
                                End If
                            Else
                                bomitem = bomlist(0)
                            End If
                        Else
                            bomitem = bomlist(0)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.Print("No BOM item for connection system found")
        End Try

        Return bomitem
    End Function

    Shared Function CreateAddDesignationCoil(circuit As CircuitData, coil As CoilData) As String
        Dim designationstring As String = ""

        Try
            If General.currentunit.ModelRangeName = "" Or General.currentunit.ModelRangeName = "NNNN" Then
                designationstring = "GCO"
                designationstring += "-" + circuit.ConnectionSide.Substring(0, 1).ToUpper + " "

                designationstring += circuit.FinType + "/"

                designationstring += coil.NoLayers.ToString + "/" + coil.NoRows.ToString + " "
                designationstring += circuit.NoPasses.ToString + "P" + circuit.NoDistributions.ToString + "S"

            Else
                designationstring = General.currentunit.ModelRangeName
                designationstring += "-" + circuit.ConnectionSide.Substring(0, 1).ToUpper + "-"

                designationstring += General.currentunit.ModelRangeSuffix + " "
                designationstring += circuit.Pressure.ToString + "b. "

                If circuit.CoreTube.Materialcodeletter = "C" Then
                    designationstring += "Cu"
                Else
                    designationstring += "VA"
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        If designationstring.Length > 30 Then
            designationstring = designationstring.Replace(" ", "")
            designationstring = designationstring.Replace("-L", "-L ")
            designationstring = designationstring.Replace("-R", "-R ")
        End If
        Return designationstring
    End Function

    Shared Function CreateAddDesignationConsys(consys As ConSysData, circuit As CircuitData) As String
        Dim designationstring As String = ""

        Try
            If General.currentunit.ModelRangeName = "" Then
                designationstring = "GCO"
                designationstring += "-" + circuit.ConnectionSide.Substring(0, 1).ToUpper + " "

                designationstring += circuit.FinType + "/"
                designationstring += circuit.NoPasses.ToString + "P" + circuit.NoDistributions.ToString + "S"
            Else
                designationstring = General.currentunit.ModelRangeName
                designationstring += "-" + circuit.ConnectionSide.Substring(0, 1).ToUpper + " "

                designationstring += General.currentunit.ModelRangeSuffix
                designationstring += circuit.Pressure.ToString + " b. "

                If circuit.CoreTube.Materialcodeletter = "C" Or circuit.CoreTube.Materialcodeletter = "D" Then
                    designationstring += "Cu"
                Else
                    designationstring += "VA"
                End If

                If consys.OutletHeaders.First.Tube.Quantity > 0 Then
                    designationstring += " D" + consys.OutletHeaders.First.Tube.Diameter.ToString
                End If
                If consys.InletHeaders.First.Tube.Quantity > 0 Then
                    designationstring += " / D" + consys.InletHeaders.First.Tube.Diameter.ToString
                End If

                designationstring += " " + consys.HeaderAlignment.Substring(0, 3)
            End If

            If designationstring.Length > 30 Then
                designationstring = designationstring.Replace(" ", "")
                designationstring = designationstring.Replace("-L", "-L ")
                designationstring = designationstring.Replace("-R", "-R ")
                If designationstring.Length > 30 Then
                    designationstring = designationstring.Substring(0, 30)
                End If
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)

        End Try

        Return designationstring
    End Function

    Shared Function CreateOrderCommentCoil(circuit As CircuitData, coil As CoilData) As String
        Dim commentstring As String = ""

        Try
            commentstring = circuit.NoPasses.ToString + "P" + circuit.NoDistributions.ToString + "S;"
            commentstring += circuit.FinType + "/"
            commentstring += coil.NoRows.ToString + "/" + coil.NoLayers.ToString + "/"
            commentstring += coil.FinnedLength.ToString + ";Circuiting:*"
            commentstring += circuit.PDMID
            If coil.EDefrostPDMID <> "" Then
                commentstring += ";Defrost:*"
                commentstring += coil.EDefrostPDMID
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return commentstring
    End Function

    Shared Function CreateOrderCommentConsys(consys As ConSysData, coilnumber As Integer) As String
        Dim commentstring As String = ""

        Try
            commentstring = "Consys No." + consys.Circnumber.ToString + " // Coil No." + coilnumber.ToString
            If consys.OutletHeaders.First.Tube.IsBrine Then
                commentstring += " Brine Defrost"
            End If

        Catch ex As Exception
            General.CreateLogEntry(ex.ToString)
        End Try

        Return commentstring
    End Function

    Shared Function CSGWebservice(website As String, unitdata As UnitData, method As String) As String
        Dim dataStream As Stream
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim responseFromServer, payload As String
        Dim byteArray As Byte()

        Try
            request = WebRequest.Create(website)
            request.ContentType = "application/json"
            request.Method = method

            If method = "POST" Then
                payload = PayloadCSG(unitdata, General.currentjob.OrderNumber, General.currentjob.OrderPosition)
                byteArray = Encoding.UTF8.GetBytes(payload)
                request.ContentLength = byteArray.Length

                dataStream = request.GetRequestStream()
                dataStream.Write(byteArray, 0, byteArray.Length)
                dataStream.Close()
            End If

            response = request.GetResponse()
            Using dataStream1 As Stream = response.GetResponseStream()
                Dim reader As New StreamReader(dataStream1)
                responseFromServer = reader.ReadToEnd()
            End Using

            response.Close()
            Debug.Print(responseFromServer)
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try

        Return responseFromServer
    End Function

    Shared Function PayloadCSG(unitdata As UnitData, orderno As String, orderpos As String) As String
        Dim payload, unitstr As String
        Dim jobji, jobm As Linq.JObject
        Dim joblist As New List(Of Linq.JObject)

        Try

            unitstr = JsonConvert.SerializeObject(unitdata)
            jobji = JsonConvert.DeserializeObject(unitstr)
            jobm = New Linq.JObject(New Linq.JProperty("unitdata", jobji),
                                    New Linq.JProperty("OrderNo", orderno),
                                    New Linq.JProperty("OrderPos", orderpos),
                                    New Linq.JProperty("SavePath", ""),
                                    New Linq.JProperty("Mastername", General.currentjob.PDMID),
                                    New Linq.JProperty("Uid", General.currentjob.Uid))

            payload = JsonConvert.SerializeObject(jobm)
            My.Computer.FileSystem.WriteAllText(General.currentjob.OrderDir + "\test.json", payload, False)

        Catch ex As Exception
            payload = ""
            Debug.Print(ex.ToString)
        End Try

        Return payload
    End Function
End Class
