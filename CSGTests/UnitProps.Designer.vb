<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class UnitProps
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.CBLocation = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CBAlignment = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.BBuild = New System.Windows.Forms.Button()
        Me.CBCType = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.BAddCircuit = New System.Windows.Forms.Button()
        Me.BAddCoil = New System.Windows.Forms.Button()
        Me.TBDPDMID = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.LDP = New System.Windows.Forms.Label()
        Me.TBDistr = New System.Windows.Forms.TextBox()
        Me.TBBTCoil = New System.Windows.Forms.TextBox()
        Me.LBT = New System.Windows.Forms.Label()
        Me.TBFinnedDepth = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TBPasses = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TBFinnedHeight = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TBFinnedLength = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.TBCTOverhang = New System.Windows.Forms.TextBox()
        Me.TBQuantity1 = New System.Windows.Forms.TextBox()
        Me.LBQ1 = New System.Windows.Forms.Label()
        Me.TBPDMID1 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CBFinType = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CBCTMaterial = New System.Windows.Forms.ComboBox()
        Me.CBOP = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CBCircuit = New System.Windows.Forms.ComboBox()
        Me.LCircuit = New System.Windows.Forms.Label()
        Me.LCoil = New System.Windows.Forms.Label()
        Me.CBCoil = New System.Windows.Forms.ComboBox()
        Me.BEditCoil = New System.Windows.Forms.Button()
        Me.BEditCircuit = New System.Windows.Forms.Button()
        Me.CBTSWT = New System.Windows.Forms.ComboBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.CBPunchingType = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.CBTSMat = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.BConSys = New System.Windows.Forms.Button()
        Me.BDrawing = New System.Windows.Forms.Button()
        Me.BSave = New System.Windows.Forms.Button()
        Me.BReset = New System.Windows.Forms.Button()
        Me.BReconnect = New System.Windows.Forms.Button()
        Me.BExport = New System.Windows.Forms.Button()
        Me.CBFin = New System.Windows.Forms.CheckBox()
        Me.BImport = New System.Windows.Forms.Button()
        Me.BEmptyConsys = New System.Windows.Forms.Button()
        Me.CheckOrbital = New System.Windows.Forms.CheckBox()
        Me.LGap = New System.Windows.Forms.Label()
        Me.TBGap = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'CBLocation
        '
        Me.CBLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBLocation.FormattingEnabled = True
        Me.CBLocation.Items.AddRange(New Object() {"right", "left"})
        Me.CBLocation.Location = New System.Drawing.Point(591, 184)
        Me.CBLocation.Name = "CBLocation"
        Me.CBLocation.Size = New System.Drawing.Size(50, 21)
        Me.CBLocation.TabIndex = 1025
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(447, 187)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(95, 13)
        Me.Label12.TabIndex = 332
        Me.Label12.Text = "Location rel. to AD"
        '
        'CBAlignment
        '
        Me.CBAlignment.AutoCompleteCustomSource.AddRange(New String() {"D", "E", "F", "G", "N"})
        Me.CBAlignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBAlignment.FormattingEnabled = True
        Me.CBAlignment.Items.AddRange(New Object() {"horizontal", "vertical"})
        Me.CBAlignment.Location = New System.Drawing.Point(203, 103)
        Me.CBAlignment.Name = "CBAlignment"
        Me.CBAlignment.Size = New System.Drawing.Size(88, 21)
        Me.CBAlignment.TabIndex = 1000
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(57, 106)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 13)
        Me.Label11.TabIndex = 330
        Me.Label11.Text = "Alignment"
        '
        'BBuild
        '
        Me.BBuild.Location = New System.Drawing.Point(59, 285)
        Me.BBuild.Name = "BBuild"
        Me.BBuild.Size = New System.Drawing.Size(169, 39)
        Me.BBuild.TabIndex = 328
        Me.BBuild.Text = "Build 3D"
        Me.BBuild.UseVisualStyleBackColor = True
        '
        'CBCType
        '
        Me.CBCType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBCType.FormattingEnabled = True
        Me.CBCType.Items.AddRange(New Object() {"First", "Second", "Brine Defrost", "Integrated Subcooler", "Sandwich"})
        Me.CBCType.Location = New System.Drawing.Point(591, 103)
        Me.CBCType.Name = "CBCType"
        Me.CBCType.Size = New System.Drawing.Size(139, 21)
        Me.CBCType.TabIndex = 1020
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(447, 106)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(31, 13)
        Me.Label7.TabIndex = 326
        Me.Label7.Text = "Type"
        '
        'BAddCircuit
        '
        Me.BAddCircuit.Location = New System.Drawing.Point(590, 51)
        Me.BAddCircuit.Name = "BAddCircuit"
        Me.BAddCircuit.Size = New System.Drawing.Size(75, 23)
        Me.BAddCircuit.TabIndex = 325
        Me.BAddCircuit.Text = "Add Circuit"
        Me.BAddCircuit.UseVisualStyleBackColor = True
        '
        'BAddCoil
        '
        Me.BAddCoil.Location = New System.Drawing.Point(203, 53)
        Me.BAddCoil.Name = "BAddCoil"
        Me.BAddCoil.Size = New System.Drawing.Size(75, 23)
        Me.BAddCoil.TabIndex = 324
        Me.BAddCoil.Text = "Add Coil"
        Me.BAddCoil.UseVisualStyleBackColor = True
        '
        'TBDPDMID
        '
        Me.TBDPDMID.Location = New System.Drawing.Point(203, 184)
        Me.TBDPDMID.Name = "TBDPDMID"
        Me.TBDPDMID.Size = New System.Drawing.Size(50, 20)
        Me.TBDPDMID.TabIndex = 1005
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(57, 187)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(140, 13)
        Me.Label10.TabIndex = 322
        Me.Label10.Text = "E. Defrost / Support PDMID"
        '
        'LDP
        '
        Me.LDP.AutoSize = True
        Me.LDP.Location = New System.Drawing.Point(622, 162)
        Me.LDP.Name = "LDP"
        Me.LDP.Size = New System.Drawing.Size(12, 13)
        Me.LDP.TabIndex = 321
        Me.LDP.Text = "x"
        '
        'TBDistr
        '
        Me.TBDistr.Location = New System.Drawing.Point(591, 158)
        Me.TBDistr.Name = "TBDistr"
        Me.TBDistr.Size = New System.Drawing.Size(25, 20)
        Me.TBDistr.TabIndex = 1023
        '
        'TBBTCoil
        '
        Me.TBBTCoil.Location = New System.Drawing.Point(203, 158)
        Me.TBBTCoil.Name = "TBBTCoil"
        Me.TBBTCoil.Size = New System.Drawing.Size(25, 20)
        Me.TBBTCoil.TabIndex = 1004
        Me.TBBTCoil.Text = "0"
        '
        'LBT
        '
        Me.LBT.AutoSize = True
        Me.LBT.Location = New System.Drawing.Point(57, 161)
        Me.LBT.Name = "LBT"
        Me.LBT.Size = New System.Drawing.Size(80, 13)
        Me.LBT.TabIndex = 320
        Me.LBT.Text = "No Blind Tubes"
        '
        'TBFinnedDepth
        '
        Me.TBFinnedDepth.Location = New System.Drawing.Point(352, 132)
        Me.TBFinnedDepth.Name = "TBFinnedDepth"
        Me.TBFinnedDepth.Size = New System.Drawing.Size(50, 20)
        Me.TBFinnedDepth.TabIndex = 1003
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(334, 135)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(12, 13)
        Me.Label4.TabIndex = 319
        Me.Label4.Text = "x"
        '
        'TBPasses
        '
        Me.TBPasses.Location = New System.Drawing.Point(640, 158)
        Me.TBPasses.Name = "TBPasses"
        Me.TBPasses.Size = New System.Drawing.Size(25, 20)
        Me.TBPasses.TabIndex = 1024
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(447, 163)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(138, 13)
        Me.Label3.TabIndex = 318
        Me.Label3.Text = "No of Distributions x Passes"
        '
        'TBFinnedHeight
        '
        Me.TBFinnedHeight.Location = New System.Drawing.Point(278, 132)
        Me.TBFinnedHeight.Name = "TBFinnedHeight"
        Me.TBFinnedHeight.Size = New System.Drawing.Size(50, 20)
        Me.TBFinnedHeight.TabIndex = 1002
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(260, 135)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(12, 13)
        Me.Label1.TabIndex = 317
        Me.Label1.Text = "x"
        '
        'TBFinnedLength
        '
        Me.TBFinnedLength.Location = New System.Drawing.Point(203, 132)
        Me.TBFinnedLength.Name = "TBFinnedLength"
        Me.TBFinnedLength.Size = New System.Drawing.Size(50, 20)
        Me.TBFinnedLength.TabIndex = 1001
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(57, 135)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(122, 13)
        Me.Label25.TabIndex = 316
        Me.Label25.Text = "Coil Size: L x H x D [mm]"
        '
        'TBCTOverhang
        '
        Me.TBCTOverhang.Location = New System.Drawing.Point(591, 294)
        Me.TBCTOverhang.Name = "TBCTOverhang"
        Me.TBCTOverhang.Size = New System.Drawing.Size(49, 20)
        Me.TBCTOverhang.TabIndex = 1029
        Me.TBCTOverhang.Text = "25"
        '
        'TBQuantity1
        '
        Me.TBQuantity1.Location = New System.Drawing.Point(662, 131)
        Me.TBQuantity1.Name = "TBQuantity1"
        Me.TBQuantity1.Size = New System.Drawing.Size(25, 20)
        Me.TBQuantity1.TabIndex = 1022
        Me.TBQuantity1.Text = "1"
        '
        'LBQ1
        '
        Me.LBQ1.AutoSize = True
        Me.LBQ1.Location = New System.Drawing.Point(646, 134)
        Me.LBQ1.Name = "LBQ1"
        Me.LBQ1.Size = New System.Drawing.Size(12, 13)
        Me.LBQ1.TabIndex = 315
        Me.LBQ1.Text = "x"
        '
        'TBPDMID1
        '
        Me.TBPDMID1.Location = New System.Drawing.Point(591, 131)
        Me.TBPDMID1.Name = "TBPDMID1"
        Me.TBPDMID1.Size = New System.Drawing.Size(50, 20)
        Me.TBPDMID1.TabIndex = 1021
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(447, 135)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(92, 13)
        Me.Label6.TabIndex = 314
        Me.Label6.Text = "PDMID x Quantity"
        '
        'CBFinType
        '
        Me.CBFinType.AutoCompleteCustomSource.AddRange(New String() {"D", "E", "F", "G", "N"})
        Me.CBFinType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBFinType.FormattingEnabled = True
        Me.CBFinType.Items.AddRange(New Object() {"D", "E", "F", "G", "H", "K", "M", "N"})
        Me.CBFinType.Location = New System.Drawing.Point(591, 210)
        Me.CBFinType.Name = "CBFinType"
        Me.CBFinType.Size = New System.Drawing.Size(50, 21)
        Me.CBFinType.TabIndex = 1026
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(447, 213)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 13)
        Me.Label5.TabIndex = 313
        Me.Label5.Text = "Fin Type"
        '
        'CBCTMaterial
        '
        Me.CBCTMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBCTMaterial.FormattingEnabled = True
        Me.CBCTMaterial.Items.AddRange(New Object() {"CU (SP01-A)", "CU (SP01-A1 (K65))", "CU (SP01-B1 (R))", "CU (SP01-C (X))", "V2A (SP03-1)", "V4A (SP03-2)"})
        Me.CBCTMaterial.Location = New System.Drawing.Point(591, 265)
        Me.CBCTMaterial.Name = "CBCTMaterial"
        Me.CBCTMaterial.Size = New System.Drawing.Size(139, 21)
        Me.CBCTMaterial.TabIndex = 1028
        '
        'CBOP
        '
        Me.CBOP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBOP.FormattingEnabled = True
        Me.CBOP.Items.AddRange(New Object() {"10", "16", "32", "41", "46", "54", "80", "120", "130"})
        Me.CBOP.Location = New System.Drawing.Point(591, 237)
        Me.CBOP.Name = "CBOP"
        Me.CBOP.Size = New System.Drawing.Size(50, 21)
        Me.CBOP.TabIndex = 1027
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(447, 268)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(97, 13)
        Me.Label9.TabIndex = 312
        Me.Label9.Text = "Core Tube Material"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(447, 240)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(121, 13)
        Me.Label8.TabIndex = 311
        Me.Label8.Text = "Operating Pressure [bar]"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(447, 297)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 13)
        Me.Label2.TabIndex = 310
        Me.Label2.Text = "Core Tube Overhang [mm]"
        '
        'CBCircuit
        '
        Me.CBCircuit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBCircuit.FormattingEnabled = True
        Me.CBCircuit.Location = New System.Drawing.Point(489, 53)
        Me.CBCircuit.Name = "CBCircuit"
        Me.CBCircuit.Size = New System.Drawing.Size(55, 21)
        Me.CBCircuit.TabIndex = 297
        '
        'LCircuit
        '
        Me.LCircuit.AutoSize = True
        Me.LCircuit.Location = New System.Drawing.Point(447, 56)
        Me.LCircuit.Name = "LCircuit"
        Me.LCircuit.Size = New System.Drawing.Size(36, 13)
        Me.LCircuit.TabIndex = 296
        Me.LCircuit.Text = "Circuit"
        '
        'LCoil
        '
        Me.LCoil.AutoSize = True
        Me.LCoil.Location = New System.Drawing.Point(57, 56)
        Me.LCoil.Name = "LCoil"
        Me.LCoil.Size = New System.Drawing.Size(24, 13)
        Me.LCoil.TabIndex = 295
        Me.LCoil.Text = "Coil"
        '
        'CBCoil
        '
        Me.CBCoil.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBCoil.FormattingEnabled = True
        Me.CBCoil.Location = New System.Drawing.Point(96, 53)
        Me.CBCoil.Name = "CBCoil"
        Me.CBCoil.Size = New System.Drawing.Size(59, 21)
        Me.CBCoil.TabIndex = 294
        '
        'BEditCoil
        '
        Me.BEditCoil.Location = New System.Drawing.Point(296, 53)
        Me.BEditCoil.Name = "BEditCoil"
        Me.BEditCoil.Size = New System.Drawing.Size(75, 23)
        Me.BEditCoil.TabIndex = 334
        Me.BEditCoil.Text = "Save Coil"
        Me.BEditCoil.UseVisualStyleBackColor = True
        '
        'BEditCircuit
        '
        Me.BEditCircuit.Location = New System.Drawing.Point(683, 51)
        Me.BEditCircuit.Name = "BEditCircuit"
        Me.BEditCircuit.Size = New System.Drawing.Size(75, 23)
        Me.BEditCircuit.TabIndex = 335
        Me.BEditCircuit.Text = "Save Circuit"
        Me.BEditCircuit.UseVisualStyleBackColor = True
        '
        'CBTSWT
        '
        Me.CBTSWT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBTSWT.FormattingEnabled = True
        Me.CBTSWT.Items.AddRange(New Object() {"1", "1,2", "1,25", "1,5", "2", "3"})
        Me.CBTSWT.Location = New System.Drawing.Point(353, 210)
        Me.CBTSWT.Name = "CBTSWT"
        Me.CBTSWT.Size = New System.Drawing.Size(50, 21)
        Me.CBTSWT.TabIndex = 1007
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(265, 213)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(81, 13)
        Me.Label24.TabIndex = 341
        Me.Label24.Text = "Thickness [mm]"
        '
        'CBPunchingType
        '
        Me.CBPunchingType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBPunchingType.FormattingEnabled = True
        Me.CBPunchingType.Items.AddRange(New Object() {"1", "2", "3"})
        Me.CBPunchingType.Location = New System.Drawing.Point(203, 237)
        Me.CBPunchingType.Name = "CBPunchingType"
        Me.CBPunchingType.Size = New System.Drawing.Size(50, 21)
        Me.CBPunchingType.TabIndex = 1008
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(57, 240)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(138, 13)
        Me.Label19.TabIndex = 340
        Me.Label19.Text = "Tube Sheet Punching Type"
        '
        'CBTSMat
        '
        Me.CBTSMat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBTSMat.FormattingEnabled = True
        Me.CBTSMat.Items.AddRange(New Object() {"G", "S", "V", "W"})
        Me.CBTSMat.Location = New System.Drawing.Point(203, 210)
        Me.CBTSMat.Name = "CBTSMat"
        Me.CBTSMat.Size = New System.Drawing.Size(50, 21)
        Me.CBTSMat.TabIndex = 1006
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(57, 213)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(103, 13)
        Me.Label18.TabIndex = 339
        Me.Label18.Text = "Tube Sheet Material"
        '
        'BConSys
        '
        Me.BConSys.Location = New System.Drawing.Point(278, 285)
        Me.BConSys.Name = "BConSys"
        Me.BConSys.Size = New System.Drawing.Size(129, 82)
        Me.BConSys.TabIndex = 342
        Me.BConSys.Text = "Open ConSys Props"
        Me.BConSys.UseVisualStyleBackColor = True
        '
        'BDrawing
        '
        Me.BDrawing.Location = New System.Drawing.Point(59, 328)
        Me.BDrawing.Name = "BDrawing"
        Me.BDrawing.Size = New System.Drawing.Size(169, 39)
        Me.BDrawing.TabIndex = 343
        Me.BDrawing.Text = "Create 2D"
        Me.BDrawing.UseVisualStyleBackColor = True
        '
        'BSave
        '
        Me.BSave.Location = New System.Drawing.Point(450, 328)
        Me.BSave.Name = "BSave"
        Me.BSave.Size = New System.Drawing.Size(129, 39)
        Me.BSave.TabIndex = 344
        Me.BSave.Text = "Save"
        Me.BSave.UseVisualStyleBackColor = True
        Me.BSave.Visible = False
        '
        'BReset
        '
        Me.BReset.Location = New System.Drawing.Point(3, 328)
        Me.BReset.Name = "BReset"
        Me.BReset.Size = New System.Drawing.Size(50, 39)
        Me.BReset.TabIndex = 345
        Me.BReset.Text = "Reset"
        Me.BReset.UseVisualStyleBackColor = True
        Me.BReset.Visible = False
        '
        'BReconnect
        '
        Me.BReconnect.Location = New System.Drawing.Point(713, 328)
        Me.BReconnect.Name = "BReconnect"
        Me.BReconnect.Size = New System.Drawing.Size(75, 39)
        Me.BReconnect.TabIndex = 346
        Me.BReconnect.Text = "Reconnect Solid Edge"
        Me.BReconnect.UseVisualStyleBackColor = True
        '
        'BExport
        '
        Me.BExport.Location = New System.Drawing.Point(713, 3)
        Me.BExport.Name = "BExport"
        Me.BExport.Size = New System.Drawing.Size(75, 23)
        Me.BExport.TabIndex = 347
        Me.BExport.Text = "Export"
        Me.BExport.UseVisualStyleBackColor = True
        '
        'CBFin
        '
        Me.CBFin.AutoSize = True
        Me.CBFin.Location = New System.Drawing.Point(693, 135)
        Me.CBFin.Name = "CBFin"
        Me.CBFin.Size = New System.Drawing.Size(95, 17)
        Me.CBFin.TabIndex = 348
        Me.CBFin.Text = "GCO Circuiting"
        Me.CBFin.UseVisualStyleBackColor = True
        Me.CBFin.Visible = False
        '
        'BImport
        '
        Me.BImport.Location = New System.Drawing.Point(59, 3)
        Me.BImport.Name = "BImport"
        Me.BImport.Size = New System.Drawing.Size(75, 23)
        Me.BImport.TabIndex = 1030
        Me.BImport.Text = "Import"
        Me.BImport.UseVisualStyleBackColor = True
        Me.BImport.Visible = False
        '
        'BEmptyConsys
        '
        Me.BEmptyConsys.Location = New System.Drawing.Point(278, 387)
        Me.BEmptyConsys.Name = "BEmptyConsys"
        Me.BEmptyConsys.Size = New System.Drawing.Size(129, 39)
        Me.BEmptyConsys.TabIndex = 1031
        Me.BEmptyConsys.Text = "Create empty Consys"
        Me.BEmptyConsys.UseVisualStyleBackColor = True
        '
        'CheckOrbital
        '
        Me.CheckOrbital.AutoSize = True
        Me.CheckOrbital.Location = New System.Drawing.Point(693, 214)
        Me.CheckOrbital.Name = "CheckOrbital"
        Me.CheckOrbital.Size = New System.Drawing.Size(56, 17)
        Me.CheckOrbital.TabIndex = 1032
        Me.CheckOrbital.Text = "Orbital"
        Me.CheckOrbital.UseVisualStyleBackColor = True
        Me.CheckOrbital.Visible = False
        '
        'LGap
        '
        Me.LGap.AutoSize = True
        Me.LGap.Location = New System.Drawing.Point(265, 161)
        Me.LGap.Name = "LGap"
        Me.LGap.Size = New System.Drawing.Size(52, 13)
        Me.LGap.TabIndex = 1033
        Me.LGap.Text = "Gap [mm]"
        Me.LGap.Visible = False
        '
        'TBGap
        '
        Me.TBGap.Location = New System.Drawing.Point(352, 158)
        Me.TBGap.Name = "TBGap"
        Me.TBGap.Size = New System.Drawing.Size(50, 20)
        Me.TBGap.TabIndex = 1034
        Me.TBGap.Text = "0"
        Me.TBGap.Visible = False
        '
        'UnitProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 377)
        Me.Controls.Add(Me.TBGap)
        Me.Controls.Add(Me.LGap)
        Me.Controls.Add(Me.CheckOrbital)
        Me.Controls.Add(Me.BEmptyConsys)
        Me.Controls.Add(Me.BImport)
        Me.Controls.Add(Me.CBFin)
        Me.Controls.Add(Me.BExport)
        Me.Controls.Add(Me.BReconnect)
        Me.Controls.Add(Me.BReset)
        Me.Controls.Add(Me.BSave)
        Me.Controls.Add(Me.BDrawing)
        Me.Controls.Add(Me.BConSys)
        Me.Controls.Add(Me.CBTSWT)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.CBPunchingType)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.CBTSMat)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.BEditCircuit)
        Me.Controls.Add(Me.BEditCoil)
        Me.Controls.Add(Me.CBLocation)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.CBAlignment)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.BBuild)
        Me.Controls.Add(Me.CBCType)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.BAddCircuit)
        Me.Controls.Add(Me.BAddCoil)
        Me.Controls.Add(Me.TBDPDMID)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.LDP)
        Me.Controls.Add(Me.TBDistr)
        Me.Controls.Add(Me.TBBTCoil)
        Me.Controls.Add(Me.LBT)
        Me.Controls.Add(Me.TBFinnedDepth)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TBPasses)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TBFinnedHeight)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TBFinnedLength)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.TBCTOverhang)
        Me.Controls.Add(Me.TBQuantity1)
        Me.Controls.Add(Me.LBQ1)
        Me.Controls.Add(Me.TBPDMID1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.CBFinType)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CBCTMaterial)
        Me.Controls.Add(Me.CBOP)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CBCircuit)
        Me.Controls.Add(Me.LCircuit)
        Me.Controls.Add(Me.LCoil)
        Me.Controls.Add(Me.CBCoil)
        Me.Name = "UnitProps"
        Me.Text = "UnitProps"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CBLocation As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents CBAlignment As ComboBox
    Friend WithEvents Label11 As Label
    Friend WithEvents BBuild As Button
    Friend WithEvents CBCType As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents BAddCircuit As Button
    Friend WithEvents BAddCoil As Button
    Friend WithEvents TBDPDMID As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents LDP As Label
    Friend WithEvents TBDistr As TextBox
    Friend WithEvents TBBTCoil As TextBox
    Friend WithEvents LBT As Label
    Friend WithEvents TBFinnedDepth As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TBPasses As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TBFinnedHeight As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TBFinnedLength As TextBox
    Friend WithEvents Label25 As Label
    Friend WithEvents TBCTOverhang As TextBox
    Friend WithEvents TBQuantity1 As TextBox
    Friend WithEvents LBQ1 As Label
    Friend WithEvents TBPDMID1 As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents CBFinType As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents CBCTMaterial As ComboBox
    Friend WithEvents CBOP As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents CBCircuit As ComboBox
    Friend WithEvents LCircuit As Label
    Friend WithEvents LCoil As Label
    Friend WithEvents CBCoil As ComboBox
    Friend WithEvents BEditCoil As Button
    Friend WithEvents BEditCircuit As Button
    Friend WithEvents CBTSWT As ComboBox
    Friend WithEvents Label24 As Label
    Friend WithEvents CBPunchingType As ComboBox
    Friend WithEvents Label19 As Label
    Friend WithEvents CBTSMat As ComboBox
    Friend WithEvents Label18 As Label
    Friend WithEvents BConSys As Button
    Friend WithEvents BDrawing As Button
    Friend WithEvents BSave As Button
    Friend WithEvents BReset As Button
    Friend WithEvents BReconnect As Button
    Friend WithEvents BExport As Button
    Friend WithEvents CBFin As CheckBox
    Friend WithEvents BImport As Button
    Friend WithEvents BEmptyConsys As Button
    Friend WithEvents CheckOrbital As CheckBox
    Friend WithEvents LGap As Label
    Friend WithEvents TBGap As TextBox
End Class
