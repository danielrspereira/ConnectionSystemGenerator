<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CopyProps
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LNewOrder = New System.Windows.Forms.Label()
        Me.LOldOrder = New System.Windows.Forms.Label()
        Me.TBNr1 = New System.Windows.Forms.TextBox()
        Me.LProject1 = New System.Windows.Forms.Label()
        Me.TBPos1 = New System.Windows.Forms.TextBox()
        Me.TBOrder1 = New System.Windows.Forms.TextBox()
        Me.LPosition1 = New System.Windows.Forms.Label()
        Me.LOrder1 = New System.Windows.Forms.Label()
        Me.TBNr2 = New System.Windows.Forms.TextBox()
        Me.LProject2 = New System.Windows.Forms.Label()
        Me.TBPos2 = New System.Windows.Forms.TextBox()
        Me.TBOrder2 = New System.Windows.Forms.TextBox()
        Me.LPosition2 = New System.Windows.Forms.Label()
        Me.LOrder2 = New System.Windows.Forms.Label()
        Me.TBMaster = New System.Windows.Forms.TextBox()
        Me.LMasterAsm = New System.Windows.Forms.Label()
        Me.BLoad = New System.Windows.Forms.Button()
        Me.LSteps = New System.Windows.Forms.Label()
        Me.TV1 = New System.Windows.Forms.TreeView()
        Me.BAnalyze = New System.Windows.Forms.Button()
        Me.BWAnalyze = New System.ComponentModel.BackgroundWorker()
        Me.BRename = New System.Windows.Forms.Button()
        Me.BSave = New System.Windows.Forms.Button()
        Me.BFind = New System.Windows.Forms.Button()
        Me.RTBLog = New System.Windows.Forms.RichTextBox()
        Me.BWRename = New System.ComponentModel.BackgroundWorker()
        Me.LStructure = New System.Windows.Forms.Label()
        Me.BExport = New System.Windows.Forms.Button()
        Me.BImport = New System.Windows.Forms.Button()
        Me.BReconnect = New System.Windows.Forms.Button()
        Me.BCopy = New System.Windows.Forms.Button()
        Me.CheckAGP = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'LNewOrder
        '
        Me.LNewOrder.AutoSize = True
        Me.LNewOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LNewOrder.Location = New System.Drawing.Point(92, 25)
        Me.LNewOrder.Name = "LNewOrder"
        Me.LNewOrder.Size = New System.Drawing.Size(122, 17)
        Me.LNewOrder.TabIndex = 0
        Me.LNewOrder.Text = "Your new Order"
        '
        'LOldOrder
        '
        Me.LOldOrder.AutoSize = True
        Me.LOldOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LOldOrder.Location = New System.Drawing.Point(342, 25)
        Me.LOldOrder.Name = "LOldOrder"
        Me.LOldOrder.Size = New System.Drawing.Size(214, 17)
        Me.LOldOrder.TabIndex = 1
        Me.LOldOrder.Text = "Order you want to copy from"
        '
        'TBNr1
        '
        Me.TBNr1.Location = New System.Drawing.Point(129, 109)
        Me.TBNr1.Name = "TBNr1"
        Me.TBNr1.ReadOnly = True
        Me.TBNr1.Size = New System.Drawing.Size(121, 20)
        Me.TBNr1.TabIndex = 71
        '
        'LProject1
        '
        Me.LProject1.AutoSize = True
        Me.LProject1.Location = New System.Drawing.Point(19, 112)
        Me.LProject1.Name = "LProject1"
        Me.LProject1.Size = New System.Drawing.Size(80, 13)
        Me.LProject1.TabIndex = 74
        Me.LProject1.Text = "Project Number"
        '
        'TBPos1
        '
        Me.TBPos1.Location = New System.Drawing.Point(129, 83)
        Me.TBPos1.Name = "TBPos1"
        Me.TBPos1.ReadOnly = True
        Me.TBPos1.Size = New System.Drawing.Size(121, 20)
        Me.TBPos1.TabIndex = 70
        '
        'TBOrder1
        '
        Me.TBOrder1.Location = New System.Drawing.Point(129, 56)
        Me.TBOrder1.Name = "TBOrder1"
        Me.TBOrder1.ReadOnly = True
        Me.TBOrder1.Size = New System.Drawing.Size(121, 20)
        Me.TBOrder1.TabIndex = 69
        '
        'LPosition1
        '
        Me.LPosition1.AutoSize = True
        Me.LPosition1.Location = New System.Drawing.Point(19, 86)
        Me.LPosition1.Name = "LPosition1"
        Me.LPosition1.Size = New System.Drawing.Size(44, 13)
        Me.LPosition1.TabIndex = 73
        Me.LPosition1.Text = "Position"
        '
        'LOrder1
        '
        Me.LOrder1.AutoSize = True
        Me.LOrder1.Location = New System.Drawing.Point(19, 59)
        Me.LOrder1.Name = "LOrder1"
        Me.LOrder1.Size = New System.Drawing.Size(73, 13)
        Me.LOrder1.TabIndex = 72
        Me.LOrder1.Text = "Order Number"
        '
        'TBNr2
        '
        Me.TBNr2.Location = New System.Drawing.Point(441, 109)
        Me.TBNr2.Name = "TBNr2"
        Me.TBNr2.Size = New System.Drawing.Size(121, 20)
        Me.TBNr2.TabIndex = 77
        '
        'LProject2
        '
        Me.LProject2.AutoSize = True
        Me.LProject2.Location = New System.Drawing.Point(331, 112)
        Me.LProject2.Name = "LProject2"
        Me.LProject2.Size = New System.Drawing.Size(80, 13)
        Me.LProject2.TabIndex = 80
        Me.LProject2.Text = "Project Number"
        '
        'TBPos2
        '
        Me.TBPos2.Location = New System.Drawing.Point(441, 83)
        Me.TBPos2.Name = "TBPos2"
        Me.TBPos2.Size = New System.Drawing.Size(121, 20)
        Me.TBPos2.TabIndex = 76
        '
        'TBOrder2
        '
        Me.TBOrder2.Location = New System.Drawing.Point(441, 56)
        Me.TBOrder2.Name = "TBOrder2"
        Me.TBOrder2.Size = New System.Drawing.Size(121, 20)
        Me.TBOrder2.TabIndex = 75
        '
        'LPosition2
        '
        Me.LPosition2.AutoSize = True
        Me.LPosition2.Location = New System.Drawing.Point(331, 86)
        Me.LPosition2.Name = "LPosition2"
        Me.LPosition2.Size = New System.Drawing.Size(44, 13)
        Me.LPosition2.TabIndex = 79
        Me.LPosition2.Text = "Position"
        '
        'LOrder2
        '
        Me.LOrder2.AutoSize = True
        Me.LOrder2.Location = New System.Drawing.Point(331, 59)
        Me.LOrder2.Name = "LOrder2"
        Me.LOrder2.Size = New System.Drawing.Size(73, 13)
        Me.LOrder2.TabIndex = 78
        Me.LOrder2.Text = "Order Number"
        '
        'TBMaster
        '
        Me.TBMaster.Location = New System.Drawing.Point(441, 135)
        Me.TBMaster.Name = "TBMaster"
        Me.TBMaster.Size = New System.Drawing.Size(121, 20)
        Me.TBMaster.TabIndex = 81
        '
        'LMasterAsm
        '
        Me.LMasterAsm.AutoSize = True
        Me.LMasterAsm.Location = New System.Drawing.Point(331, 138)
        Me.LMasterAsm.Name = "LMasterAsm"
        Me.LMasterAsm.Size = New System.Drawing.Size(109, 13)
        Me.LMasterAsm.TabIndex = 82
        Me.LMasterAsm.Text = "PDM No. Master Asm"
        '
        'BLoad
        '
        Me.BLoad.Location = New System.Drawing.Point(582, 196)
        Me.BLoad.Name = "BLoad"
        Me.BLoad.Size = New System.Drawing.Size(89, 26)
        Me.BLoad.TabIndex = 83
        Me.BLoad.Text = "Load old Asm"
        Me.BLoad.UseVisualStyleBackColor = True
        '
        'LSteps
        '
        Me.LSteps.AutoSize = True
        Me.LSteps.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LSteps.Location = New System.Drawing.Point(19, 353)
        Me.LSteps.Name = "LSteps"
        Me.LSteps.Size = New System.Drawing.Size(35, 17)
        Me.LSteps.TabIndex = 84
        Me.LSteps.Text = "Log"
        '
        'TV1
        '
        Me.TV1.Location = New System.Drawing.Point(22, 196)
        Me.TV1.Name = "TV1"
        Me.TV1.Size = New System.Drawing.Size(493, 154)
        Me.TV1.TabIndex = 86
        '
        'BAnalyze
        '
        Me.BAnalyze.Location = New System.Drawing.Point(582, 228)
        Me.BAnalyze.Name = "BAnalyze"
        Me.BAnalyze.Size = New System.Drawing.Size(89, 26)
        Me.BAnalyze.TabIndex = 87
        Me.BAnalyze.Text = "Analyze"
        Me.BAnalyze.UseVisualStyleBackColor = True
        '
        'BWAnalyze
        '
        '
        'BRename
        '
        Me.BRename.Location = New System.Drawing.Point(582, 260)
        Me.BRename.Name = "BRename"
        Me.BRename.Size = New System.Drawing.Size(89, 26)
        Me.BRename.TabIndex = 88
        Me.BRename.Text = "Rename"
        Me.BRename.UseVisualStyleBackColor = True
        '
        'BSave
        '
        Me.BSave.Enabled = False
        Me.BSave.Location = New System.Drawing.Point(582, 292)
        Me.BSave.Name = "BSave"
        Me.BSave.Size = New System.Drawing.Size(89, 26)
        Me.BSave.TabIndex = 89
        Me.BSave.Text = "Save"
        Me.BSave.UseVisualStyleBackColor = True
        '
        'BFind
        '
        Me.BFind.Location = New System.Drawing.Point(582, 133)
        Me.BFind.Name = "BFind"
        Me.BFind.Size = New System.Drawing.Size(43, 23)
        Me.BFind.TabIndex = 90
        Me.BFind.Text = "Find"
        Me.BFind.UseVisualStyleBackColor = True
        '
        'RTBLog
        '
        Me.RTBLog.Location = New System.Drawing.Point(22, 385)
        Me.RTBLog.Name = "RTBLog"
        Me.RTBLog.Size = New System.Drawing.Size(493, 96)
        Me.RTBLog.TabIndex = 91
        Me.RTBLog.Text = ""
        '
        'BWRename
        '
        '
        'LStructure
        '
        Me.LStructure.AutoSize = True
        Me.LStructure.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LStructure.Location = New System.Drawing.Point(19, 167)
        Me.LStructure.Name = "LStructure"
        Me.LStructure.Size = New System.Drawing.Size(75, 17)
        Me.LStructure.TabIndex = 93
        Me.LStructure.Text = "Structure"
        '
        'BExport
        '
        Me.BExport.Location = New System.Drawing.Point(582, 49)
        Me.BExport.Name = "BExport"
        Me.BExport.Size = New System.Drawing.Size(89, 26)
        Me.BExport.TabIndex = 94
        Me.BExport.Text = "Export"
        Me.BExport.UseVisualStyleBackColor = True
        '
        'BImport
        '
        Me.BImport.Location = New System.Drawing.Point(582, 19)
        Me.BImport.Name = "BImport"
        Me.BImport.Size = New System.Drawing.Size(89, 26)
        Me.BImport.TabIndex = 95
        Me.BImport.Text = "Import"
        Me.BImport.UseVisualStyleBackColor = True
        Me.BImport.Visible = False
        '
        'BReconnect
        '
        Me.BReconnect.Location = New System.Drawing.Point(582, 454)
        Me.BReconnect.Name = "BReconnect"
        Me.BReconnect.Size = New System.Drawing.Size(89, 26)
        Me.BReconnect.TabIndex = 96
        Me.BReconnect.Text = "Reconnect"
        Me.BReconnect.UseVisualStyleBackColor = True
        '
        'BCopy
        '
        Me.BCopy.Location = New System.Drawing.Point(268, 21)
        Me.BCopy.Name = "BCopy"
        Me.BCopy.Size = New System.Drawing.Size(43, 23)
        Me.BCopy.TabIndex = 97
        Me.BCopy.Text = "Copy"
        Me.BCopy.UseVisualStyleBackColor = True
        '
        'CheckAGP
        '
        Me.CheckAGP.AutoSize = True
        Me.CheckAGP.Location = New System.Drawing.Point(582, 173)
        Me.CheckAGP.Name = "CheckAGP"
        Me.CheckAGP.Size = New System.Drawing.Size(101, 17)
        Me.CheckAGP.TabIndex = 98
        Me.CheckAGP.Text = "New AGP Items"
        Me.CheckAGP.UseVisualStyleBackColor = True
        '
        'CopyProps
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(686, 504)
        Me.Controls.Add(Me.CheckAGP)
        Me.Controls.Add(Me.BCopy)
        Me.Controls.Add(Me.BReconnect)
        Me.Controls.Add(Me.BImport)
        Me.Controls.Add(Me.BExport)
        Me.Controls.Add(Me.LStructure)
        Me.Controls.Add(Me.RTBLog)
        Me.Controls.Add(Me.BFind)
        Me.Controls.Add(Me.BSave)
        Me.Controls.Add(Me.BRename)
        Me.Controls.Add(Me.BAnalyze)
        Me.Controls.Add(Me.TV1)
        Me.Controls.Add(Me.LSteps)
        Me.Controls.Add(Me.BLoad)
        Me.Controls.Add(Me.TBMaster)
        Me.Controls.Add(Me.LMasterAsm)
        Me.Controls.Add(Me.TBNr2)
        Me.Controls.Add(Me.LProject2)
        Me.Controls.Add(Me.TBPos2)
        Me.Controls.Add(Me.TBOrder2)
        Me.Controls.Add(Me.LPosition2)
        Me.Controls.Add(Me.LOrder2)
        Me.Controls.Add(Me.TBNr1)
        Me.Controls.Add(Me.LProject1)
        Me.Controls.Add(Me.TBPos1)
        Me.Controls.Add(Me.TBOrder1)
        Me.Controls.Add(Me.LPosition1)
        Me.Controls.Add(Me.LOrder1)
        Me.Controls.Add(Me.LOldOrder)
        Me.Controls.Add(Me.LNewOrder)
        Me.Name = "CopyProps"
        Me.Text = "Copy from Order"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LNewOrder As Label
    Friend WithEvents LOldOrder As Label
    Friend WithEvents TBNr1 As TextBox
    Friend WithEvents LProject1 As Label
    Friend WithEvents TBPos1 As TextBox
    Friend WithEvents TBOrder1 As TextBox
    Friend WithEvents LPosition1 As Label
    Friend WithEvents LOrder1 As Label
    Friend WithEvents TBNr2 As TextBox
    Friend WithEvents LProject2 As Label
    Friend WithEvents TBPos2 As TextBox
    Friend WithEvents TBOrder2 As TextBox
    Friend WithEvents LPosition2 As Label
    Friend WithEvents LOrder2 As Label
    Friend WithEvents TBMaster As TextBox
    Friend WithEvents LMasterAsm As Label
    Friend WithEvents BLoad As Button
    Friend WithEvents LSteps As Label
    Friend WithEvents TV1 As TreeView
    Friend WithEvents BAnalyze As Button
    Friend WithEvents BWAnalyze As System.ComponentModel.BackgroundWorker
    Friend WithEvents BRename As Button
    Friend WithEvents BSave As Button
    Friend WithEvents BFind As Button
    Friend WithEvents RTBLog As RichTextBox
    Friend WithEvents BWRename As System.ComponentModel.BackgroundWorker
    Friend WithEvents LStructure As Label
    Friend WithEvents BExport As Button
    Friend WithEvents BImport As Button
    Friend WithEvents BReconnect As Button
    Friend WithEvents BCopy As Button
    Friend WithEvents CheckAGP As CheckBox
End Class
