<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class JobSelection
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CSGTable = New System.Windows.Forms.DataGridView()
        Me.CUid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COrderNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COrderPos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CProjectNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CPrio = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CStatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CUsers = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CRequestT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PDMID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BOK = New System.Windows.Forms.Button()
        CType(Me.CSGTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(278, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Finished Jobs in Database of CSG for this Order"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(131, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Select 1 and click OK"
        '
        'CSGTable
        '
        Me.CSGTable.AllowUserToAddRows = False
        Me.CSGTable.AllowUserToDeleteRows = False
        Me.CSGTable.AllowUserToResizeRows = False
        Me.CSGTable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.CSGTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.CSGTable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CUid, Me.COrderNo, Me.COrderPos, Me.CProjectNo, Me.CPrio, Me.CStatus, Me.CUsers, Me.CRequestT, Me.PDMID})
        Me.CSGTable.Location = New System.Drawing.Point(15, 59)
        Me.CSGTable.MaximumSize = New System.Drawing.Size(850, 900)
        Me.CSGTable.MinimumSize = New System.Drawing.Size(625, 60)
        Me.CSGTable.Name = "CSGTable"
        Me.CSGTable.RowHeadersVisible = False
        Me.CSGTable.Size = New System.Drawing.Size(693, 138)
        Me.CSGTable.TabIndex = 9
        '
        'CUid
        '
        Me.CUid.HeaderText = "Uid"
        Me.CUid.Name = "CUid"
        Me.CUid.ReadOnly = True
        Me.CUid.Width = 48
        '
        'COrderNo
        '
        Me.COrderNo.HeaderText = "Order Number"
        Me.COrderNo.Name = "COrderNo"
        Me.COrderNo.ReadOnly = True
        Me.COrderNo.Width = 90
        '
        'COrderPos
        '
        Me.COrderPos.HeaderText = "Order Position"
        Me.COrderPos.Name = "COrderPos"
        Me.COrderPos.ReadOnly = True
        Me.COrderPos.Width = 90
        '
        'CProjectNo
        '
        Me.CProjectNo.HeaderText = "Project Number"
        Me.CProjectNo.Name = "CProjectNo"
        Me.CProjectNo.ReadOnly = True
        Me.CProjectNo.Width = 97
        '
        'CPrio
        '
        Me.CPrio.HeaderText = "Prio"
        Me.CPrio.Name = "CPrio"
        Me.CPrio.ReadOnly = True
        Me.CPrio.Width = 50
        '
        'CStatus
        '
        Me.CStatus.HeaderText = "Status"
        Me.CStatus.Name = "CStatus"
        Me.CStatus.ReadOnly = True
        Me.CStatus.Width = 62
        '
        'CUsers
        '
        Me.CUsers.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CUsers.HeaderText = "User"
        Me.CUsers.Name = "CUsers"
        Me.CUsers.ReadOnly = True
        Me.CUsers.Width = 62
        '
        'CRequestT
        '
        Me.CRequestT.HeaderText = "Request Time"
        Me.CRequestT.Name = "CRequestT"
        Me.CRequestT.ReadOnly = True
        Me.CRequestT.Width = 90
        '
        'PDMID
        '
        Me.PDMID.HeaderText = "PDMID"
        Me.PDMID.Name = "PDMID"
        Me.PDMID.ReadOnly = True
        Me.PDMID.Width = 67
        '
        'BOK
        '
        Me.BOK.Location = New System.Drawing.Point(15, 204)
        Me.BOK.Name = "BOK"
        Me.BOK.Size = New System.Drawing.Size(97, 35)
        Me.BOK.TabIndex = 10
        Me.BOK.Text = "OK"
        Me.BOK.UseVisualStyleBackColor = True
        '
        'JobSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(720, 250)
        Me.Controls.Add(Me.BOK)
        Me.Controls.Add(Me.CSGTable)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "JobSelection"
        Me.Text = "Job Selection"
        CType(Me.CSGTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents CSGTable As DataGridView
    Friend WithEvents CUid As DataGridViewTextBoxColumn
    Friend WithEvents COrderNo As DataGridViewTextBoxColumn
    Friend WithEvents COrderPos As DataGridViewTextBoxColumn
    Friend WithEvents CProjectNo As DataGridViewTextBoxColumn
    Friend WithEvents CPrio As DataGridViewTextBoxColumn
    Friend WithEvents CStatus As DataGridViewTextBoxColumn
    Friend WithEvents CUsers As DataGridViewTextBoxColumn
    Friend WithEvents CRequestT As DataGridViewTextBoxColumn
    Friend WithEvents PDMID As DataGridViewTextBoxColumn
    Friend WithEvents BOK As Button
End Class
