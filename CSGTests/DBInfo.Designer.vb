<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DBInfo
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
        Me.CSGTable = New System.Windows.Forms.DataGridView()
        Me.CUid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COrderNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.COrderPos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CProjectNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CPrio = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CStatus = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CUsers = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CRequestT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BRefresh = New System.Windows.Forms.Button()
        Me.BShow = New System.Windows.Forms.Button()
        Me.Linfo = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.BCancel = New System.Windows.Forms.Button()
        CType(Me.CSGTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CSGTable
        '
        Me.CSGTable.AllowUserToAddRows = False
        Me.CSGTable.AllowUserToDeleteRows = False
        Me.CSGTable.AllowUserToResizeRows = False
        Me.CSGTable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.CSGTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.CSGTable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CUid, Me.COrderNo, Me.COrderPos, Me.CProjectNo, Me.CPrio, Me.CStatus, Me.CUsers, Me.CRequestT})
        Me.CSGTable.Location = New System.Drawing.Point(12, 132)
        Me.CSGTable.MaximumSize = New System.Drawing.Size(850, 900)
        Me.CSGTable.MinimumSize = New System.Drawing.Size(625, 60)
        Me.CSGTable.Name = "CSGTable"
        Me.CSGTable.RowHeadersVisible = False
        Me.CSGTable.Size = New System.Drawing.Size(625, 138)
        Me.CSGTable.TabIndex = 1
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
        'BRefresh
        '
        Me.BRefresh.Location = New System.Drawing.Point(12, 75)
        Me.BRefresh.Name = "BRefresh"
        Me.BRefresh.Size = New System.Drawing.Size(75, 23)
        Me.BRefresh.TabIndex = 2
        Me.BRefresh.Text = "Refresh"
        Me.BRefresh.UseVisualStyleBackColor = True
        '
        'BShow
        '
        Me.BShow.Location = New System.Drawing.Point(155, 75)
        Me.BShow.Name = "BShow"
        Me.BShow.Size = New System.Drawing.Size(112, 23)
        Me.BShow.TabIndex = 3
        Me.BShow.Text = "Show Batch Queue"
        Me.BShow.UseVisualStyleBackColor = True
        '
        'Linfo
        '
        Me.Linfo.AutoSize = True
        Me.Linfo.Location = New System.Drawing.Point(13, 113)
        Me.Linfo.Name = "Linfo"
        Me.Linfo.Size = New System.Drawing.Size(39, 13)
        Me.Linfo.TabIndex = 4
        Me.Linfo.Text = "Label1"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(255, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Search based on input on the main window."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(789, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Uses order number, position and selected mode (Desktop = 1, Batch = 3, Copied = 4" &
    ") to list all order requests from the job table in the list."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 35)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(721, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Refresh button updates the current search. Show Batch Queue lists all open jobs, " &
    "which are scheduled for the batch process."
        '
        'BCancel
        '
        Me.BCancel.Location = New System.Drawing.Point(655, 246)
        Me.BCancel.Name = "BCancel"
        Me.BCancel.Size = New System.Drawing.Size(75, 23)
        Me.BCancel.TabIndex = 8
        Me.BCancel.Text = "Cancel Job"
        Me.BCancel.UseVisualStyleBackColor = True
        '
        'DBInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(801, 281)
        Me.Controls.Add(Me.BCancel)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Linfo)
        Me.Controls.Add(Me.BShow)
        Me.Controls.Add(Me.BRefresh)
        Me.Controls.Add(Me.CSGTable)
        Me.Name = "DBInfo"
        Me.Text = "DBInfo"
        CType(Me.CSGTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CSGTable As DataGridView
    Friend WithEvents BRefresh As Button
    Friend WithEvents BShow As Button
    Friend WithEvents Linfo As Label
    Friend WithEvents CUid As DataGridViewTextBoxColumn
    Friend WithEvents COrderNo As DataGridViewTextBoxColumn
    Friend WithEvents COrderPos As DataGridViewTextBoxColumn
    Friend WithEvents CProjectNo As DataGridViewTextBoxColumn
    Friend WithEvents CPrio As DataGridViewTextBoxColumn
    Friend WithEvents CStatus As DataGridViewTextBoxColumn
    Friend WithEvents CUsers As DataGridViewTextBoxColumn
    Friend WithEvents CRequestT As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents BCancel As Button
End Class
