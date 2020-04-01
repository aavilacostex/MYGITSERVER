<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPDevelopmentseecomments
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPDevelopmentseecomments))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdprint = New System.Windows.Forms.Button()
        Me.cmddelete = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.dgvProjectMessages = New System.Windows.Forms.DataGridView()
        Me.clSubject = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clDateEntered = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clTimeEntered = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clUser = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtCode = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtpartno = New System.Windows.Forms.TextBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.dgvProjectMessage2 = New System.Windows.Forms.DataGridView()
        Me.clSelection = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.clComments = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblNotVisible = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvProjectMessages, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvProjectMessage2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lblNotVisible)
        Me.Panel1.Controls.Add(Me.cmdExit)
        Me.Panel1.Controls.Add(Me.cmdprint)
        Me.Panel1.Controls.Add(Me.cmddelete)
        Me.Panel1.Location = New System.Drawing.Point(12, 404)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(646, 34)
        Me.Panel1.TabIndex = 2
        '
        'cmdExit
        '
        Me.cmdExit.Image = CType(resources.GetObject("cmdExit.Image"), System.Drawing.Image)
        Me.cmdExit.Location = New System.Drawing.Point(588, 3)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(43, 28)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdprint
        '
        Me.cmdprint.Image = CType(resources.GetObject("cmdprint.Image"), System.Drawing.Image)
        Me.cmdprint.Location = New System.Drawing.Point(539, 3)
        Me.cmdprint.Name = "cmdprint"
        Me.cmdprint.Size = New System.Drawing.Size(43, 28)
        Me.cmdprint.TabIndex = 4
        Me.cmdprint.UseVisualStyleBackColor = True
        '
        'cmddelete
        '
        Me.cmddelete.Location = New System.Drawing.Point(484, 3)
        Me.cmddelete.Name = "cmddelete"
        Me.cmddelete.Size = New System.Drawing.Size(49, 28)
        Me.cmddelete.TabIndex = 3
        Me.cmddelete.Text = "Delete"
        Me.cmddelete.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(655, 382)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Menu
        Me.TabPage1.Controls.Add(Me.dgvProjectMessages)
        Me.TabPage1.Controls.Add(Me.TableLayoutPanel1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(647, 356)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TabPage1"
        '
        'dgvProjectMessages
        '
        Me.dgvProjectMessages.AllowUserToResizeColumns = False
        Me.dgvProjectMessages.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvProjectMessages.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProjectMessages.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clSubject, Me.clDateEntered, Me.clTimeEntered, Me.clUser})
        Me.dgvProjectMessages.Location = New System.Drawing.Point(11, 60)
        Me.dgvProjectMessages.Name = "dgvProjectMessages"
        Me.dgvProjectMessages.RowHeadersVisible = False
        Me.dgvProjectMessages.Size = New System.Drawing.Size(619, 265)
        Me.dgvProjectMessages.TabIndex = 2
        '
        'clSubject
        '
        Me.clSubject.HeaderText = "Subject"
        Me.clSubject.Name = "clSubject"
        '
        'clDateEntered
        '
        Me.clDateEntered.HeaderText = "Date Entered"
        Me.clDateEntered.Name = "clDateEntered"
        '
        'clTimeEntered
        '
        Me.clTimeEntered.HeaderText = "Time Entered"
        Me.clTimeEntered.Name = "clTimeEntered"
        '
        'clUser
        '
        Me.clUser.HeaderText = "User"
        Me.clUser.Name = "clUser"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.12439!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.63328!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.28595!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.27948!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtCode, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtpartno, 3, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(11, 6)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(619, 37)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Margin = New System.Windows.Forms.Padding(10, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Project No."
        '
        'txtCode
        '
        Me.txtCode.Location = New System.Drawing.Point(108, 8)
        Me.txtCode.Margin = New System.Windows.Forms.Padding(3, 8, 3, 3)
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(185, 20)
        Me.txtCode.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(316, 10)
        Me.Label2.Margin = New System.Windows.Forms.Padding(10, 10, 3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Part No."
        '
        'txtpartno
        '
        Me.txtpartno.Location = New System.Drawing.Point(415, 8)
        Me.txtpartno.Margin = New System.Windows.Forms.Padding(3, 8, 3, 3)
        Me.txtpartno.Name = "txtpartno"
        Me.txtpartno.Size = New System.Drawing.Size(185, 20)
        Me.txtpartno.TabIndex = 3
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Menu
        Me.TabPage2.Controls.Add(Me.dgvProjectMessage2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(647, 356)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        '
        'dgvProjectMessage2
        '
        Me.dgvProjectMessage2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvProjectMessage2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProjectMessage2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clSelection, Me.clComments})
        Me.dgvProjectMessage2.Location = New System.Drawing.Point(7, 7)
        Me.dgvProjectMessage2.Name = "dgvProjectMessage2"
        Me.dgvProjectMessage2.RowHeadersVisible = False
        Me.dgvProjectMessage2.Size = New System.Drawing.Size(634, 213)
        Me.dgvProjectMessage2.TabIndex = 0
        '
        'clSelection
        '
        Me.clSelection.HeaderText = "Selection"
        Me.clSelection.Name = "clSelection"
        '
        'clComments
        '
        Me.clComments.HeaderText = "Comments"
        Me.clComments.Name = "clComments"
        '
        'lblNotVisible
        '
        Me.lblNotVisible.AutoSize = True
        Me.lblNotVisible.Location = New System.Drawing.Point(411, 4)
        Me.lblNotVisible.Name = "lblNotVisible"
        Me.lblNotVisible.Size = New System.Drawing.Size(0, 13)
        Me.lblNotVisible.TabIndex = 6
        Me.lblNotVisible.Visible = False
        '
        'frmPDevelopmentseecomments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(666, 450)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmPDevelopmentseecomments"
        Me.Text = "frmPDevelopmentseecomments"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvProjectMessages, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvProjectMessage2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As Panel
    Friend WithEvents cmdExit As Button
    Friend WithEvents cmdprint As Button
    Friend WithEvents cmddelete As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents dgvProjectMessages As DataGridView
    Friend WithEvents clSubject As DataGridViewTextBoxColumn
    Friend WithEvents clDateEntered As DataGridViewTextBoxColumn
    Friend WithEvents clTimeEntered As DataGridViewTextBoxColumn
    Friend WithEvents clUser As DataGridViewTextBoxColumn
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents txtCode As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtpartno As TextBox
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents dgvProjectMessage2 As DataGridView
    Friend WithEvents clSelection As DataGridViewCheckBoxColumn
    Friend WithEvents clComments As DataGridViewTextBoxColumn
    Friend WithEvents lblNotVisible As Label
End Class
