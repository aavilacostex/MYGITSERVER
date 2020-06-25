<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmLoadExcel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLoadExcel))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblVendorDesc = New System.Windows.Forms.Label()
        Me.btnValidVendor = New System.Windows.Forms.Button()
        Me.txtVendorNo = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmdhidden = New System.Windows.Forms.Button()
        Me.cmdExcel = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.clError = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVMVNUM2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDPTN2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AddReference = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.EditReference = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.clPRDSTS = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clVMVNUM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDMFR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDCTP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRDPTN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clPRHCOD = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.cmbPerCharge = New System.Windows.Forms.ComboBox()
        Me.dtProjectDate = New System.Windows.Forms.DateTimePicker()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnCheck = New System.Windows.Forms.Button()
        Me.btnSuccess = New System.Windows.Forms.Button()
        Me.txtDesc = New System.Windows.Forms.TextBox()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.txtProjectNo = New System.Windows.Forms.TextBox()
        Me.btnInsert = New System.Windows.Forms.Button()
        Me.lblDesc = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblPerCharge = New System.Windows.Forms.Label()
        Me.lblProjectDate = New System.Windows.Forms.Label()
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.lblProjectNo = New System.Windows.Forms.Label()
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.lblVendorNo = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "excel.png")
        Me.ImageList1.Images.SetKeyName(1, "493-4936787_free-png-search-icon-magnifying-glass-icon-png.png")
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 81.81818!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.18182!))
        Me.TableLayoutPanel4.Controls.Add(Me.txtVendorNo, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.btnValidVendor, 1, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.lblVendorDesc, 0, 1)
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(248, 163)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 2
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(239, 51)
        Me.TableLayoutPanel4.TabIndex = 31
        '
        'lblVendorDesc
        '
        Me.TableLayoutPanel4.SetColumnSpan(Me.lblVendorDesc, 2)
        Me.lblVendorDesc.Location = New System.Drawing.Point(3, 36)
        Me.lblVendorDesc.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.lblVendorDesc.Name = "lblVendorDesc"
        Me.lblVendorDesc.Size = New System.Drawing.Size(152, 10)
        Me.lblVendorDesc.TabIndex = 32
        Me.lblVendorDesc.Text = "  "
        '
        'btnValidVendor
        '
        Me.btnValidVendor.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnValidVendor.ImageIndex = 1
        Me.btnValidVendor.ImageList = Me.ImageList1
        Me.btnValidVendor.Location = New System.Drawing.Point(198, 3)
        Me.btnValidVendor.Name = "btnValidVendor"
        Me.btnValidVendor.Size = New System.Drawing.Size(24, 17)
        Me.btnValidVendor.TabIndex = 31
        Me.btnValidVendor.Text = " "
        Me.btnValidVendor.UseVisualStyleBackColor = True
        '
        'txtVendorNo
        '
        Me.txtVendorNo.Location = New System.Drawing.Point(3, 3)
        Me.txtVendorNo.Multiline = True
        Me.txtVendorNo.Name = "txtVendorNo"
        Me.txtVendorNo.Size = New System.Drawing.Size(189, 27)
        Me.txtVendorNo.TabIndex = 30
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.cmdExcel, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.cmdhidden, 1, 0)
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 624)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(239, 42)
        Me.TableLayoutPanel3.TabIndex = 28
        '
        'cmdhidden
        '
        Me.cmdhidden.Location = New System.Drawing.Point(122, 3)
        Me.cmdhidden.Name = "cmdhidden"
        Me.cmdhidden.Size = New System.Drawing.Size(50, 15)
        Me.cmdhidden.TabIndex = 30
        Me.cmdhidden.Text = "Button1"
        Me.cmdhidden.UseVisualStyleBackColor = True
        Me.cmdhidden.Visible = False
        '
        'cmdExcel
        '
        Me.cmdExcel.ImageIndex = 0
        Me.cmdExcel.ImageList = Me.ImageList1
        Me.cmdExcel.Location = New System.Drawing.Point(3, 3)
        Me.cmdExcel.Name = "cmdExcel"
        Me.cmdExcel.Size = New System.Drawing.Size(29, 19)
        Me.cmdExcel.TabIndex = 29
        Me.cmdExcel.UseVisualStyleBackColor = True
        Me.cmdExcel.Visible = False
        '
        'SplitContainer1
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.SplitContainer1, 3)
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 330)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView1)
        Me.SplitContainer1.Panel1Collapsed = True
        Me.SplitContainer1.Panel1MinSize = 60
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DataGridView2)
        Me.SplitContainer1.Panel2MinSize = 60
        Me.SplitContainer1.Size = New System.Drawing.Size(769, 254)
        Me.SplitContainer1.SplitterDistance = 60
        Me.SplitContainer1.TabIndex = 27
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.EditReference, Me.AddReference, Me.clPRDPTN2, Me.clVMVNUM2, Me.clError})
        Me.DataGridView2.Location = New System.Drawing.Point(4, 4)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(762, 246)
        Me.DataGridView2.TabIndex = 0
        Me.DataGridView2.Visible = False
        '
        'clError
        '
        Me.clError.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clError.HeaderText = "Error Description"
        Me.clError.Name = "clError"
        '
        'clVMVNUM2
        '
        Me.clVMVNUM2.HeaderText = "Vendor Number"
        Me.clVMVNUM2.Name = "clVMVNUM2"
        Me.clVMVNUM2.Width = 150
        '
        'clPRDPTN2
        '
        Me.clPRDPTN2.HeaderText = "Part Number"
        Me.clPRDPTN2.Name = "clPRDPTN2"
        Me.clPRDPTN2.Width = 150
        '
        'AddReference
        '
        Me.AddReference.HeaderText = "Add"
        Me.AddReference.Name = "AddReference"
        Me.AddReference.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.AddReference.Width = 50
        '
        'EditReference
        '
        Me.EditReference.HeaderText = "Edit"
        Me.EditReference.Name = "EditReference"
        Me.EditReference.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.EditReference.Width = 50
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clPRHCOD, Me.clPRDPTN, Me.clPRDCTP, Me.clPRDMFR, Me.clVMVNUM, Me.clPRDSTS})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView1.Location = New System.Drawing.Point(8, 8)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(758, 277)
        Me.DataGridView1.TabIndex = 10
        '
        'clPRDSTS
        '
        Me.clPRDSTS.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDSTS.FillWeight = 102.9442!
        Me.clPRDSTS.HeaderText = "Status"
        Me.clPRDSTS.MinimumWidth = 8
        Me.clPRDSTS.Name = "clPRDSTS"
        Me.clPRDSTS.ReadOnly = True
        '
        'clVMVNUM
        '
        Me.clVMVNUM.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clVMVNUM.FillWeight = 102.9442!
        Me.clVMVNUM.HeaderText = "Vendor No."
        Me.clVMVNUM.MinimumWidth = 8
        Me.clVMVNUM.Name = "clVMVNUM"
        Me.clVMVNUM.ReadOnly = True
        '
        'clPRDMFR
        '
        Me.clPRDMFR.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDMFR.FillWeight = 102.9442!
        Me.clPRDMFR.HeaderText = "Manufacturer No."
        Me.clPRDMFR.MinimumWidth = 8
        Me.clPRDMFR.Name = "clPRDMFR"
        Me.clPRDMFR.ReadOnly = True
        '
        'clPRDCTP
        '
        Me.clPRDCTP.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDCTP.FillWeight = 102.9442!
        Me.clPRDCTP.HeaderText = "CTP No."
        Me.clPRDCTP.MinimumWidth = 8
        Me.clPRDCTP.Name = "clPRDCTP"
        Me.clPRDCTP.ReadOnly = True
        '
        'clPRDPTN
        '
        Me.clPRDPTN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRDPTN.FillWeight = 102.9442!
        Me.clPRDPTN.HeaderText = "Part No."
        Me.clPRDPTN.MinimumWidth = 8
        Me.clPRDPTN.Name = "clPRDPTN"
        Me.clPRDPTN.ReadOnly = True
        '
        'clPRHCOD
        '
        Me.clPRHCOD.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.clPRHCOD.FillWeight = 85.27919!
        Me.clPRHCOD.HeaderText = "Project No."
        Me.clPRHCOD.MinimumWidth = 8
        Me.clPRHCOD.Name = "clPRHCOD"
        Me.clPRHCOD.ReadOnly = True
        '
        'cmbStatus
        '
        Me.cmbStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.IntegralHeight = False
        Me.cmbStatus.ItemHeight = 17
        Me.cmbStatus.Location = New System.Drawing.Point(493, 166)
        Me.cmbStatus.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(186, 25)
        Me.cmbStatus.TabIndex = 26
        '
        'cmbPerCharge
        '
        Me.cmbPerCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPerCharge.FormattingEnabled = True
        Me.cmbPerCharge.Location = New System.Drawing.Point(3, 166)
        Me.cmbPerCharge.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.cmbPerCharge.Name = "cmbPerCharge"
        Me.cmbPerCharge.Size = New System.Drawing.Size(239, 25)
        Me.cmbPerCharge.TabIndex = 25
        '
        'dtProjectDate
        '
        Me.dtProjectDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtProjectDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtProjectDate.Location = New System.Drawing.Point(493, 97)
        Me.dtProjectDate.Name = "dtProjectDate"
        Me.dtProjectDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dtProjectDate.Size = New System.Drawing.Size(279, 24)
        Me.dtProjectDate.TabIndex = 24
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TableLayoutPanel2.SetColumnSpan(Me.Panel2, 3)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(769, 50)
        Me.Panel2.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(85, 11)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(604, 25)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Please add the project header data and then add the excel with the references."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 285)
        Me.Label1.Margin = New System.Windows.Forms.Padding(3, 10, 3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(239, 42)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please select the file that you want to process"
        '
        'btnSelect
        '
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelect.Location = New System.Drawing.Point(315, 285)
        Me.btnSelect.Margin = New System.Windows.Forms.Padding(70, 10, 3, 3)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(147, 39)
        Me.btnSelect.TabIndex = 0
        Me.btnSelect.Text = "Load File"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Location = New System.Drawing.Point(493, 278)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(279, 46)
        Me.Panel1.TabIndex = 2
        Me.Panel1.Visible = False
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(95, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "RadioButton1"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(3, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(90, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "RadioButton2"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnSuccess, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCheck, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(248, 624)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(239, 42)
        Me.TableLayoutPanel1.TabIndex = 18
        '
        'btnCheck
        '
        Me.btnCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCheck.Location = New System.Drawing.Point(122, 6)
        Me.btnCheck.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.btnCheck.Name = "btnCheck"
        Me.btnCheck.Size = New System.Drawing.Size(75, 20)
        Me.btnCheck.TabIndex = 3
        Me.btnCheck.Text = "Check Errors"
        Me.btnCheck.UseVisualStyleBackColor = True
        '
        'btnSuccess
        '
        Me.btnSuccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSuccess.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSuccess.Location = New System.Drawing.Point(3, 6)
        Me.btnSuccess.Margin = New System.Windows.Forms.Padding(3, 6, 3, 3)
        Me.btnSuccess.Name = "btnSuccess"
        Me.btnSuccess.Size = New System.Drawing.Size(75, 20)
        Me.btnSuccess.TabIndex = 4
        Me.btnSuccess.Text = "Show Success"
        Me.btnSuccess.UseVisualStyleBackColor = True
        '
        'txtDesc
        '
        Me.TableLayoutPanel2.SetColumnSpan(Me.txtDesc, 3)
        Me.txtDesc.Location = New System.Drawing.Point(3, 242)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDesc.Size = New System.Drawing.Size(769, 30)
        Me.txtDesc.TabIndex = 16
        '
        'txtProjectName
        '
        Me.txtProjectName.Location = New System.Drawing.Point(248, 97)
        Me.txtProjectName.Multiline = True
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(239, 24)
        Me.txtProjectName.TabIndex = 12
        '
        'txtProjectNo
        '
        Me.txtProjectNo.Location = New System.Drawing.Point(3, 97)
        Me.txtProjectNo.Multiline = True
        Me.txtProjectNo.Name = "txtProjectNo"
        Me.txtProjectNo.Size = New System.Drawing.Size(239, 24)
        Me.txtProjectNo.TabIndex = 11
        '
        'btnInsert
        '
        Me.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnInsert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInsert.Location = New System.Drawing.Point(560, 629)
        Me.btnInsert.Margin = New System.Windows.Forms.Padding(70, 8, 3, 3)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(161, 37)
        Me.btnInsert.TabIndex = 10
        Me.btnInsert.Text = "Insert to DB"
        Me.btnInsert.UseVisualStyleBackColor = True
        Me.btnInsert.Visible = False
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.TableLayoutPanel2.SetColumnSpan(Me.lblDesc, 3)
        Me.lblDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesc.Location = New System.Drawing.Point(3, 222)
        Me.lblDesc.Margin = New System.Windows.Forms.Padding(3, 5, 3, 0)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(80, 15)
        Me.lblDesc.TabIndex = 8
        Me.lblDesc.Text = "Description"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(493, 139)
        Me.lblStatus.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(116, 15)
        Me.lblStatus.TabIndex = 7
        Me.lblStatus.Text = "Project Status (*)"
        '
        'lblPerCharge
        '
        Me.lblPerCharge.AutoSize = True
        Me.lblPerCharge.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPerCharge.Location = New System.Drawing.Point(3, 139)
        Me.lblPerCharge.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblPerCharge.Name = "lblPerCharge"
        Me.lblPerCharge.Size = New System.Drawing.Size(138, 15)
        Me.lblPerCharge.TabIndex = 6
        Me.lblPerCharge.Text = "Person in Charge (*)"
        '
        'lblProjectDate
        '
        Me.lblProjectDate.AutoSize = True
        Me.lblProjectDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectDate.Location = New System.Drawing.Point(493, 71)
        Me.lblProjectDate.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblProjectDate.Name = "lblProjectDate"
        Me.lblProjectDate.Size = New System.Drawing.Size(106, 15)
        Me.lblProjectDate.TabIndex = 5
        Me.lblProjectDate.Text = "Project Date (*)"
        '
        'lblProjectName
        '
        Me.lblProjectName.AutoSize = True
        Me.lblProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectName.Location = New System.Drawing.Point(248, 71)
        Me.lblProjectName.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(114, 15)
        Me.lblProjectName.TabIndex = 4
        Me.lblProjectName.Text = "Project Name (*)"
        '
        'lblProjectNo
        '
        Me.lblProjectNo.AutoSize = True
        Me.lblProjectNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectNo.Location = New System.Drawing.Point(3, 71)
        Me.lblProjectNo.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblProjectNo.Name = "lblProjectNo"
        Me.lblProjectNo.Size = New System.Drawing.Size(78, 15)
        Me.lblProjectNo.TabIndex = 3
        Me.lblProjectNo.Text = "Project No."
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.AutoSize = False
        Me.TableLayoutPanel2.SetColumnSpan(Me.BindingNavigator1, 3)
        Me.BindingNavigator1.CountItem = Me.BindingNavigatorCountItem
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.None
        Me.BindingNavigator1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(235, 587)
        Me.BindingNavigator1.Margin = New System.Windows.Forms.Padding(235, 0, 0, 0)
        Me.BindingNavigator1.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.BindingNavigator1.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.BindingNavigator1.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.BindingNavigator1.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Me.BindingNavigatorPositionItem
        Me.BindingNavigator1.Size = New System.Drawing.Size(378, 34)
        Me.BindingNavigator1.TabIndex = 1
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(28, 31)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(28, 31)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 34)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(75, 35)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 31)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 34)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(28, 31)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(28, 31)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 34)
        '
        'lblVendorNo
        '
        Me.lblVendorNo.AutoSize = True
        Me.lblVendorNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorNo.Location = New System.Drawing.Point(248, 139)
        Me.lblVendorNo.Margin = New System.Windows.Forms.Padding(3, 15, 3, 0)
        Me.lblVendorNo.Name = "lblVendorNo"
        Me.lblVendorNo.Size = New System.Drawing.Size(127, 15)
        Me.lblVendorNo.TabIndex = 30
        Me.lblVendorNo.Text = "Vendor Number (*)"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 3
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 285.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.lblVendorNo, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.BindingNavigator1, 0, 9)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectNo, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectName, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblProjectDate, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblPerCharge, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.lblStatus, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.lblDesc, 0, 5)
        Me.TableLayoutPanel2.Controls.Add(Me.btnInsert, 2, 10)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectNo, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtProjectName, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.txtDesc, 0, 6)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel1, 1, 10)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel1, 2, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.btnSelect, 1, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 0, 7)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.dtProjectDate, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbPerCharge, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.cmbStatus, 2, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.SplitContainer1, 0, 8)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel3, 0, 10)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel4, 1, 4)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(13, 21)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 11
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 57.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 52.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 260.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(775, 669)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'frmLoadExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 701)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Name = "frmLoadExcel"
        Me.Text = "frmLoadExcel"
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents BindingSource1 As BindingSource
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents txtVendorNo As TextBox
    Friend WithEvents btnValidVendor As Button
    Friend WithEvents lblVendorDesc As Label
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents cmdExcel As Button
    Friend WithEvents cmdhidden As Button
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents lblVendorNo As Label
    Friend WithEvents BindingNavigator1 As BindingNavigator
    Friend WithEvents BindingNavigatorCountItem As ToolStripLabel
    Friend WithEvents BindingNavigatorMoveFirstItem As ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As ToolStripSeparator
    Friend WithEvents lblProjectNo As Label
    Friend WithEvents lblProjectName As Label
    Friend WithEvents lblProjectDate As Label
    Friend WithEvents lblPerCharge As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblDesc As Label
    Friend WithEvents btnInsert As Button
    Friend WithEvents txtProjectNo As TextBox
    Friend WithEvents txtProjectName As TextBox
    Friend WithEvents txtDesc As TextBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents btnSuccess As Button
    Friend WithEvents btnCheck As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents btnSelect As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents dtProjectDate As DateTimePicker
    Friend WithEvents cmbPerCharge As ComboBox
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents clPRHCOD As DataGridViewTextBoxColumn
    Friend WithEvents clPRDPTN As DataGridViewTextBoxColumn
    Friend WithEvents clPRDCTP As DataGridViewTextBoxColumn
    Friend WithEvents clPRDMFR As DataGridViewTextBoxColumn
    Friend WithEvents clVMVNUM As DataGridViewTextBoxColumn
    Friend WithEvents clPRDSTS As DataGridViewTextBoxColumn
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents EditReference As DataGridViewLinkColumn
    Friend WithEvents AddReference As DataGridViewLinkColumn
    Friend WithEvents clPRDPTN2 As DataGridViewTextBoxColumn
    Friend WithEvents clVMVNUM2 As DataGridViewTextBoxColumn
    Friend WithEvents clError As DataGridViewTextBoxColumn
End Class
