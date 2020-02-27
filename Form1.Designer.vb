<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.PurchasingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomerClaimsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductsDevelopmentToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProductsDevelopmentToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.LoginToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PurchasingToolStripMenuItem, Me.LoginToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'PurchasingToolStripMenuItem
        '
        Me.PurchasingToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClaimsToolStripMenuItem, Me.ProductsDevelopmentToolStripMenuItem})
        Me.PurchasingToolStripMenuItem.Name = "PurchasingToolStripMenuItem"
        Me.PurchasingToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.PurchasingToolStripMenuItem.Text = "Purchasing"
        '
        'ClaimsToolStripMenuItem
        '
        Me.ClaimsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SupplierClaimsToolStripMenuItem, Me.CustomerClaimsToolStripMenuItem})
        Me.ClaimsToolStripMenuItem.Name = "ClaimsToolStripMenuItem"
        Me.ClaimsToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.ClaimsToolStripMenuItem.Text = "Claims"
        '
        'SupplierClaimsToolStripMenuItem
        '
        Me.SupplierClaimsToolStripMenuItem.Name = "SupplierClaimsToolStripMenuItem"
        Me.SupplierClaimsToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.SupplierClaimsToolStripMenuItem.Text = "Supplier Claims"
        '
        'CustomerClaimsToolStripMenuItem
        '
        Me.CustomerClaimsToolStripMenuItem.Name = "CustomerClaimsToolStripMenuItem"
        Me.CustomerClaimsToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.CustomerClaimsToolStripMenuItem.Text = "Customer Claims"
        '
        'ProductsDevelopmentToolStripMenuItem
        '
        Me.ProductsDevelopmentToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProductsDevelopmentToolStripMenuItem1})
        Me.ProductsDevelopmentToolStripMenuItem.Name = "ProductsDevelopmentToolStripMenuItem"
        Me.ProductsDevelopmentToolStripMenuItem.Size = New System.Drawing.Size(195, 22)
        Me.ProductsDevelopmentToolStripMenuItem.Text = "Products Development"
        '
        'ProductsDevelopmentToolStripMenuItem1
        '
        Me.ProductsDevelopmentToolStripMenuItem1.Name = "ProductsDevelopmentToolStripMenuItem1"
        Me.ProductsDevelopmentToolStripMenuItem1.Size = New System.Drawing.Size(195, 22)
        Me.ProductsDevelopmentToolStripMenuItem1.Text = "Products Development"
        '
        'LoginToolStripMenuItem
        '
        Me.LoginToolStripMenuItem.Name = "LoginToolStripMenuItem"
        Me.LoginToolStripMenuItem.Size = New System.Drawing.Size(49, 20)
        Me.LoginToolStripMenuItem.Text = "Login"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "CTP INFORMATION SYSTEM"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents PurchasingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SupplierClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CustomerClaimsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductsDevelopmentToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProductsDevelopmentToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents LoginToolStripMenuItem As ToolStripMenuItem
End Class
