﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LoadingExcel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LoadingExcel))
        Me.ExcelLoader = New System.Windows.Forms.PictureBox()
        CType(Me.ExcelLoader, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ExcelLoader
        '
        Me.ExcelLoader.Image = CType(resources.GetObject("ExcelLoader.Image"), System.Drawing.Image)
        Me.ExcelLoader.Location = New System.Drawing.Point(12, 12)
        Me.ExcelLoader.Name = "ExcelLoader"
        Me.ExcelLoader.Size = New System.Drawing.Size(108, 82)
        Me.ExcelLoader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.ExcelLoader.TabIndex = 0
        Me.ExcelLoader.TabStop = False
        '
        'LoadingExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(132, 106)
        Me.ControlBox = False
        Me.Controls.Add(Me.ExcelLoader)
        Me.Name = "LoadingExcel"
        Me.Text = "Loading.."
        CType(Me.ExcelLoader, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ExcelLoader As PictureBox
End Class
