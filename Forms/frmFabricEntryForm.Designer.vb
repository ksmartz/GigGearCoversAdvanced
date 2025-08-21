<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFabricEntryForm
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
        Me.cmbSupplier = New System.Windows.Forms.ComboBox()
        Me.cmbBrand = New System.Windows.Forms.ComboBox()
        Me.cmbFabricType = New System.Windows.Forms.ComboBox()
        Me.cmbProduct = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtShippingCost = New System.Windows.Forms.TextBox()
        Me.txtCostPerLinearYard = New System.Windows.Forms.TextBox()
        Me.txtWeightPerLinearYard = New System.Windows.Forms.TextBox()
        Me.txtSquareInchesPerLinearYard = New System.Windows.Forms.TextBox()
        Me.txtFabricRollWidth = New System.Windows.Forms.TextBox()
        Me.txtWeightPerSquareInch = New System.Windows.Forms.TextBox()
        Me.txtTotalYards = New System.Windows.Forms.TextBox()
        Me.txtCostPerSquareInch = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'cmbSupplier
        '
        Me.cmbSupplier.FormattingEnabled = True
        Me.cmbSupplier.Location = New System.Drawing.Point(40, 32)
        Me.cmbSupplier.Name = "cmbSupplier"
        Me.cmbSupplier.Size = New System.Drawing.Size(169, 28)
        Me.cmbSupplier.TabIndex = 0
        '
        'cmbBrand
        '
        Me.cmbBrand.FormattingEnabled = True
        Me.cmbBrand.Location = New System.Drawing.Point(53, 114)
        Me.cmbBrand.Name = "cmbBrand"
        Me.cmbBrand.Size = New System.Drawing.Size(169, 28)
        Me.cmbBrand.TabIndex = 1
        '
        'cmbFabricType
        '
        Me.cmbFabricType.FormattingEnabled = True
        Me.cmbFabricType.Location = New System.Drawing.Point(70, 258)
        Me.cmbFabricType.Name = "cmbFabricType"
        Me.cmbFabricType.Size = New System.Drawing.Size(169, 28)
        Me.cmbFabricType.TabIndex = 2
        '
        'cmbProduct
        '
        Me.cmbProduct.FormattingEnabled = True
        Me.cmbProduct.Location = New System.Drawing.Point(53, 177)
        Me.cmbProduct.Name = "cmbProduct"
        Me.cmbProduct.Size = New System.Drawing.Size(169, 28)
        Me.cmbProduct.TabIndex = 3
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(104, 340)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(203, 59)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtShippingCost
        '
        Me.txtShippingCost.Location = New System.Drawing.Point(335, 19)
        Me.txtShippingCost.Name = "txtShippingCost"
        Me.txtShippingCost.Size = New System.Drawing.Size(150, 26)
        Me.txtShippingCost.TabIndex = 5
        '
        'txtCostPerLinearYard
        '
        Me.txtCostPerLinearYard.Location = New System.Drawing.Point(352, 85)
        Me.txtCostPerLinearYard.Name = "txtCostPerLinearYard"
        Me.txtCostPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtCostPerLinearYard.TabIndex = 6
        '
        'txtWeightPerLinearYard
        '
        Me.txtWeightPerLinearYard.Location = New System.Drawing.Point(352, 136)
        Me.txtWeightPerLinearYard.Name = "txtWeightPerLinearYard"
        Me.txtWeightPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtWeightPerLinearYard.TabIndex = 7
        '
        'txtSquareInchesPerLinearYard
        '
        Me.txtSquareInchesPerLinearYard.Location = New System.Drawing.Point(352, 191)
        Me.txtSquareInchesPerLinearYard.Name = "txtSquareInchesPerLinearYard"
        Me.txtSquareInchesPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtSquareInchesPerLinearYard.TabIndex = 8
        '
        'txtFabricRollWidth
        '
        Me.txtFabricRollWidth.Location = New System.Drawing.Point(369, 236)
        Me.txtFabricRollWidth.Name = "txtFabricRollWidth"
        Me.txtFabricRollWidth.Size = New System.Drawing.Size(150, 26)
        Me.txtFabricRollWidth.TabIndex = 9
        '
        'txtWeightPerSquareInch
        '
        Me.txtWeightPerSquareInch.Location = New System.Drawing.Point(382, 387)
        Me.txtWeightPerSquareInch.Name = "txtWeightPerSquareInch"
        Me.txtWeightPerSquareInch.Size = New System.Drawing.Size(150, 26)
        Me.txtWeightPerSquareInch.TabIndex = 10
        '
        'txtTotalYards
        '
        Me.txtTotalYards.Location = New System.Drawing.Point(369, 294)
        Me.txtTotalYards.Name = "txtTotalYards"
        Me.txtTotalYards.Size = New System.Drawing.Size(150, 26)
        Me.txtTotalYards.TabIndex = 11
        '
        'txtCostPerSquareInch
        '
        Me.txtCostPerSquareInch.Location = New System.Drawing.Point(382, 340)
        Me.txtCostPerSquareInch.Name = "txtCostPerSquareInch"
        Me.txtCostPerSquareInch.Size = New System.Drawing.Size(150, 26)
        Me.txtCostPerSquareInch.TabIndex = 12
        '
        'frmFabricEntryForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1169, 450)
        Me.Controls.Add(Me.txtCostPerSquareInch)
        Me.Controls.Add(Me.txtTotalYards)
        Me.Controls.Add(Me.txtWeightPerSquareInch)
        Me.Controls.Add(Me.txtFabricRollWidth)
        Me.Controls.Add(Me.txtSquareInchesPerLinearYard)
        Me.Controls.Add(Me.txtWeightPerLinearYard)
        Me.Controls.Add(Me.txtCostPerLinearYard)
        Me.Controls.Add(Me.txtShippingCost)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cmbProduct)
        Me.Controls.Add(Me.cmbFabricType)
        Me.Controls.Add(Me.cmbBrand)
        Me.Controls.Add(Me.cmbSupplier)
        Me.Name = "frmFabricEntryForm"
        Me.Text = "frmFabricEntryForm"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbSupplier As Windows.Forms.ComboBox
    Friend WithEvents cmbBrand As Windows.Forms.ComboBox
    Friend WithEvents cmbFabricType As Windows.Forms.ComboBox
    Friend WithEvents cmbProduct As Windows.Forms.ComboBox
    Friend WithEvents btnSave As Windows.Forms.Button
    Friend WithEvents txtShippingCost As Windows.Forms.TextBox
    Friend WithEvents txtCostPerLinearYard As Windows.Forms.TextBox
    Friend WithEvents txtWeightPerLinearYard As Windows.Forms.TextBox
    Friend WithEvents txtSquareInchesPerLinearYard As Windows.Forms.TextBox
    Friend WithEvents txtFabricRollWidth As Windows.Forms.TextBox
    Friend WithEvents txtWeightPerSquareInch As Windows.Forms.TextBox
    Friend WithEvents txtTotalYards As Windows.Forms.TextBox
    Friend WithEvents txtCostPerSquareInch As Windows.Forms.TextBox
End Class
