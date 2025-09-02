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
        Me.txtFabricBrandName = New System.Windows.Forms.TextBox()
        Me.txtFabricBrandProductName = New System.Windows.Forms.TextBox()
        Me.lblSupplier = New System.Windows.Forms.Label()
        Me.lblProductBrand = New System.Windows.Forms.Label()
        Me.lblProductName = New System.Windows.Forms.Label()
        Me.lblFabricType = New System.Windows.Forms.Label()
        Me.lblAddBrandName = New System.Windows.Forms.Label()
        Me.lblAddEditProductName = New System.Windows.Forms.Label()
        Me.lblShippingCost = New System.Windows.Forms.Label()
        Me.lblCostPerLinearYard = New System.Windows.Forms.Label()
        Me.lblWeightPerLinearYard = New System.Windows.Forms.Label()
        Me.lblSquareInchesPerLinearYard = New System.Windows.Forms.Label()
        Me.lblFabricRollWidth = New System.Windows.Forms.Label()
        Me.lblTotalYards = New System.Windows.Forms.Label()
        Me.lblCostPerSquareInch = New System.Windows.Forms.Label()
        Me.lblWeightPerSquareInch = New System.Windows.Forms.Label()
        Me.lbAddFabricType = New System.Windows.Forms.Label()
        Me.txtAddFabricType = New System.Windows.Forms.TextBox()
        Me.btnAddFabricType = New System.Windows.Forms.Button()
        Me.chkIsActiveForMarketPlace = New System.Windows.Forms.CheckBox()
        Me.txtFabricTypeNameAbbreviation = New System.Windows.Forms.TextBox()
        Me.lblFabricTypeNameAbbreviation = New System.Windows.Forms.Label()
        Me.chkAssignToSupplier = New System.Windows.Forms.CheckBox()
        Me.dgvAssignFabrics = New System.Windows.Forms.DataGridView()
        CType(Me.dgvAssignFabrics, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbSupplier
        '
        Me.cmbSupplier.FormattingEnabled = True
        Me.cmbSupplier.Location = New System.Drawing.Point(12, 51)
        Me.cmbSupplier.Name = "cmbSupplier"
        Me.cmbSupplier.Size = New System.Drawing.Size(169, 28)
        Me.cmbSupplier.TabIndex = 0
        '
        'cmbBrand
        '
        Me.cmbBrand.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBrand.FormattingEnabled = True
        Me.cmbBrand.Location = New System.Drawing.Point(12, 116)
        Me.cmbBrand.Name = "cmbBrand"
        Me.cmbBrand.Size = New System.Drawing.Size(169, 28)
        Me.cmbBrand.TabIndex = 1
        '
        'cmbFabricType
        '
        Me.cmbFabricType.FormattingEnabled = True
        Me.cmbFabricType.Location = New System.Drawing.Point(12, 246)
        Me.cmbFabricType.Name = "cmbFabricType"
        Me.cmbFabricType.Size = New System.Drawing.Size(169, 28)
        Me.cmbFabricType.TabIndex = 2
        '
        'cmbProduct
        '
        Me.cmbProduct.FormattingEnabled = True
        Me.cmbProduct.Location = New System.Drawing.Point(12, 182)
        Me.cmbProduct.Name = "cmbProduct"
        Me.cmbProduct.Size = New System.Drawing.Size(169, 28)
        Me.cmbProduct.TabIndex = 3
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(12, 293)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(203, 59)
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtShippingCost
        '
        Me.txtShippingCost.Location = New System.Drawing.Point(437, 180)
        Me.txtShippingCost.Name = "txtShippingCost"
        Me.txtShippingCost.Size = New System.Drawing.Size(150, 26)
        Me.txtShippingCost.TabIndex = 5
        '
        'txtCostPerLinearYard
        '
        Me.txtCostPerLinearYard.Location = New System.Drawing.Point(437, 246)
        Me.txtCostPerLinearYard.Name = "txtCostPerLinearYard"
        Me.txtCostPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtCostPerLinearYard.TabIndex = 6
        '
        'txtWeightPerLinearYard
        '
        Me.txtWeightPerLinearYard.Location = New System.Drawing.Point(437, 321)
        Me.txtWeightPerLinearYard.Name = "txtWeightPerLinearYard"
        Me.txtWeightPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtWeightPerLinearYard.TabIndex = 7
        '
        'txtSquareInchesPerLinearYard
        '
        Me.txtSquareInchesPerLinearYard.Location = New System.Drawing.Point(680, 53)
        Me.txtSquareInchesPerLinearYard.Name = "txtSquareInchesPerLinearYard"
        Me.txtSquareInchesPerLinearYard.Size = New System.Drawing.Size(150, 26)
        Me.txtSquareInchesPerLinearYard.TabIndex = 8
        '
        'txtFabricRollWidth
        '
        Me.txtFabricRollWidth.Location = New System.Drawing.Point(680, 111)
        Me.txtFabricRollWidth.Name = "txtFabricRollWidth"
        Me.txtFabricRollWidth.Size = New System.Drawing.Size(150, 26)
        Me.txtFabricRollWidth.TabIndex = 9
        '
        'txtWeightPerSquareInch
        '
        Me.txtWeightPerSquareInch.Location = New System.Drawing.Point(680, 324)
        Me.txtWeightPerSquareInch.Name = "txtWeightPerSquareInch"
        Me.txtWeightPerSquareInch.Size = New System.Drawing.Size(150, 26)
        Me.txtWeightPerSquareInch.TabIndex = 10
        '
        'txtTotalYards
        '
        Me.txtTotalYards.Location = New System.Drawing.Point(680, 180)
        Me.txtTotalYards.Name = "txtTotalYards"
        Me.txtTotalYards.Size = New System.Drawing.Size(150, 26)
        Me.txtTotalYards.TabIndex = 11
        '
        'txtCostPerSquareInch
        '
        Me.txtCostPerSquareInch.Location = New System.Drawing.Point(680, 246)
        Me.txtCostPerSquareInch.Name = "txtCostPerSquareInch"
        Me.txtCostPerSquareInch.Size = New System.Drawing.Size(150, 26)
        Me.txtCostPerSquareInch.TabIndex = 12
        '
        'txtFabricBrandName
        '
        Me.txtFabricBrandName.Location = New System.Drawing.Point(437, 52)
        Me.txtFabricBrandName.Name = "txtFabricBrandName"
        Me.txtFabricBrandName.Size = New System.Drawing.Size(203, 26)
        Me.txtFabricBrandName.TabIndex = 13
        '
        'txtFabricBrandProductName
        '
        Me.txtFabricBrandProductName.Location = New System.Drawing.Point(437, 111)
        Me.txtFabricBrandProductName.Name = "txtFabricBrandProductName"
        Me.txtFabricBrandProductName.Size = New System.Drawing.Size(203, 26)
        Me.txtFabricBrandProductName.TabIndex = 14
        '
        'lblSupplier
        '
        Me.lblSupplier.AutoSize = True
        Me.lblSupplier.Location = New System.Drawing.Point(12, 24)
        Me.lblSupplier.Name = "lblSupplier"
        Me.lblSupplier.Size = New System.Drawing.Size(67, 20)
        Me.lblSupplier.TabIndex = 15
        Me.lblSupplier.Text = "Supplier"
        '
        'lblProductBrand
        '
        Me.lblProductBrand.AutoSize = True
        Me.lblProductBrand.Location = New System.Drawing.Point(12, 88)
        Me.lblProductBrand.Name = "lblProductBrand"
        Me.lblProductBrand.Size = New System.Drawing.Size(52, 20)
        Me.lblProductBrand.TabIndex = 16
        Me.lblProductBrand.Text = "Brand"
        '
        'lblProductName
        '
        Me.lblProductName.AutoSize = True
        Me.lblProductName.Location = New System.Drawing.Point(12, 156)
        Me.lblProductName.Name = "lblProductName"
        Me.lblProductName.Size = New System.Drawing.Size(110, 20)
        Me.lblProductName.TabIndex = 17
        Me.lblProductName.Text = "Product Name"
        '
        'lblFabricType
        '
        Me.lblFabricType.AutoSize = True
        Me.lblFabricType.Location = New System.Drawing.Point(12, 223)
        Me.lblFabricType.Name = "lblFabricType"
        Me.lblFabricType.Size = New System.Drawing.Size(91, 20)
        Me.lblFabricType.TabIndex = 18
        Me.lblFabricType.Text = "Fabric Type"
        '
        'lblAddBrandName
        '
        Me.lblAddBrandName.AutoSize = True
        Me.lblAddBrandName.Location = New System.Drawing.Point(433, 24)
        Me.lblAddBrandName.Name = "lblAddBrandName"
        Me.lblAddBrandName.Size = New System.Drawing.Size(163, 20)
        Me.lblAddBrandName.TabIndex = 19
        Me.lblAddBrandName.Text = "Add/Edit Brand Name"
        '
        'lblAddEditProductName
        '
        Me.lblAddEditProductName.AutoSize = True
        Me.lblAddEditProductName.Location = New System.Drawing.Point(433, 88)
        Me.lblAddEditProductName.Name = "lblAddEditProductName"
        Me.lblAddEditProductName.Size = New System.Drawing.Size(175, 20)
        Me.lblAddEditProductName.TabIndex = 20
        Me.lblAddEditProductName.Text = "Add/Edit Product Name"
        '
        'lblShippingCost
        '
        Me.lblShippingCost.AutoSize = True
        Me.lblShippingCost.Location = New System.Drawing.Point(433, 152)
        Me.lblShippingCost.Name = "lblShippingCost"
        Me.lblShippingCost.Size = New System.Drawing.Size(108, 20)
        Me.lblShippingCost.TabIndex = 21
        Me.lblShippingCost.Text = "Shipping Cost"
        '
        'lblCostPerLinearYard
        '
        Me.lblCostPerLinearYard.AutoSize = True
        Me.lblCostPerLinearYard.Location = New System.Drawing.Point(433, 219)
        Me.lblCostPerLinearYard.Name = "lblCostPerLinearYard"
        Me.lblCostPerLinearYard.Size = New System.Drawing.Size(136, 20)
        Me.lblCostPerLinearYard.TabIndex = 22
        Me.lblCostPerLinearYard.Text = "Cost / Linear Yard"
        '
        'lblWeightPerLinearYard
        '
        Me.lblWeightPerLinearYard.AutoSize = True
        Me.lblWeightPerLinearYard.Location = New System.Drawing.Point(433, 289)
        Me.lblWeightPerLinearYard.Name = "lblWeightPerLinearYard"
        Me.lblWeightPerLinearYard.Size = New System.Drawing.Size(145, 20)
        Me.lblWeightPerLinearYard.TabIndex = 23
        Me.lblWeightPerLinearYard.Text = "Weight/Linear Yard"
        '
        'lblSquareInchesPerLinearYard
        '
        Me.lblSquareInchesPerLinearYard.AutoSize = True
        Me.lblSquareInchesPerLinearYard.Location = New System.Drawing.Point(676, 24)
        Me.lblSquareInchesPerLinearYard.Name = "lblSquareInchesPerLinearYard"
        Me.lblSquareInchesPerLinearYard.Size = New System.Drawing.Size(171, 20)
        Me.lblSquareInchesPerLinearYard.TabIndex = 24
        Me.lblSquareInchesPerLinearYard.Text = "Sq. Inches/Linear Yard"
        '
        'lblFabricRollWidth
        '
        Me.lblFabricRollWidth.AutoSize = True
        Me.lblFabricRollWidth.Location = New System.Drawing.Point(676, 88)
        Me.lblFabricRollWidth.Name = "lblFabricRollWidth"
        Me.lblFabricRollWidth.Size = New System.Drawing.Size(129, 20)
        Me.lblFabricRollWidth.TabIndex = 25
        Me.lblFabricRollWidth.Text = "Fabric Roll Width"
        '
        'lblTotalYards
        '
        Me.lblTotalYards.AutoSize = True
        Me.lblTotalYards.Location = New System.Drawing.Point(676, 152)
        Me.lblTotalYards.Name = "lblTotalYards"
        Me.lblTotalYards.Size = New System.Drawing.Size(101, 20)
        Me.lblTotalYards.TabIndex = 26
        Me.lblTotalYards.Text = "lblTotalYards"
        '
        'lblCostPerSquareInch
        '
        Me.lblCostPerSquareInch.AutoSize = True
        Me.lblCostPerSquareInch.Location = New System.Drawing.Point(676, 219)
        Me.lblCostPerSquareInch.Name = "lblCostPerSquareInch"
        Me.lblCostPerSquareInch.Size = New System.Drawing.Size(133, 20)
        Me.lblCostPerSquareInch.TabIndex = 27
        Me.lblCostPerSquareInch.Text = "Cost/Square Inch"
        '
        'lblWeightPerSquareInch
        '
        Me.lblWeightPerSquareInch.AutoSize = True
        Me.lblWeightPerSquareInch.Location = New System.Drawing.Point(676, 289)
        Me.lblWeightPerSquareInch.Name = "lblWeightPerSquareInch"
        Me.lblWeightPerSquareInch.Size = New System.Drawing.Size(146, 20)
        Me.lblWeightPerSquareInch.TabIndex = 28
        Me.lblWeightPerSquareInch.Text = "Weight/SquareInch"
        '
        'lbAddFabricType
        '
        Me.lbAddFabricType.AutoSize = True
        Me.lbAddFabricType.Location = New System.Drawing.Point(994, 36)
        Me.lbAddFabricType.Name = "lbAddFabricType"
        Me.lbAddFabricType.Size = New System.Drawing.Size(124, 20)
        Me.lbAddFabricType.TabIndex = 30
        Me.lbAddFabricType.Text = "Add Fabric Type"
        '
        'txtAddFabricType
        '
        Me.txtAddFabricType.Location = New System.Drawing.Point(968, 59)
        Me.txtAddFabricType.Name = "txtAddFabricType"
        Me.txtAddFabricType.Size = New System.Drawing.Size(150, 26)
        Me.txtAddFabricType.TabIndex = 29
        '
        'btnAddFabricType
        '
        Me.btnAddFabricType.Location = New System.Drawing.Point(915, 211)
        Me.btnAddFabricType.Name = "btnAddFabricType"
        Me.btnAddFabricType.Size = New System.Drawing.Size(203, 59)
        Me.btnAddFabricType.TabIndex = 31
        Me.btnAddFabricType.Text = "Add Fabric Type"
        Me.btnAddFabricType.UseVisualStyleBackColor = True
        '
        'chkIsActiveForMarketPlace
        '
        Me.chkIsActiveForMarketPlace.AutoSize = True
        Me.chkIsActiveForMarketPlace.Location = New System.Drawing.Point(206, 103)
        Me.chkIsActiveForMarketPlace.Name = "chkIsActiveForMarketPlace"
        Me.chkIsActiveForMarketPlace.Size = New System.Drawing.Size(192, 24)
        Me.chkIsActiveForMarketPlace.TabIndex = 32
        Me.chkIsActiveForMarketPlace.Text = "Active for Marketplace"
        Me.chkIsActiveForMarketPlace.UseVisualStyleBackColor = True
        '
        'txtFabricTypeNameAbbreviation
        '
        Me.txtFabricTypeNameAbbreviation.Location = New System.Drawing.Point(968, 152)
        Me.txtFabricTypeNameAbbreviation.Name = "txtFabricTypeNameAbbreviation"
        Me.txtFabricTypeNameAbbreviation.Size = New System.Drawing.Size(150, 26)
        Me.txtFabricTypeNameAbbreviation.TabIndex = 33
        '
        'lblFabricTypeNameAbbreviation
        '
        Me.lblFabricTypeNameAbbreviation.AutoSize = True
        Me.lblFabricTypeNameAbbreviation.Location = New System.Drawing.Point(978, 107)
        Me.lblFabricTypeNameAbbreviation.Name = "lblFabricTypeNameAbbreviation"
        Me.lblFabricTypeNameAbbreviation.Size = New System.Drawing.Size(140, 20)
        Me.lblFabricTypeNameAbbreviation.TabIndex = 34
        Me.lblFabricTypeNameAbbreviation.Text = "Fabric Type ABBR"
        '
        'chkAssignToSupplier
        '
        Me.chkAssignToSupplier.AutoSize = True
        Me.chkAssignToSupplier.Location = New System.Drawing.Point(206, 51)
        Me.chkAssignToSupplier.Name = "chkAssignToSupplier"
        Me.chkAssignToSupplier.Size = New System.Drawing.Size(163, 24)
        Me.chkAssignToSupplier.TabIndex = 35
        Me.chkAssignToSupplier.Text = "Assign to Supplier"
        Me.chkAssignToSupplier.UseVisualStyleBackColor = True
        '
        'dgvAssignFabrics
        '
        Me.dgvAssignFabrics.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAssignFabrics.Location = New System.Drawing.Point(39, 492)
        Me.dgvAssignFabrics.Name = "dgvAssignFabrics"
        Me.dgvAssignFabrics.RowHeadersWidth = 62
        Me.dgvAssignFabrics.RowTemplate.Height = 28
        Me.dgvAssignFabrics.Size = New System.Drawing.Size(1167, 216)
        Me.dgvAssignFabrics.TabIndex = 36
        '
        'frmFabricEntryForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1287, 884)
        Me.Controls.Add(Me.dgvAssignFabrics)
        Me.Controls.Add(Me.chkAssignToSupplier)
        Me.Controls.Add(Me.lblFabricTypeNameAbbreviation)
        Me.Controls.Add(Me.txtFabricTypeNameAbbreviation)
        Me.Controls.Add(Me.chkIsActiveForMarketPlace)
        Me.Controls.Add(Me.btnAddFabricType)
        Me.Controls.Add(Me.lbAddFabricType)
        Me.Controls.Add(Me.txtAddFabricType)
        Me.Controls.Add(Me.lblWeightPerSquareInch)
        Me.Controls.Add(Me.lblCostPerSquareInch)
        Me.Controls.Add(Me.lblTotalYards)
        Me.Controls.Add(Me.lblFabricRollWidth)
        Me.Controls.Add(Me.lblSquareInchesPerLinearYard)
        Me.Controls.Add(Me.lblWeightPerLinearYard)
        Me.Controls.Add(Me.lblCostPerLinearYard)
        Me.Controls.Add(Me.lblShippingCost)
        Me.Controls.Add(Me.lblAddEditProductName)
        Me.Controls.Add(Me.lblAddBrandName)
        Me.Controls.Add(Me.lblFabricType)
        Me.Controls.Add(Me.lblProductName)
        Me.Controls.Add(Me.lblProductBrand)
        Me.Controls.Add(Me.lblSupplier)
        Me.Controls.Add(Me.txtFabricBrandProductName)
        Me.Controls.Add(Me.txtFabricBrandName)
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
        CType(Me.dgvAssignFabrics, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents txtFabricBrandName As Windows.Forms.TextBox
    Friend WithEvents txtFabricBrandProductName As Windows.Forms.TextBox
    Friend WithEvents lblSupplier As Windows.Forms.Label
    Friend WithEvents lblProductBrand As Windows.Forms.Label
    Friend WithEvents lblProductName As Windows.Forms.Label
    Friend WithEvents lblFabricType As Windows.Forms.Label
    Friend WithEvents lblAddBrandName As Windows.Forms.Label
    Friend WithEvents lblAddEditProductName As Windows.Forms.Label
    Friend WithEvents lblShippingCost As Windows.Forms.Label
    Friend WithEvents lblCostPerLinearYard As Windows.Forms.Label
    Friend WithEvents lblWeightPerLinearYard As Windows.Forms.Label
    Friend WithEvents lblSquareInchesPerLinearYard As Windows.Forms.Label
    Friend WithEvents lblFabricRollWidth As Windows.Forms.Label
    Friend WithEvents lblTotalYards As Windows.Forms.Label
    Friend WithEvents lblCostPerSquareInch As Windows.Forms.Label
    Friend WithEvents lblWeightPerSquareInch As Windows.Forms.Label
    Friend WithEvents lbAddFabricType As Windows.Forms.Label
    Friend WithEvents txtAddFabricType As Windows.Forms.TextBox
    Friend WithEvents btnAddFabricType As Windows.Forms.Button
    Friend WithEvents chkIsActiveForMarketPlace As Windows.Forms.CheckBox
    Friend WithEvents txtFabricTypeNameAbbreviation As Windows.Forms.TextBox
    Friend WithEvents lblFabricTypeNameAbbreviation As Windows.Forms.Label
    Friend WithEvents chkAssignToSupplier As Windows.Forms.CheckBox
    Friend WithEvents dgvAssignFabrics As Windows.Forms.DataGridView
End Class
