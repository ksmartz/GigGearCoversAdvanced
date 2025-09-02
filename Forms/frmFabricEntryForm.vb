Imports System.Data.SqlClient
Imports System.Windows.Forms

Partial Public Class frmFabricEntryForm
    Inherits Form

    Private isFormLoading As Boolean = False
    Private suppressProductSelectionEvent As Boolean = False

    Private Sub frmFabricEntryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isFormLoading = True
        Dim db As New DbConnectionManager()
        Dim suppliers = db.GetAllSuppliers() ' Returns DataTable or List(Of SupplierInformation)
        cmbSupplier.DataSource = suppliers
        cmbSupplier.DisplayMember = "SupplierName" ' Adjust to your actual column/property name
        cmbSupplier.ValueMember = "SupplierID"     ' Adjust to your actual column/property name

        InitializeAssignFabricsGrid()
        isFormLoading = False
    End Sub
    '' Populate Supplier ComboBox ONLY
    'Using conn = DbConnectionManager.GetConnection()
    '    Using cmd = conn.CreateCommand()
    '        cmd.CommandText = "SELECT PK_SupplierNameId, CompanyName FROM SupplierInformation"
    '        Using reader = cmd.ExecuteReader()
    '            Dim suppliers As New List(Of SupplierInformation)
    '            While reader.Read()
    '                suppliers.Add(New SupplierInformation With {
    '                .PK_SupplierNameId = reader.GetInt32(0),
    '                .CompanyName = reader.GetString(1)
    '            })
    '            End While
    '            cmbSupplier.DataSource = suppliers
    '            cmbSupplier.DisplayMember = "CompanyName"
    '            cmbSupplier.ValueMember = "PK_SupplierNameId"
    '        End Using
    '    End Using
    'End Using

    'cmbSupplier.SelectedIndex = -1

    '' Load all brands (not filtered)
    'LoadAllBrands()
    'cmbProduct.DataSource = Nothing

    'LoadFabricTypeCombo(-1)
    'cmbFabricType.SelectedIndex = -1

    '' Clear all textboxes
    'txtFabricBrandName.Clear()
    'txtFabricBrandProductName.Clear()
    'txtWeightPerLinearYard.Clear()
    'txtFabricRollWidth.Clear()
    'txtSquareInchesPerLinearYard.Clear()
    'txtWeightPerSquareInch.Clear()
    'txtTotalYards.Clear()
    'txtShippingCost.Clear()
    'txtCostPerLinearYard.Clear()
    'txtCostPerSquareInch.Clear()
    Private Sub InitializeAssignFabricsGrid()
        dgvAssignFabrics.Columns.Clear()

        ' BrandName dropdown
        Dim db As New DbConnectionManager()
        Dim brandNames = db.GetAllFabricBrandNames() ' Returns DataTable or List(Of FabricBrandName)
        Dim brandCol As New DataGridViewComboBoxColumn With {
        .Name = "BrandName",
        .HeaderText = "Brand Name",
        .DataSource = brandNames,
        .DisplayMember = "BrandName",
        .ValueMember = "BrandName"
    }
        dgvAssignFabrics.Columns.Add(brandCol)

        ' ProductName dropdown (will be filled dynamically)
        Dim productCol As New DataGridViewComboBoxColumn With {
        .Name = "ProductName",
        .HeaderText = "Product Name"
    }
        dgvAssignFabrics.Columns.Add(productCol)

        ' FabricType dropdown
        Dim fabricTypes = db.GetAllFabricTypes() ' Returns DataTable or List(Of FabricTypeName)
        Dim typeCol As New DataGridViewComboBoxColumn With {
        .Name = "FabricType",
        .HeaderText = "Fabric Type",
        .DataSource = fabricTypes,
        .DisplayMember = "FabricType",
        .ValueMember = "FabricType"
    }
        dgvAssignFabrics.Columns.Add(typeCol)

        ' Pricing columns
        dgvAssignFabrics.Columns.Add("ShippingCost", "Shipping Cost")
        dgvAssignFabrics.Columns.Add("CostPerLinearYard", "Cost Per Linear Yard")
        dgvAssignFabrics.Columns.Add("CostPerSquareInch", "Cost Per Square Inch")
        dgvAssignFabrics.Columns.Add("WeightPerSquareInch", "Weight Per Square Inch")
        dgvAssignFabrics.Columns.Add("WeightPerLinearYard", "Weight Per Linear Yard")
        dgvAssignFabrics.Columns.Add("FabricRollWidth", "Fabric Roll Width")
    End Sub
    Private Sub dgvAssignFabrics_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgvAssignFabrics.EditingControlShowing
        If dgvAssignFabrics.CurrentCell.ColumnIndex = dgvAssignFabrics.Columns("ProductName").Index Then
            Dim combo As ComboBox = TryCast(e.Control, ComboBox)
            If combo IsNot Nothing Then
                RemoveHandler combo.DropDown, AddressOf ProductNameDropDown
                AddHandler combo.DropDown, AddressOf ProductNameDropDown
            End If
        End If
    End Sub
    '*************************************************
    Private Sub dgvAssignFabrics_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAssignFabrics.CellValueChanged
        If e.RowIndex < 0 Then Return

        Dim row = dgvAssignFabrics.Rows(e.RowIndex)
        Dim supplierId = cmbSupplier.SelectedValue
        Dim brandName = row.Cells("BrandName").Value?.ToString()
        Dim productName = row.Cells("ProductName").Value?.ToString()

        If String.IsNullOrEmpty(brandName) OrElse String.IsNullOrEmpty(productName) Then Return

        Dim db As New DbConnectionManager()
        Dim pricing = db.GetFabricPricingHistory(supplierId, brandName, productName) ' Returns object or DataRow
        Dim productInfo = db.GetFabricProductInfo(brandName, productName) ' Returns object or DataRow

        ' Fill pricing columns if data exists, else leave blank
        row.Cells("ShippingCost").Value = If(pricing IsNot Nothing, pricing("ShippingCost"), "")
        row.Cells("CostPerLinearYard").Value = If(pricing IsNot Nothing, pricing("CostPerLinearYard"), "")
        row.Cells("CostPerSquareInch").Value = If(pricing IsNot Nothing, pricing("CostPerSquareInch"), "")
        row.Cells("WeightPerSquareInch").Value = If(pricing IsNot Nothing, pricing("WeightPerSquareInch"), "")

        row.Cells("WeightPerLinearYard").Value = If(productInfo IsNot Nothing, productInfo("WeightPerLinearYard"), "")
        row.Cells("FabricRollWidth").Value = If(productInfo IsNot Nothing, productInfo("FabricRollWidth"), "")
    End Sub

    Private Sub ProductNameDropDown(sender As Object, e As EventArgs)
        Dim rowIndex = dgvAssignFabrics.CurrentCell.RowIndex
        Dim brandName = dgvAssignFabrics.Rows(rowIndex).Cells("BrandName").Value?.ToString()
        If String.IsNullOrEmpty(brandName) Then Return

        Dim db As New DbConnectionManager()
        Dim products = db.GetProductsByBrandName(brandName) ' Returns DataTable or List(Of FabricBrandProductName)
        Dim combo As ComboBox = CType(dgvAssignFabrics.EditingControl, ComboBox)
        combo.DataSource = products
        combo.DisplayMember = "ProductName"
        combo.ValueMember = "ProductName"
    End Sub

    Private Sub LoadAllBrands() '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim brands = db.GetAllFabricBrandNames() ' Should return a List(Of FabricBrandName) or DataTable
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub
    Private Sub LoadBrandsForSupplier(supplierId As Integer, Optional selectBrandId As Integer = -1)
        Dim brands As New List(Of FabricBrandName)
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT DISTINCT b.PK_FabricBrandNameId, b.BrandName " &
                              "FROM SupplierProductNameData s " &
                              "INNER JOIN FabricBrandProductName p ON s.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId " &
                              "INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId " &
                              "WHERE s.FK_SupplierNameId = @SupplierId"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        brands.Add(New FabricBrandName With {
                        .PK_FabricBrandNameId = reader.GetInt32(0),
                        .BrandName = reader.GetString(1)
                    })
                    End While
                End Using
            End Using
        End Using
        cmbBrand.DataSource = brands
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        If selectBrandId <> -1 Then
            cmbBrand.SelectedValue = selectBrandId
        ElseIf brands.Count > 0 Then
            cmbBrand.SelectedIndex = 0
        End If
        cmbBrand.Refresh()
    End Sub
    Private Sub LoadAllProductsForBrand(brandId As Integer) '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim products = db.GetProductsByBrandId(brandId) ' You may need to add this method if not present
        cmbProduct.DataSource = products
        cmbProduct.DisplayMember = "BrandProductName"
        cmbProduct.ValueMember = "PK_FabricBrandProductNameId"
        cmbProduct.SelectedIndex = -1
    End Sub
    Private Sub LoadProductWeightAndWidth(productId As Integer) '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim productInfo = db.GetFabricProductInfoById(productId) ' You may need to add this method
        If productInfo IsNot Nothing Then
            txtWeightPerLinearYard.Text = productInfo("WeightPerLinearYard").ToString()
            txtFabricRollWidth.Text = productInfo("FabricRollWidth").ToString()
        Else
            txtWeightPerLinearYard.Clear()
            txtFabricRollWidth.Clear()
        End If
    End Sub

    Private Sub cmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSupplier.SelectedIndexChanged
        If isFormLoading Then Return

        Dim prevBrandId As Integer = -1
        Dim prevProductId As Integer = -1

        ' Remember the current brand and product selection
        If cmbBrand.SelectedValue IsNot Nothing AndAlso TypeOf cmbBrand.SelectedValue Is Integer Then
            prevBrandId = CInt(cmbBrand.SelectedValue)
        End If
        If cmbProduct.SelectedValue IsNot Nothing AndAlso TypeOf cmbProduct.SelectedValue Is Integer Then
            prevProductId = CInt(cmbProduct.SelectedValue)
        End If

        ' Clear brand and product selection
        cmbBrand.DataSource = Nothing
        cmbBrand.Items.Clear()
        cmbBrand.DisplayMember = Nothing
        cmbBrand.ValueMember = Nothing
        cmbProduct.DataSource = Nothing

        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ' Always reload all brands
        LoadAllBrands()

        If prevBrandId <> -1 Then
            cmbBrand.SelectedValue = prevBrandId
            ' Always reload all products for the brand and reselect previous
            LoadAllProductsForBrand(prevBrandId)
            If prevProductId <> -1 Then
                cmbProduct.SelectedValue = prevProductId

                ' Only load product-specific data if both were previously selected
                Using conn = DbConnectionManager.GetConnection()
                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "
                        SELECT b.BrandName, p.BrandProductName, p.WeightPerLinearYard, p.FabricRollWidth, p.FK_FabricTypeNameId
                        FROM FabricBrandProductName p
                        INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                        WHERE p.PK_FabricBrandProductNameId = @ProductId"
                        cmd.Parameters.AddWithValue("@ProductId", prevProductId)
                        Using reader = cmd.ExecuteReader()
                            If reader.Read() Then
                                txtFabricBrandName.Text = reader.GetString(0)
                                txtFabricBrandProductName.Text = reader.GetString(1)
                                If Not reader.IsDBNull(2) Then txtWeightPerLinearYard.Text = reader.GetDecimal(2).ToString() Else txtWeightPerLinearYard.Clear()
                                If Not reader.IsDBNull(3) Then
                                    txtFabricRollWidth.Text = reader.GetDecimal(3).ToString()
                                    txtSquareInchesPerLinearYard.Text = (reader.GetDecimal(3) * 36D).ToString()
                                Else
                                    txtFabricRollWidth.Clear()
                                    txtSquareInchesPerLinearYard.Clear()
                                End If
                                If Not reader.IsDBNull(4) Then
                                    cmbFabricType.SelectedValue = reader.GetInt32(4)
                                Else
                                    cmbFabricType.SelectedIndex = -1
                                End If
                                ' Calculate and fill weight per square inch
                                If Not reader.IsDBNull(2) AndAlso Not reader.IsDBNull(3) Then
                                    Dim weight = reader.GetDecimal(2)
                                    Dim rollWidth = reader.GetDecimal(3)
                                    Dim sqInches = rollWidth * 36D
                                    If sqInches > 0 Then
                                        txtWeightPerSquareInch.Text = Math.Round(weight / sqInches, 5).ToString()
                                    Else
                                        txtWeightPerSquareInch.Clear()
                                    End If
                                Else
                                    txtWeightPerSquareInch.Clear()
                                End If
                            End If
                        End Using
                    End Using
                End Using

                ' Now, if a supplier is selected, try to load supplier-specific data
                If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
                    Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
                    Using conn = DbConnectionManager.GetConnection()
                        Using cmd = conn.CreateCommand()
                            cmd.CommandText = "SELECT PK_SupplierProductNameDataId, SquareInchesPerLinearYard, TotalYards
                                   FROM SupplierProductNameData
                                   WHERE FK_SupplierNameId = @SupplierId AND FK_FabricBrandProductNameId = @ProductId"
                            cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                            cmd.Parameters.AddWithValue("@ProductId", prevProductId)
                            Using reader = cmd.ExecuteReader()
                                If reader.Read() Then
                                    If Not reader.IsDBNull(2) Then txtTotalYards.Text = reader.GetDecimal(2).ToString() Else txtTotalYards.Clear()
                                Else
                                    txtTotalYards.Clear()
                                    txtShippingCost.Clear()
                                    txtCostPerLinearYard.Clear()
                                    txtCostPerSquareInch.Clear()
                                End If
                            End Using
                        End Using
                    End Using
                End If
            Else
                ' No previous product, so clear product selection and textboxes
                cmbProduct.SelectedIndex = -1
                txtFabricBrandName.Clear()
                txtFabricBrandProductName.Clear()
                txtWeightPerLinearYard.Clear()
                txtFabricRollWidth.Clear()
                txtSquareInchesPerLinearYard.Clear()
                txtWeightPerSquareInch.Clear()
                txtTotalYards.Clear()
                txtShippingCost.Clear()
                txtCostPerLinearYard.Clear()
                txtCostPerSquareInch.Clear()
            End If
        Else
            ' No previous brand, so clear brand and product selection and textboxes
            cmbBrand.SelectedIndex = -1
            cmbProduct.SelectedIndex = -1
            txtFabricBrandName.Clear()
            txtFabricBrandProductName.Clear()
            txtWeightPerLinearYard.Clear()
            txtFabricRollWidth.Clear()
            txtSquareInchesPerLinearYard.Clear()
            txtWeightPerSquareInch.Clear()
            txtTotalYards.Clear()
            txtShippingCost.Clear()
            txtCostPerLinearYard.Clear()
            txtCostPerSquareInch.Clear()
        End If
    End Sub

    Private Sub cmbBrand_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBrand.SelectedIndexChanged
        If isFormLoading Then Return

        suppressProductSelectionEvent = True

        cmbProduct.DataSource = Nothing
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
            suppressProductSelectionEvent = False
            Return
        End If

        cmbBrand.SelectedIndex = 0

        If cmbBrand.SelectedValue Is Nothing OrElse Not TypeOf cmbBrand.SelectedValue Is Integer Then
            suppressProductSelectionEvent = False
            Return
        End If

        Dim brandId As Integer = CInt(cmbBrand.SelectedValue)

        If chkAssignToSupplier.Checked Then
            LoadAllProductsForBrand(brandId)
        Else
            Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
            ' Load only products for this supplier and brand
            Using conn = DbConnectionManager.GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "
                SELECT p.PK_FabricBrandProductNameId, p.BrandProductName
                FROM FabricBrandProductName p
                INNER JOIN SupplierProductNameData s ON s.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                WHERE s.FK_SupplierNameId = @SupplierId AND p.FK_FabricBrandNameId = @BrandId"
                    cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                    cmd.Parameters.AddWithValue("@BrandId", brandId)
                    Using reader = cmd.ExecuteReader()
                        Dim products As New List(Of FabricBrandProductName)
                        While reader.Read()
                            products.Add(New FabricBrandProductName With {
                            .PK_FabricBrandProductNameId = reader.GetInt32(0),
                            .BrandProductName = reader.GetString(1)
                        })
                        End While
                        cmbProduct.DataSource = products
                        cmbProduct.DisplayMember = "BrandProductName"
                        cmbProduct.ValueMember = "PK_FabricBrandProductNameId"
                    End Using
                End Using
            End Using
        End If
        cmbProduct.SelectedIndex = -1

        suppressProductSelectionEvent = False
    End Sub

    Private Sub cmbProduct_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProduct.SelectedIndexChanged '************REFACTOR TO DB MANAGER
        If suppressProductSelectionEvent Then Return

        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        If cmbProduct.SelectedValue Is Nothing OrElse Not TypeOf cmbProduct.SelectedValue Is Integer Then Return
        Dim productId As Integer = CInt(cmbProduct.SelectedValue)
        Dim db As New DbConnectionManager()

        If chkAssignToSupplier.Checked Then
            ' Assign mode: Only show product-specific data, clear supplier-specific fields
            LoadProductWeightAndWidth(productId)

            ' Fill in the brand and product name textboxes for editing
            Dim productInfo = db.GetFabricProductInfoById(productId)
            If productInfo IsNot Nothing Then
                txtFabricBrandName.Text = productInfo("BrandName").ToString()
                txtFabricBrandProductName.Text = productInfo("BrandProductName").ToString()
                If Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
                    cmbFabricType.SelectedValue = CInt(productInfo("FK_FabricTypeNameId"))
                Else
                    cmbFabricType.SelectedIndex = -1
                End If
            End If

            ' Clear supplier-specific fields ONLY
            txtTotalYards.Clear()
            txtShippingCost.Clear()
            txtCostPerLinearYard.Clear()
            txtCostPerSquareInch.Clear()
            ' DO NOT clear txtWeightPerSquareInch or txtSquareInchesPerLinearYard or txtFabricRollWidth
            Return
        End If

        ' Not in assign mode: Load supplier-specific data as before
        Dim supplierProductNameDataId As Integer = -1
        Dim fabricTypeId As Integer = -1
        Dim squareInchesPerLinearYard As Decimal = 0D
        Dim totalYards As Decimal = 0D

        If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
            Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)

            ' Use helper to get supplier-product data
            Dim supplierProduct = db.GetSupplierProductNameData(supplierId, productId)
            If supplierProduct IsNot Nothing Then
                If Not IsDBNull(supplierProduct("PK_SupplierProductNameDataId")) Then supplierProductNameDataId = CInt(supplierProduct("PK_SupplierProductNameDataId"))
                If Not IsDBNull(supplierProduct("FK_FabricTypeNameId")) Then fabricTypeId = CInt(supplierProduct("FK_FabricTypeNameId"))
                If Not IsDBNull(supplierProduct("SquareInchesPerLinearYard")) Then squareInchesPerLinearYard = CDec(supplierProduct("SquareInchesPerLinearYard"))
                If Not IsDBNull(supplierProduct("TotalYards")) Then totalYards = CDec(supplierProduct("TotalYards"))
                txtSquareInchesPerLinearYard.Text = squareInchesPerLinearYard.ToString()
                txtTotalYards.Text = totalYards.ToString()
                LoadFabricTypeCombo(fabricTypeId)
            End If

            ' Load weight and width from the product table
            Dim productInfo = db.GetFabricProductInfoById(productId)
            If productInfo IsNot Nothing Then
                txtWeightPerLinearYard.Text = productInfo("WeightPerLinearYard").ToString()
                txtFabricRollWidth.Text = productInfo("FabricRollWidth").ToString()
                txtFabricBrandName.Text = productInfo("BrandName").ToString()
                txtFabricBrandProductName.Text = productInfo("BrandProductName").ToString()
                If Not IsDBNull(productInfo("FK_FabricTypeNameId")) Then
                    cmbFabricType.SelectedValue = CInt(productInfo("FK_FabricTypeNameId"))
                End If
                ' Calculate and fill weight per square inch
                Dim weight As Decimal = 0D
                Dim rollWidth As Decimal = 0D
                If Not IsDBNull(productInfo("WeightPerLinearYard")) Then weight = CDec(productInfo("WeightPerLinearYard"))
                If Not IsDBNull(productInfo("FabricRollWidth")) Then rollWidth = CDec(productInfo("FabricRollWidth"))
                Dim sqInches = rollWidth * 36D
                If sqInches > 0 Then
                    txtWeightPerSquareInch.Text = Math.Round(weight / sqInches, 5).ToString()
                Else
                    txtWeightPerSquareInch.Clear()
                End If
            End If

            ' Load latest pricing from FabricPricingHistory using helper
            Dim pricing = db.GetFabricPricingHistoryByProductId(supplierId, productId)
            If pricing IsNot Nothing Then
                txtShippingCost.Text = pricing("ShippingCost").ToString()
                txtCostPerLinearYard.Text = pricing("CostPerLinearYard").ToString()
                txtCostPerSquareInch.Text = pricing("CostPerSquareInch").ToString()
                txtWeightPerSquareInch.Text = pricing("WeightPerSquareInch").ToString()
            End If
        End If
    End Sub

    Private Sub LoadFabricTypeCombo(selectedId As Integer) '************REFACTOR TO DB MANAGER
        Dim db As New DbConnectionManager()
        Dim types = db.GetAllFabricTypes()
        cmbFabricType.DataSource = types
        cmbFabricType.DisplayMember = "FabricType"
        cmbFabricType.ValueMember = "PK_FabricTypeNameId"
        cmbFabricType.SelectedValue = selectedId
    End Sub



    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim isActive As Boolean = chkIsActiveForMarketPlace.Checked
            Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
            Dim fabricTypeId As Integer = CInt(cmbFabricType.SelectedValue)

            ' Use the textboxes for new brand and product names
            Dim brandName As String = txtFabricBrandName.Text.Trim()
            If String.IsNullOrWhiteSpace(brandName) Then
                MessageBox.Show("Please enter a brand name.")
                Return
            End If

            Dim productName As String = txtFabricBrandProductName.Text.Trim()
            If String.IsNullOrWhiteSpace(productName) Then
                MessageBox.Show("Please enter a product name.")
                Return
            End If

            Dim brandId As Integer = -1

            ' Check if the brand exists, otherwise insert it
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then conn.Open()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT PK_FabricBrandNameId FROM FabricBrandName WHERE LOWER(BrandName) = @BrandName"
                    cmd.Parameters.AddWithValue("@BrandName", brandName.ToLower())
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        brandId = CInt(result)
                    Else
                        cmd.CommandText = "INSERT INTO FabricBrandName (BrandName) VALUES (@BrandName); SELECT CAST(SCOPE_IDENTITY() AS int);"
                        cmd.Parameters.Clear()
                        cmd.Parameters.AddWithValue("@BrandName", brandName)
                        brandId = CInt(cmd.ExecuteScalar())
                    End If
                End Using
            End Using

            ' Ensure the product exists for the selected brand, or insert it and get the ID
            Dim productId As Integer
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then conn.Open()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT PK_FabricBrandProductNameId FROM FabricBrandProductName WHERE BrandProductName = @ProductName AND FK_FabricBrandNameId = @BrandId"
                    cmd.Parameters.AddWithValue("@ProductName", productName)
                    cmd.Parameters.AddWithValue("@BrandId", brandId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        productId = CInt(result)
                    Else
                        cmd.CommandText = "INSERT INTO FabricBrandProductName (BrandProductName, FK_FabricBrandNameId) VALUES (@ProductName, @BrandId); SELECT CAST(SCOPE_IDENTITY() AS int);"
                        cmd.Parameters.Clear()
                        cmd.Parameters.AddWithValue("@ProductName", productName)
                        cmd.Parameters.AddWithValue("@BrandId", brandId)
                        productId = CInt(cmd.ExecuteScalar())
                    End If
                End Using
            End Using

            Dim weight As Decimal = Decimal.Parse(txtWeightPerLinearYard.Text)
            Dim rollWidth As Decimal = Decimal.Parse(txtFabricRollWidth.Text)
            Dim totalYards As Decimal = Decimal.Parse(txtTotalYards.Text)
            Dim shippingCost As Decimal = Decimal.Parse(txtShippingCost.Text)
            Dim costPerLinearYard As Decimal = Decimal.Parse(txtCostPerLinearYard.Text)

            ' Calculate square inches per linear yard
            Dim squareInchesPerLinearYard As Decimal = rollWidth * 36D

            ' Calculate cost per square inch (rounded to 5 decimal places)
            Dim costPerSquareInch As Decimal = 0D
            If squareInchesPerLinearYard > 0 Then
                costPerSquareInch = Math.Round(costPerLinearYard / squareInchesPerLinearYard, 5)
            End If

            ' Calculate weight per square inch (rounded to 5 decimal places)
            Dim weightPerSquareInch As Decimal = 0D
            If squareInchesPerLinearYard > 0 Then
                weightPerSquareInch = Math.Round(weight / squareInchesPerLinearYard, 5)
            End If

            ' Unset IsActiveForMarketplace for all other records with this product if needed
            If isActive Then
                Using conn = DbConnectionManager.GetConnection()
                    If conn.State <> ConnectionState.Open Then
                        conn.Open()
                    End If
                    Using cmd = conn.CreateCommand()
                        cmd.CommandText = "UPDATE SupplierProductNameData SET IsActiveForMarketplace = 0 WHERE FK_FabricBrandProductNameId = @ProductId"
                        cmd.Parameters.AddWithValue("@ProductId", productId)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
            End If

            ' Insert into SupplierProductNameData and get the new ID
            Dim newSupplierProductId As Integer
            newSupplierProductId = InsertSupplierProductNameData(supplierId, productId, fabricTypeId, weight, squareInchesPerLinearYard, rollWidth, totalYards, isActive)

            ' Insert into FabricPricingHistory with Quantity = totalYards
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then
                    conn.Open()
                End If
                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
"INSERT INTO FabricPricingHistory " &
"(FK_SupplierProductNameDataId, DateFrom, ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch, Quantity) " &
"VALUES (@SupplierProductNameDataId, @DateFrom, @ShippingCost, @CostPerLinearYard, @CostPerSquareInch, @WeightPerSquareInch, @Quantity)"
                    cmd.Parameters.AddWithValue("@SupplierProductNameDataId", newSupplierProductId)
                    cmd.Parameters.AddWithValue("@DateFrom", Date.Now)
                    cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                    cmd.Parameters.AddWithValue("@CostPerLinearYard", costPerLinearYard)
                    cmd.Parameters.AddWithValue("@CostPerSquareInch", costPerSquareInch)
                    cmd.Parameters.AddWithValue("@WeightPerSquareInch", weightPerSquareInch)
                    cmd.Parameters.AddWithValue("@Quantity", totalYards)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Saved!")
            If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
                LoadBrandsForSupplier(CInt(cmbSupplier.SelectedValue), brandId)
            End If
        Catch ex As Exception
            MessageBox.Show("Error saving data: " & ex.Message)
        End Try
    End Sub

    Private Sub btnAddFabricType_Click(sender As Object, e As EventArgs) Handles btnAddFabricType.Click
        Dim newTypeName As String = txtAddFabricType.Text.Trim()
        Dim newAbbreviation As String = txtFabricTypeNameAbbreviation.Text.Trim()

        If String.IsNullOrWhiteSpace(newTypeName) Then
            MessageBox.Show("Please enter a fabric type name.")
            Return
        End If

        ' Check if the fabric type already exists (case-insensitive)
        Dim exists As Boolean = False
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM FabricTypeName WHERE LOWER(FabricType) = @TypeName"
                cmd.Parameters.AddWithValue("@TypeName", newTypeName.ToLower())
                exists = CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using

        If exists Then
            MessageBox.Show("This fabric type already exists.")
            Return
        End If

        ' Insert the new fabric type
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO FabricTypeName (FabricType, FabricTypeNameAbbreviation) VALUES (@TypeName, @Abbreviation)"
                cmd.Parameters.AddWithValue("@TypeName", newTypeName)
                cmd.Parameters.AddWithValue("@Abbreviation", newAbbreviation)
                cmd.ExecuteNonQuery()
            End Using
        End Using

        txtAddFabricType.Clear()
        txtFabricTypeNameAbbreviation.Clear()
        MessageBox.Show("Fabric type added.")
    End Sub

    Private Sub chkAssignToSupplier_CheckedChanged(sender As Object, e As EventArgs) Handles chkAssignToSupplier.CheckedChanged
        ' Only clear supplier-specific fields
        txtTotalYards.Clear()
        txtShippingCost.Clear()
        txtCostPerLinearYard.Clear()
        txtCostPerSquareInch.Clear()
        ' DO NOT clear or reset cmbProduct, cmbFabricType, or any product-specific fields here!
    End Sub
    Private Sub ClearTextBoxes()
        txtTotalYards.Clear()
        txtShippingCost.Clear()
        txtCostPerLinearYard.Clear()
        txtCostPerSquareInch.Clear()
        ' Do NOT clear product-specific fields here!
    End Sub



    Private Function InsertSupplierProductNameData(
    supplierId As Integer,
    productId As Integer,
    fabricTypeId As Integer,
    weight As Decimal,
    squareInchesPerLinearYard As Decimal,
    rollWidth As Decimal,
    totalYards As Decimal,
    isActive As Boolean
) As Integer
        Dim newSupplierProductId As Integer
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                    "INSERT INTO SupplierProductNameData " &
                    "(FK_SupplierNameId, FK_FabricBrandProductNameId, FK_FabricTypeNameId, WeightPerLinearYard, SquareInchesPerLinearYard, FabricRollWidth, TotalYards, IsActiveForMarketplace) " &
                    "VALUES (@SupplierId, @ProductId, @FabricTypeId, @Weight, @SquareInches, @RollWidth, @TotalYards, @IsActive); " &
                    "SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                cmd.Parameters.AddWithValue("@FabricTypeId", fabricTypeId)
                cmd.Parameters.AddWithValue("@Weight", weight)
                cmd.Parameters.AddWithValue("@SquareInches", squareInchesPerLinearYard)
                cmd.Parameters.AddWithValue("@RollWidth", rollWidth)
                cmd.Parameters.AddWithValue("@TotalYards", totalYards)
                cmd.Parameters.AddWithValue("@IsActive", isActive)
                newSupplierProductId = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return newSupplierProductId
    End Function

    Private Sub LoadAllBrandsWithSupplierMark(supplierId As Integer)
        Dim allBrands As New List(Of BrandDisplayItem)
        Dim supplierBrandIds As New HashSet(Of Integer)

        ' Get all brands associated with the supplier
        If supplierId > 0 Then
            Using conn = DbConnectionManager.GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "SELECT DISTINCT b.PK_FabricBrandNameId
                                   FROM SupplierProductNameData s
                                   INNER JOIN FabricBrandProductName p ON s.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE s.FK_SupplierNameId = @SupplierId"
                    cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            supplierBrandIds.Add(reader.GetInt32(0))
                        End While
                    End Using
                End Using
            End Using
        End If

        ' Get all brands
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandNameId, BrandName FROM FabricBrandName"
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim id = reader.GetInt32(0)
                        allBrands.Add(New BrandDisplayItem With {
                        .PK_FabricBrandNameId = id,
                        .BrandName = reader.GetString(1),
                        .IsSupplierBrand = supplierBrandIds.Contains(id)
                    })
                    End While
                End Using
            End Using
        End Using

        cmbBrand.DataSource = allBrands
        cmbBrand.DisplayMember = "DisplayText"
        cmbBrand.ValueMember = "PK_FabricBrandNameId"
        cmbBrand.SelectedIndex = -1
    End Sub

    Private Sub lblAddEditProductName_Click(sender As Object, e As EventArgs) Handles lblAddEditProductName.Click

    End Sub

    Private Sub txtShippingCost_TextChanged(sender As Object, e As EventArgs) Handles txtShippingCost.TextChanged

    End Sub

    Private Sub txtAddFabricType_TextChanged(sender As Object, e As EventArgs) Handles txtAddFabricType.TextChanged

    End Sub
End Class

Public Class ProductDisplayItem
    Public Property PK_FabricBrandProductNameId As Integer
    Public Property BrandProductName As String
    Public Property IsSupplierProduct As Boolean
    Public ReadOnly Property DisplayText As String
        Get
            If IsSupplierProduct Then
                Return "★ " & BrandProductName
            Else
                Return BrandProductName
            End If
        End Get
    End Property
End Class