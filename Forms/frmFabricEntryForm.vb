Imports System.Data.SqlClient
Imports System.Windows.Forms

Partial Public Class frmFabricEntryForm
    Inherits Form

    Private isFormLoading As Boolean = False

    Private Sub frmFabricEntryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isFormLoading = True

        ' Populate Supplier ComboBox
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SupplierNameId, CompanyName FROM SupplierInformation"
                Using reader = cmd.ExecuteReader()
                    Dim suppliers As New List(Of SupplierInformation)
                    While reader.Read()
                        suppliers.Add(New SupplierInformation With {
                        .PK_SupplierNameId = reader.GetInt32(0),
                        .CompanyName = reader.GetString(1)
                    })
                    End While
                    cmbSupplier.DataSource = suppliers
                    cmbSupplier.DisplayMember = "CompanyName"
                    cmbSupplier.ValueMember = "PK_SupplierNameId"
                End Using
            End Using
        End Using

        cmbSupplier.SelectedIndex = -1

        ' Load all brands (not filtered)
        LoadAllBrands()

        cmbProduct.DataSource = Nothing

        ' Load fabric types into ComboBox on form load
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        isFormLoading = False
    End Sub
    Private Sub LoadAllBrands()
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandNameId, BrandName FROM FabricBrandName"
                Using reader = cmd.ExecuteReader()
                    Dim brands As New List(Of FabricBrandName)
                    While reader.Read()
                        brands.Add(New FabricBrandName With {
                        .PK_FabricBrandNameId = reader.GetInt32(0),
                        .BrandName = reader.GetString(1)
                    })
                    End While
                    cmbBrand.DataSource = brands
                    cmbBrand.DisplayMember = "BrandName"
                    cmbBrand.ValueMember = "PK_FabricBrandNameId"
                End Using
            End Using
        End Using
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
    Private Sub LoadAllProductsForBrand(brandId As Integer)
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandProductNameId, BrandProductName FROM FabricBrandProductName WHERE FK_FabricBrandNameId = @BrandId"
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
        cmbProduct.SelectedIndex = -1
    End Sub
    Private Sub LoadProductWeightAndWidth(productId As Integer)
        Using conn = DbConnectionManager.GetConnection()
            If conn.State <> ConnectionState.Open Then
                conn.Open()
            End If
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT WeightPerLinearYard, FabricRollWidth
                FROM FabricBrandProductName
                WHERE PK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        If Not reader.IsDBNull(0) Then
                            txtWeightPerLinearYard.Text = reader.GetDecimal(0).ToString()
                        Else
                            txtWeightPerLinearYard.Clear()
                        End If
                        If Not reader.IsDBNull(1) Then
                            txtFabricRollWidth.Text = reader.GetDecimal(1).ToString()
                        Else
                            txtFabricRollWidth.Clear()
                        End If
                    Else
                        txtWeightPerLinearYard.Clear()
                        txtFabricRollWidth.Clear()
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub cmbSupplier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSupplier.SelectedIndexChanged
        If isFormLoading Then Return

        cmbBrand.DataSource = Nothing
        cmbBrand.Items.Clear()
        cmbBrand.DisplayMember = Nothing
        cmbBrand.ValueMember = Nothing
        cmbProduct.DataSource = Nothing

        ' Always reload all fabric types
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        ' If no supplier is selected, show all brands
        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then
            LoadAllBrands()
            Return
        End If

        ' Load brands for the selected supplier
        Dim supplierId As Integer = CInt(cmbSupplier.SelectedValue)
        LoadBrandsForSupplier(supplierId)
    End Sub

    Private Sub cmbBrand_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBrand.SelectedIndexChanged
        cmbProduct.DataSource = Nothing
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        If cmbSupplier.SelectedValue Is Nothing OrElse Not TypeOf cmbSupplier.SelectedValue Is Integer Then Return

        cmbBrand.SelectedIndex = 0

        If cmbBrand.SelectedValue Is Nothing OrElse Not TypeOf cmbBrand.SelectedValue Is Integer Then Return
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
    End Sub

    Private Sub cmbProduct_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbProduct.SelectedIndexChanged
        LoadFabricTypeCombo(-1)
        cmbFabricType.SelectedIndex = -1

        ClearTextBoxes()

        If cmbProduct.SelectedValue Is Nothing OrElse Not TypeOf cmbProduct.SelectedValue Is Integer Then Return
        Dim productId As Integer = CInt(cmbProduct.SelectedValue)

        ' Load SupplierProductNameData for this product
        Dim supplierProductNameDataId As Integer = -1
        Dim fabricTypeId As Integer = -1

        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT TOP 1 PK_SupplierProductNameDataId, FK_FabricTypeNameId, WeightPerLinearYard, SquareInchesPerLinearYard, FabricRollWidth, TotalYards
                FROM SupplierProductNameData
                WHERE FK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        supplierProductNameDataId = reader.GetInt32(0)
                        fabricTypeId = reader.GetInt32(1)
                        txtSquareInchesPerLinearYard.Text = reader.GetDecimal(3).ToString()

                        txtTotalYards.Text = reader.GetDecimal(5).ToString()
                        LoadFabricTypeCombo(fabricTypeId)
                    End If
                End Using
            End Using
        End Using
        ' load weight and width from the product table:
        LoadProductWeightAndWidth(productId)
        ' Load latest pricing from FabricPricingHistory
        If supplierProductNameDataId <> -1 Then
            Using conn = DbConnectionManager.GetConnection()
                Using cmd = conn.CreateCommand()
                    cmd.CommandText = "
                    SELECT TOP 1 ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch
                    FROM FabricPricingHistory
                    WHERE FK_SupplierProductNameDataId = @SupplierProductNameDataId
                    ORDER BY DateFrom DESC"
                    cmd.Parameters.AddWithValue("@SupplierProductNameDataId", supplierProductNameDataId)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            txtShippingCost.Text = reader.GetDecimal(0).ToString()
                            txtCostPerLinearYard.Text = reader.GetDecimal(1).ToString()
                            txtCostPerSquareInch.Text = reader.GetDecimal(2).ToString()
                            txtWeightPerSquareInch.Text = reader.GetDecimal(3).ToString()
                        End If
                    End Using
                End Using
            End Using
        End If

        ' Fill in the brand and product name textboxes for editing
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT b.BrandName, p.BrandProductName
                FROM FabricBrandProductName p
                INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                WHERE p.PK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        txtFabricBrandName.Text = reader.GetString(0)
                        txtFabricBrandProductName.Text = reader.GetString(1)
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub LoadFabricTypeCombo(selectedId As Integer)
        Using conn = DbConnectionManager.GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricTypeNameId, FabricType FROM FabricTypeName"
                Using reader = cmd.ExecuteReader()
                    Dim types As New List(Of FabricTypeName)
                    While reader.Read()
                        types.Add(New FabricTypeName With {
                            .PK_FabricTypeNameId = reader.GetInt32(0),
                            .FabricType = reader.GetString(1)
                        })
                    End While
                    cmbFabricType.DataSource = types
                    cmbFabricType.DisplayMember = "FabricType"
                    cmbFabricType.ValueMember = "PK_FabricTypeNameId"
                End Using
            End Using
        End Using
        cmbFabricType.SelectedValue = selectedId
    End Sub

    Private Sub ClearTextBoxes()
        txtWeightPerLinearYard.Clear()
        txtFabricRollWidth.Clear()
        txtTotalYards.Clear()
        txtShippingCost.Clear()
        txtCostPerLinearYard.Clear()
        ' Add any other textboxes you want to clear
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

            ' Insert into FabricPricingHistory
            Using conn = DbConnectionManager.GetConnection()
                If conn.State <> ConnectionState.Open Then
                    conn.Open()
                End If
                Using cmd = conn.CreateCommand()
                    cmd.CommandText =
            "INSERT INTO FabricPricingHistory " &
            "(FK_SupplierProductNameDataId, DateFrom, ShippingCost, CostPerLinearYard, CostPerSquareInch, WeightPerSquareInch) " &
            "VALUES (@SupplierProductNameDataId, @DateFrom, @ShippingCost, @CostPerLinearYard, @CostPerSquareInch, @WeightPerSquareInch)"
                    cmd.Parameters.AddWithValue("@SupplierProductNameDataId", newSupplierProductId)
                    cmd.Parameters.AddWithValue("@DateFrom", Date.Now)
                    cmd.Parameters.AddWithValue("@ShippingCost", shippingCost)
                    cmd.Parameters.AddWithValue("@CostPerLinearYard", costPerLinearYard)
                    cmd.Parameters.AddWithValue("@CostPerSquareInch", costPerSquareInch)
                    cmd.Parameters.AddWithValue("@WeightPerSquareInch", weightPerSquareInch)
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
        If chkAssignToSupplier.Checked Then
            ' Show all brands
            LoadAllBrands()
            cmbProduct.DataSource = Nothing
        Else
            ' Show only brands for the selected supplier
            If cmbSupplier.SelectedValue IsNot Nothing AndAlso TypeOf cmbSupplier.SelectedValue Is Integer Then
                LoadBrandsForSupplier(CInt(cmbSupplier.SelectedValue))
            Else
                LoadAllBrands()
            End If
            cmbProduct.DataSource = Nothing
        End If
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
End Class