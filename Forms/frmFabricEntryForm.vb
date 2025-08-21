Imports System.Data.Entity
Imports System.Windows.Forms

Partial Public Class frmFabricEntryForm
    Inherits Form

    Private db As New FabricDbContext()

    Private Sub frmFabricEntryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Populate ComboBoxes with data from the database
        cmbSupplier.DataSource = db.Suppliers.ToList()
        cmbSupplier.DisplayMember = "Name"
        cmbSupplier.ValueMember = "SupplierId"

        cmbBrand.DataSource = db.FabricBrands.ToList()
        cmbBrand.DisplayMember = "BrandName"
        cmbBrand.ValueMember = "FabricBrandId"

        cmbProduct.DataSource = db.FabricProducts.ToList()
        cmbProduct.DisplayMember = "ProductName"
        cmbProduct.ValueMember = "FabricProductId"

        cmbFabricType.DataSource = db.FabricTypes.ToList()
        cmbFabricType.DisplayMember = "TypeName"
        cmbFabricType.ValueMember = "FabricTypeId"
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            ' Create and save the Fabric entity
            Dim fabric As New Fabric With {
                .SupplierId = CInt(cmbSupplier.SelectedValue),
                .FabricProductId = CInt(cmbProduct.SelectedValue),
                .FabricTypeId = CInt(cmbFabricType.SelectedValue)
            }
            db.Fabrics.Add(fabric)
            db.SaveChanges()

            ' Create and save the FabricPricingHistory entity
            Dim pricing As New FabricPricingHistory With {
                .FabricId = fabric.FabricId,
                .DateFrom = Date.Now,
                .ShippingCost = Decimal.Parse(txtShippingCost.Text),
                .CostPerLinearYard = Decimal.Parse(txtCostPerLinearYard.Text),
                .WeightPerLinearYard = Decimal.Parse(txtWeightPerLinearYard.Text),
                .SquareInchesPerLinearYard = Decimal.Parse(txtSquareInchesPerLinearYard.Text),
                .FabricRollWidth = Decimal.Parse(txtFabricRollWidth.Text),
                .TotalYards = Decimal.Parse(txtTotalYards.Text),
                .CostPerSquareInch = Decimal.Parse(txtCostPerSquareInch.Text),
                .WeightPerSquareInch = Decimal.Parse(txtWeightPerSquareInch.Text)
            }
            db.FabricPricingHistories.Add(pricing)
            db.SaveChanges()

            MessageBox.Show("Saved!")
        Catch ex As Exception
            MessageBox.Show("Error saving data: " & ex.Message)
        End Try
    End Sub
End Class