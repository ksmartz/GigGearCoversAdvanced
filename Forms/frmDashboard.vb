Public Class frmDashboard
    Private Sub btnAddSupplier_Click(sender As Object, e As EventArgs) Handles btnAddSupplier.Click
        Dim supplierForm As New frmSuppliers()
        supplierForm.ShowDialog()
    End Sub

    Private Sub btnMaterialsDataEntry_Click(sender As Object, e As EventArgs) Handles btnMaterialsDataEntry.Click
        Dim fabricEntryForm As New frmFabricEntryForm()
        fabricEntryForm.ShowDialog() ' Use Show() if you want it non-modal
    End Sub
End Class