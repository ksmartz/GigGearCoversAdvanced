Imports System.Data.Entity

Public Class FabricDbContext
    Inherits DbContext

    Public Property Suppliers As DbSet(Of Supplier)
    Public Property FabricBrands As DbSet(Of FabricBrand)
    Public Property FabricProducts As DbSet(Of FabricProduct)
    Public Property FabricTypes As DbSet(Of FabricType)
    Public Property Fabrics As DbSet(Of Fabric)
    Public Property FabricPricingHistories As DbSet(Of FabricPricingHistory)
End Class