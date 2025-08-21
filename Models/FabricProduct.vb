Public Class FabricProduct
    Public Property FabricProductId As Integer
    Public Property ProductName As String
    Public Property FabricBrandId As Integer
    Public Overridable Property FabricBrand As FabricBrand
    Public Overridable Property Fabrics As ICollection(Of Fabric)
End Class