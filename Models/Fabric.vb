Public Class Fabric
    Public Property FabricId As Integer
    Public Property FabricProductId As Integer
    Public Property SupplierId As Integer
    Public Property FabricTypeId As Integer
    Public Overridable Property FabricProduct As FabricProduct
    Public Overridable Property Supplier As Supplier
    Public Overridable Property FabricType As FabricType
    Public Overridable Property PricingHistory As ICollection(Of FabricPricingHistory)
End Class