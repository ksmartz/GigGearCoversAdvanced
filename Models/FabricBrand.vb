Public Class FabricBrand
    Public Property FabricBrandId As Integer
    Public Property BrandName As String
    Public Overridable Property FabricProducts As ICollection(Of FabricProduct)
End Class