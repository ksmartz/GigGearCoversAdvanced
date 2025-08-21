Public Class FabricType
    Public Property FabricTypeId As Integer
    Public Property TypeName As String
    Public Overridable Property Fabrics As ICollection(Of Fabric)
End Class