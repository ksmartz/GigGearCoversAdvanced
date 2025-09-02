Public Class Manufacturer
    Public Property ManufacturerId As Integer
    Public Property Name As String ' Unique
    Public Property SeriesList As New List(Of Series)
End Class
