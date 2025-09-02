Public Class Model
    Public Property ModelId As Integer
    Public Property SeriesId As Integer
    Public Property Name As String ' Unique per Series
    Public Property Width As Decimal
    Public Property Depth As Decimal
    Public Property Height As Decimal
    Public Property DesignFeatures As New List(Of DesignFeature)

End Class

