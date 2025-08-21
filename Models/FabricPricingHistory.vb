Public Class FabricPricingHistory
    Public Property FabricPricingHistoryId As Integer
    Public Property FabricId As Integer
    Public Property DateFrom As Date
    Public Property DateTo As Nullable(Of Date)
    Public Property ShippingCost As Decimal
    Public Property CostPerLinearYard As Decimal
    Public Property WeightPerLinearYard As Decimal
    Public Property SquareInchesPerLinearYard As Decimal
    Public Property FabricRollWidth As Decimal
    Public Property TotalYards As Decimal
    Public Property CostPerSquareInch As Decimal
    Public Property WeightPerSquareInch As Decimal
    Public Overridable Property Fabric As Fabric
End Class