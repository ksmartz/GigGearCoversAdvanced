Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Public Class DbConnectionManager
    ' Adjust your connection string as needed
    Private Shared ReadOnly Property ConnectionString As String
        Get
            Return "Data Source=MYPC\SQLEXPRESS;Initial Catalog=GigGearCoversDb;Integrated Security=True"
        End Get
    End Property

    Public Shared Function GetConnection() As SqlConnection
        Return New SqlConnection(ConnectionString)
    End Function

    ' 1. Get all suppliers
    Public Function GetAllSuppliers() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SupplierNameId AS SupplierID, CompanyName AS SupplierName FROM SupplierInformation"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 2. Get all fabric brand names
    Public Function GetAllFabricBrandNames() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandNameId, BrandName FROM FabricBrandName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 3. Get all fabric types
    Public Function GetAllFabricTypes() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricTypeNameId, FabricType FROM FabricTypeName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 4. Get products by brand name (for DataGridView dropdown)
    Public Function GetProductsByBrandName(brandName As String) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.PK_FabricBrandProductNameId, p.BrandProductName AS ProductName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE b.BrandName = @BrandName"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 5. Get products by brand ID (for ComboBox)
    Public Function GetProductsByBrandId(brandId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_FabricBrandProductNameId, BrandProductName
                                   FROM FabricBrandProductName
                                   WHERE FK_FabricBrandNameId = @BrandId"
                cmd.Parameters.AddWithValue("@BrandId", brandId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 6. Get product info by brand and product name (for DataGridView)
    Public Function GetFabricProductInfo(brandName As String, productName As String) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.*, b.BrandName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE b.BrandName = @BrandName AND p.BrandProductName = @ProductName"
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 7. Get product info by product ID (for ComboBox and details)
    Public Function GetFabricProductInfoById(productId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT p.*, b.BrandName
                                   FROM FabricBrandProductName p
                                   INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                                   WHERE p.PK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 8. Get latest pricing history by supplier, brand, and product name (for DataGridView)
    Public Function GetFabricPricingHistory(supplierId As Integer, brandName As String, productName As String) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                    SELECT TOP 1 h.*
                    FROM FabricPricingHistory h
                    INNER JOIN SupplierProductNameData s ON h.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                    INNER JOIN FabricBrandProductName p ON s.FK_FabricBrandProductNameId = p.PK_FabricBrandProductNameId
                    INNER JOIN FabricBrandName b ON p.FK_FabricBrandNameId = b.PK_FabricBrandNameId
                    WHERE s.FK_SupplierNameId = @SupplierId AND b.BrandName = @BrandName AND p.BrandProductName = @ProductName
                    ORDER BY h.DateFrom DESC"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@BrandName", brandName)
                cmd.Parameters.AddWithValue("@ProductName", productName)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 9. Get latest pricing history by supplier and product ID (for ComboBox/details)
    Public Function GetFabricPricingHistoryByProductId(supplierId As Integer, productId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                    SELECT TOP 1 h.*
                    FROM FabricPricingHistory h
                    INNER JOIN SupplierProductNameData s ON h.FK_SupplierProductNameDataId = s.PK_SupplierProductNameDataId
                    WHERE s.FK_SupplierNameId = @SupplierId AND s.FK_FabricBrandProductNameId = @ProductId
                    ORDER BY h.DateFrom DESC"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function

    ' 10. Get supplier-product data (for ComboBox/details)
    Public Function GetSupplierProductNameData(supplierId As Integer, productId As Integer) As DataRow
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                    SELECT *
                    FROM SupplierProductNameData
                    WHERE FK_SupplierNameId = @SupplierId AND FK_FabricBrandProductNameId = @ProductId"
                cmd.Parameters.AddWithValue("@SupplierId", supplierId)
                cmd.Parameters.AddWithValue("@ProductId", productId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0) Else Return Nothing
                End Using
            End Using
        End Using
    End Function
    ' Add these methods to your DbConnectionManager class

    ' 1. Get all equipment types
    Public Function GetEquipmentTypes() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_EquipmentTypeId, EquipmentTypeName FROM EquipmentType"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 2. Get all manufacturer names
    Public Function GetManufacturerNames() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_ManufacturerId, ManufacturerName FROM Manufacturer ORDER BY ManufacturerName"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    ' Set the primary key for easier DataRow lookup by ID
                    If dt.Columns.Contains("PK_ManufacturerId") Then
                        dt.PrimaryKey = New DataColumn() {dt.Columns("PK_ManufacturerId")}
                    End If
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 3. Insert a new manufacturer
    Public Function InsertManufacturer(manufacturerName As String) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO Manufacturer (ManufacturerName) VALUES (@Name); SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@Name", manufacturerName)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' 4. Get all series names for a manufacturer
    Public Function GetSeriesNamesByManufacturer(manufacturerId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_SeriesId, SeriesName FROM Series WHERE FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 5. Insert a new series
    Public Function InsertSeries(seriesName As String, manufacturerId As Integer, equipmentTypeId As Integer) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO Series (SeriesName, FK_ManufacturerId, FK_EquipmentTypeId) VALUES (@Name, @ManufacturerId, @EquipmentTypeId); SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@Name", seriesName)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' 6. Get equipment type ID by manufacturer and series
    Public Function GetEquipmentTypeIdByManufacturerAndSeries(manufacturerId As Integer, seriesId As Integer) As Integer?
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT FK_EquipmentTypeId FROM Series WHERE PK_SeriesId = @SeriesId AND FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                conn.Open() ' <-- REQUIRED!
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return CInt(result)
                Else
                    Return Nothing
                End If
            End Using
        End Using
    End Function

    ' 7. Get models by manufacturer and series
    Public Function GetModelsByManufacturerAndSeries(manufacturerId As Integer, seriesId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT m.*
                FROM Model m
                INNER JOIN Series s ON m.FK_SeriesId = s.PK_SeriesId
                WHERE m.FK_SeriesId = @SeriesId AND s.FK_ManufacturerId = @ManufacturerId"
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    ' 8. Check if a model exists
    Public Function ModelExists(manufacturerId As Integer, seriesId As Integer, modelName As String) As Boolean
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM Model WHERE FK_ManufacturerId = @ManufacturerId AND FK_SeriesId = @SeriesId AND ModelName = @ModelName"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                Return CInt(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function

    ' 9. Insert a new model
    Public Function InsertModel(
    manufacturerId As Integer,
    seriesId As Integer,
    modelName As String,
    width As Decimal,
    depth As Decimal,
    height As Decimal,
    OptionalHeight As Object,
    OptionalDepth As Object,
    AngleType As Object,
    AmpHandleLocation As Object,
    TAHWidth As Object,
    TAHHeight As Object,
    SAHWidth As Object,
    SAHHeight As Object,
    TAHRearOffset As Object,
    SAHRearOffset As Object,
    SAHTopDownOffset As Object,
    MusicRestDesign As Object,
    Chart_Template As Object,
    Notes As Object
) As Integer
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "INSERT INTO Model (FK_ManufacturerId, FK_SeriesId, ModelName, Width, Depth, Height, OptionalHeight, OptionalDepth, AngleType, AmpHandleLocation, TAHWidth, TAHHeight, SAHWidth, SAHHeight, TAHRearOffset, SAHRearOffset, SAHTopDownOffset, MusicRestDesign, Chart_Template, Notes) " &
                                  "VALUES (@ManufacturerId, @SeriesId, @ModelName, @Width, @Depth, @Height, @OptionalHeight, @OptionalDepth, @AngleType, @AmpHandleLocation, @TAHWidth, @TAHHeight, @SAHWidth, @SAHHeight, @TAHRearOffset, @SAHRearOffset, @SAHTopDownOffset, @MusicRestDesign, @Chart_Template, @Notes); " &
                                  "SELECT CAST(SCOPE_IDENTITY() AS int);"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                cmd.Parameters.AddWithValue("@Width", width)
                cmd.Parameters.AddWithValue("@Depth", depth)
                cmd.Parameters.AddWithValue("@Height", height)
                cmd.Parameters.AddWithValue("@OptionalHeight", If(OptionalHeight Is Nothing, DBNull.Value, OptionalHeight))
                cmd.Parameters.AddWithValue("@OptionalDepth", If(OptionalDepth Is Nothing, DBNull.Value, OptionalDepth))
                cmd.Parameters.AddWithValue("@AngleType", If(AngleType Is Nothing, DBNull.Value, AngleType))
                cmd.Parameters.AddWithValue("@AmpHandleLocation", If(AmpHandleLocation Is Nothing, DBNull.Value, AmpHandleLocation))
                cmd.Parameters.AddWithValue("@TAHWidth", If(TAHWidth Is Nothing, DBNull.Value, TAHWidth))
                cmd.Parameters.AddWithValue("@TAHHeight", If(TAHHeight Is Nothing, DBNull.Value, TAHHeight))
                cmd.Parameters.AddWithValue("@SAHWidth", If(SAHWidth Is Nothing, DBNull.Value, SAHWidth))
                cmd.Parameters.AddWithValue("@SAHHeight", If(SAHHeight Is Nothing, DBNull.Value, SAHHeight))
                cmd.Parameters.AddWithValue("@TAHRearOffset", If(TAHRearOffset Is Nothing, DBNull.Value, TAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHRearOffset", If(SAHRearOffset Is Nothing, DBNull.Value, SAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHTopDownOffset", If(SAHTopDownOffset Is Nothing, DBNull.Value, SAHTopDownOffset))
                cmd.Parameters.AddWithValue("@MusicRestDesign", If(MusicRestDesign Is Nothing, DBNull.Value, MusicRestDesign))
                cmd.Parameters.AddWithValue("@Chart_Template", If(Chart_Template Is Nothing, DBNull.Value, Chart_Template))
                cmd.Parameters.AddWithValue("@Notes", If(Notes Is Nothing, DBNull.Value, Notes))
                Return CInt(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ' 10. Update an existing model
    Public Function UpdateModel(
    modelId As Integer,
    modelName As String,
    width As Decimal,
    depth As Decimal,
    height As Decimal,
    OptionalHeight As Object,
    OptionalDepth As Object,
    AngleType As Object,
    AmpHandleLocation As Object,
    TAHWidth As Object,
    TAHHeight As Object,
    SAHWidth As Object,
    SAHHeight As Object,
    TAHRearOffset As Object,
    SAHRearOffset As Object,
    SAHTopDownOffset As Object,
    MusicRestDesign As Object,
    Chart_Template As Object,
    Notes As Object
) As Boolean
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "UPDATE Model SET ModelName=@ModelName, Width=@Width, Depth=@Depth, Height=@Height, OptionalHeight=@OptionalHeight, OptionalDepth=@OptionalDepth, AngleType=@AngleType, AmpHandleLocation=@AmpHandleLocation, TAHWidth=@TAHWidth, TAHHeight=@TAHHeight, SAHWidth=@SAHWidth, SAHHeight=@SAHHeight, TAHRearOffset=@TAHRearOffset, SAHRearOffset=@SAHRearOffset, SAHTopDownOffset=@SAHTopDownOffset, MusicRestDesign=@MusicRestDesign, Chart_Template=@Chart_Template, Notes=@Notes WHERE PK_ModelId=@ModelId"
                cmd.Parameters.AddWithValue("@ModelId", modelId)
                cmd.Parameters.AddWithValue("@ModelName", modelName)
                cmd.Parameters.AddWithValue("@Width", width)
                cmd.Parameters.AddWithValue("@Depth", depth)
                cmd.Parameters.AddWithValue("@Height", height)
                cmd.Parameters.AddWithValue("@OptionalHeight", If(OptionalHeight Is Nothing, DBNull.Value, OptionalHeight))
                cmd.Parameters.AddWithValue("@OptionalDepth", If(OptionalDepth Is Nothing, DBNull.Value, OptionalDepth))
                cmd.Parameters.AddWithValue("@AngleType", If(AngleType Is Nothing, DBNull.Value, AngleType))
                cmd.Parameters.AddWithValue("@AmpHandleLocation", If(AmpHandleLocation Is Nothing, DBNull.Value, AmpHandleLocation))
                cmd.Parameters.AddWithValue("@TAHWidth", If(TAHWidth Is Nothing, DBNull.Value, TAHWidth))
                cmd.Parameters.AddWithValue("@TAHHeight", If(TAHHeight Is Nothing, DBNull.Value, TAHHeight))
                cmd.Parameters.AddWithValue("@SAHWidth", If(SAHWidth Is Nothing, DBNull.Value, SAHWidth))
                cmd.Parameters.AddWithValue("@SAHHeight", If(SAHHeight Is Nothing, DBNull.Value, SAHHeight))
                cmd.Parameters.AddWithValue("@TAHRearOffset", If(TAHRearOffset Is Nothing, DBNull.Value, TAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHRearOffset", If(SAHRearOffset Is Nothing, DBNull.Value, SAHRearOffset))
                cmd.Parameters.AddWithValue("@SAHTopDownOffset", If(SAHTopDownOffset Is Nothing, DBNull.Value, SAHTopDownOffset))
                cmd.Parameters.AddWithValue("@MusicRestDesign", If(MusicRestDesign Is Nothing, DBNull.Value, MusicRestDesign))
                cmd.Parameters.AddWithValue("@Chart_Template", If(Chart_Template Is Nothing, DBNull.Value, Chart_Template))
                cmd.Parameters.AddWithValue("@Notes", If(Notes Is Nothing, DBNull.Value, Notes))
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    Public Function UpdateSeriesEquipmentType(seriesId As Integer, equipmentTypeId As Integer) As Boolean
        Using conn = GetConnection()
            conn.Open()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "UPDATE Series SET FK_EquipmentTypeId = @EquipmentTypeId WHERE PK_SeriesId = @SeriesId"
                cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    Public Function GetSeriesAndEquipmentTypeByManufacturer(manufacturerId As Integer) As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                SELECT 
                    s.PK_SeriesId, 
                    s.SeriesName, 
                    e.EquipmentTypeName
                FROM 
                    Series s
                    INNER JOIN EquipmentType e ON s.FK_EquipmentTypeId = e.PK_EquipmentTypeId
                WHERE 
                    s.FK_ManufacturerId = @ManufacturerId
                ORDER BY s.SeriesName"
                cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function GetAllAngleTypesTable() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_AngleTypeId, AngleTypeName FROM AngleType"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function

    Public Function GetAllAmpHandleLocationsTable() As DataTable
        Using conn = GetConnection()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "SELECT PK_AmpHandleLocationId, AmpHandleLocationName FROM AmpHandleLocation"
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)
                    Return dt
                End Using
            End Using
        End Using
    End Function
End Class