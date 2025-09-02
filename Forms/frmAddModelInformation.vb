Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms
Imports System.Linq

Public Class frmAddModelInformation
    Private isFormLoaded As Boolean = False
    Private isUserSelectingSeries As Boolean = False
    Private isLoadingSeries As Boolean = False
    Private isLoadingModels As Boolean = False
    Private originalRowValues As Dictionary(Of Integer, Dictionary(Of String, Object)) = New Dictionary(Of Integer, Dictionary(Of String, Object))()


    Private Sub AddModelInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim db As New DbConnectionManager()

        ' Manufacturer ComboBox: table-driven
        cmbManufacturerName.DropDownStyle = ComboBoxStyle.DropDownList
        cmbManufacturerName.MaxDropDownItems = 10
        cmbManufacturerName.DataSource = Nothing
        cmbManufacturerName.DataSource = db.GetManufacturerNames()
        cmbManufacturerName.DisplayMember = "ManufacturerName"
        cmbManufacturerName.ValueMember = "PK_ManufacturerId"

        ' Equipment Type ComboBox: table-driven ONLY
        cmbEquipmentType.DataSource = Nothing
        cmbEquipmentType.DataSource = db.GetEquipmentTypes()
        cmbEquipmentType.DisplayMember = "EquipmentTypeName"
        cmbEquipmentType.ValueMember = "PK_EquipmentTypeId"
        cmbEquipmentType.DropDownStyle = ComboBoxStyle.DropDownList

        AddHandler dgvSeriesName.MouseDown, AddressOf dgvSeriesName_MouseDown
        AddHandler dgvSeriesName.KeyDown, AddressOf dgvSeriesName_KeyDown
        dgvModelInformation.AutoGenerateColumns = False

        isFormLoaded = True
        ' InitializeModelGridColumns() ' Uncomment if you want to initialize columns on load
    End Sub
    Private Sub dgvSeriesName_MouseDown(sender As Object, e As MouseEventArgs)
        isUserSelectingSeries = True
    End Sub

    Private Sub dgvSeriesName_KeyDown(sender As Object, e As KeyEventArgs)
        isUserSelectingSeries = True
    End Sub
    Private Sub cmbManufacturerName_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbManufacturerName.KeyDown
        If e.KeyCode = Keys.Enter Then
            CheckAndAddManufacturer()
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub cmbManufacturerName_Leave(sender As Object, e As EventArgs) Handles cmbManufacturerName.Leave
        CheckAndAddManufacturer()
    End Sub

    Private Sub CheckAndAddManufacturer()
        Dim enteredName As String = cmbManufacturerName.Text.Trim()
        If String.IsNullOrEmpty(enteredName) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingNames = db.GetManufacturerNames()
        If Not existingNames.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("ManufacturerName"), enteredName, StringComparison.OrdinalIgnoreCase)) Then
            Dim result = MessageBox.Show(
            $"'{enteredName}' is not in the list. Do you want to add it as a new manufacturer?",
            "Add Manufacturer",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        )
            If result = DialogResult.Yes Then
                db.InsertManufacturer(enteredName)
                ' Refresh ComboBox
                cmbManufacturerName.DataSource = Nothing
                cmbManufacturerName.DataSource = db.GetManufacturerNames()
                cmbManufacturerName.SelectedItem = enteredName
            End If
        End If
    End Sub

    Private Sub cmbManufacturerName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManufacturerName.SelectedIndexChanged
        If Not isFormLoaded Then Return
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then Return

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then Return

        Dim db As New DbConnectionManager()
        Dim seriesList = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        isLoadingSeries = True
        cmbSeriesName.DataSource = Nothing
        cmbSeriesName.DataSource = seriesList
        cmbSeriesName.DisplayMember = "SeriesName"
        cmbSeriesName.ValueMember = "PK_SeriesId"
        cmbSeriesName.SelectedIndex = -1
        isLoadingSeries = False

        LoadModelsForSelectedSeries()
        LoadSeriesGridForManufacturer()
    End Sub

    Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
        If isLoadingSeries OrElse Not isFormLoaded Then Return

        SetEquipmentTypeForSelectedSeries()

        ' Get the selected equipment type name
        Dim equipmentTypeName As String = cmbEquipmentType.Text.Trim()

        ' Initialize columns for the selected equipment type
        InitializeModelGridColumns(equipmentTypeName)

        LoadModelsForSelectedSeries()

        If Not String.IsNullOrEmpty(cmbSeriesName.Text.Trim()) Then
            dgvModelInformation.Enabled = True
        Else
            dgvModelInformation.Enabled = False
        End If
    End Sub

    Private Sub cmbSeriesName_Leave(sender As Object, e As EventArgs) Handles cmbSeriesName.Leave
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then Exit Sub
        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then Exit Sub

        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            Dim equipmentType As String = cmbEquipmentType.Text.Trim()
            If String.IsNullOrEmpty(equipmentType) Then
                MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Exit Sub
            End If

            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                            $"Manufacturer ID: {selectedManufacturerId}" & vbCrLf &
                            $"Series: {enteredSeries}" & vbCrLf &
                            $"Equipment Type: {equipmentType}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(enteredSeries, selectedManufacturerId, equipmentTypeValue)
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturerId)
                cmbSeriesName.DisplayMember = "SeriesName"
                cmbSeriesName.ValueMember = "PK_SeriesId"
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub

    Private Sub SetEquipmentTypeForSelectedSeries()
        If isLoadingSeries Then Return
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If
        Dim selectedSeriesId As Integer
        If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
            cmbEquipmentType.SelectedIndex = -1
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim equipmentTypeId = db.GetEquipmentTypeIdByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)

        If equipmentTypeId.HasValue Then
            cmbEquipmentType.SelectedValue = equipmentTypeId.Value
        Else
            cmbEquipmentType.SelectedIndex = -1
        End If
    End Sub

    Private Sub cmbEquipmentType_Leave(sender As Object, e As EventArgs) Handles cmbEquipmentType.Leave
        TrySaveNewSeries()
    End Sub

    Private Sub LoadModelsForSelectedSeries()
        isLoadingModels = True
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If
        Dim selectedSeriesId As Integer
        If Not Integer.TryParse(cmbSeriesName.SelectedValue.ToString(), selectedSeriesId) Then
            dgvModelInformation.DataSource = Nothing
            isLoadingModels = False
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturerId, selectedSeriesId)
        dgvModelInformation.DataSource = dt
        isLoadingModels = False
    End Sub

    'Private Sub dgvModelInformation_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvModelInformation.RowValidating
    '    Dim row = dgvModelInformation.Rows(e.RowIndex)
    '    If isLoadingModels Then Return
    '    ' Only skip if the row is the placeholder for a new row and is empty
    '    If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
    '        Return
    '    End If

    '    Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
    '    Dim selectedSeries As String = cmbSeriesName.Text.Trim()
    '    If String.IsNullOrEmpty(selectedSeries) Then
    '        MessageBox.Show("Please select a series before adding a model.", "Series Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '        e.Cancel = True
    '        Return
    '    End If

    '    dgvModelInformation.EndEdit()

    '    Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
    '    Dim width As Decimal = If(dgvModelInformation.Columns.Contains("Width"), Convert.ToDecimal(row.Cells("Width").Value), 0D)
    '    Dim depth As Decimal = If(dgvModelInformation.Columns.Contains("Depth"), Convert.ToDecimal(row.Cells("Depth").Value), 0D)
    '    Dim height As Decimal = If(dgvModelInformation.Columns.Contains("Height"), Convert.ToDecimal(row.Cells("Height").Value), 0D)
    '    Dim optionalHeight As Decimal? = Nothing
    '    Dim optionalDepth As Decimal? = Nothing
    '    Dim angleType As String = Nothing
    '    Dim ampHandleLocation As String = Nothing
    '    Dim tahWidth As Decimal? = Nothing
    '    Dim tahHeight As Decimal? = Nothing
    '    Dim sahWidth As Decimal? = Nothing
    '    Dim sahHeight As Decimal? = Nothing
    '    Dim tahRearOffset As Decimal? = Nothing
    '    Dim sahRearOffset As Decimal? = Nothing
    '    Dim sahTopDownOffset As Decimal? = Nothing
    '    Dim musicRestDesign As Boolean? = Nothing
    '    Dim chartTemplate As Boolean? = Nothing
    '    Dim notes As String = Nothing

    '    If dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso row.Cells("OptionalHeight").Value IsNot Nothing AndAlso row.Cells("OptionalHeight").Value.ToString() <> "" Then
    '        optionalHeight = Convert.ToDecimal(row.Cells("OptionalHeight").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso row.Cells("OptionalDepth").Value IsNot Nothing AndAlso row.Cells("OptionalDepth").Value.ToString() <> "" Then
    '        optionalDepth = Convert.ToDecimal(row.Cells("OptionalDepth").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("AngleType") AndAlso row.Cells("AngleType").Value IsNot Nothing Then
    '        angleType = Convert.ToString(row.Cells("AngleType").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("AmpHandleLocation") AndAlso row.Cells("AmpHandleLocation").Value IsNot Nothing Then
    '        ampHandleLocation = Convert.ToString(row.Cells("AmpHandleLocation").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("TAHWidth") AndAlso row.Cells("TAHWidth").Value IsNot Nothing AndAlso row.Cells("TAHWidth").Value.ToString() <> "" Then
    '        tahWidth = Convert.ToDecimal(row.Cells("TAHWidth").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("TAHHeight") AndAlso row.Cells("TAHHeight").Value IsNot Nothing AndAlso row.Cells("TAHHeight").Value.ToString() <> "" Then
    '        tahHeight = Convert.ToDecimal(row.Cells("TAHHeight").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("SAHWidth") AndAlso row.Cells("SAHWidth").Value IsNot Nothing AndAlso row.Cells("SAHWidth").Value.ToString() <> "" Then
    '        sahWidth = Convert.ToDecimal(row.Cells("SAHWidth").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("SAHHeight") AndAlso row.Cells("SAHHeight").Value IsNot Nothing AndAlso row.Cells("SAHHeight").Value.ToString() <> "" Then
    '        sahHeight = Convert.ToDecimal(row.Cells("SAHHeight").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso row.Cells("TAHRearOffset").Value IsNot Nothing AndAlso row.Cells("TAHRearOffset").Value.ToString() <> "" Then
    '        tahRearOffset = Convert.ToDecimal(row.Cells("TAHRearOffset").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso row.Cells("SAHRearOffset").Value IsNot Nothing AndAlso row.Cells("SAHRearOffset").Value.ToString() <> "" Then
    '        sahRearOffset = Convert.ToDecimal(row.Cells("SAHRearOffset").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso row.Cells("SAHTopDownOffset").Value IsNot Nothing AndAlso row.Cells("SAHTopDownOffset").Value.ToString() <> "" Then
    '        sahTopDownOffset = Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing Then
    '        musicRestDesign = Convert.ToBoolean(row.Cells("MusicRestDesign").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing Then
    '        chartTemplate = Convert.ToBoolean(row.Cells("Chart_Template").Value)
    '    End If
    '    If dgvModelInformation.Columns.Contains("Notes") AndAlso row.Cells("Notes").Value IsNot Nothing Then
    '        notes = Convert.ToString(row.Cells("Notes").Value)
    '    End If

    '    If String.IsNullOrEmpty(modelName) Then
    '        MessageBox.Show("Model name is required.")
    '        e.Cancel = True
    '        Return
    '    End If

    '    Dim db As New DbConnectionManager()

    '    ' Get the original model name from the DataRowView
    '    Dim originalModelName As String = ""
    '    If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
    '        originalModelName = Convert.ToString(DirectCast(row.DataBoundItem, DataRowView)("ModelName"))
    '    End If

    '    If String.IsNullOrEmpty(originalModelName) OrElse row.IsNewRow Then
    '        If db.ModelExists(selectedManufacturer, selectedSeries, modelName) Then
    '            MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '            e.Cancel = True
    '            Return
    '        End If
    '        db.InsertModel(selectedManufacturer, selectedSeries, modelName, width, depth, height,
    '        optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '        tahWidth, tahHeight, sahWidth, sahHeight,
    '        tahRearOffset, sahRearOffset, sahTopDownOffset,
    '        musicRestDesign, chartTemplate, notes)
    '        LoadModelsForSelectedSeries()
    '    Else
    '        db.UpdateModel(selectedManufacturer, selectedSeries, originalModelName, modelName, width, depth, height,
    '        optionalHeight, optionalDepth, angleType, ampHandleLocation,
    '        tahWidth, tahHeight, sahWidth, sahHeight,
    '        tahRearOffset, sahRearOffset, sahTopDownOffset,
    '        musicRestDesign, chartTemplate, notes)
    '    End If
    'End Sub

    Private Sub dgvModelInformation_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellContentClick
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

        ' Force commit of any edit and new row
        dgvModelInformation.CurrentCell = Nothing
        Application.DoEvents()

        If dgvModelInformation.Columns(e.ColumnIndex).Name = "Save" Then
            Dim row = dgvModelInformation.Rows(e.RowIndex)
            If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
                MessageBox.Show("Please enter data before saving.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
            Dim selectedSeries As String = cmbSeriesName.Text.Trim()
            If String.IsNullOrEmpty(selectedSeries) Then
                MessageBox.Show("Please select a series before saving a model.", "Series Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Gather all values
            Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
            Dim width As Decimal = 0D
            If dgvModelInformation.Columns.Contains("Width") AndAlso row.Cells("Width").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()) Then
                width = Convert.ToDecimal(row.Cells("Width").Value)
            End If
            Dim depth As Decimal = 0D
            If dgvModelInformation.Columns.Contains("Depth") AndAlso row.Cells("Depth").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()) Then
                depth = Convert.ToDecimal(row.Cells("Depth").Value)
            End If
            Dim height As Decimal = 0D
            If dgvModelInformation.Columns.Contains("Height") AndAlso row.Cells("Height").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()) Then
                height = Convert.ToDecimal(row.Cells("Height").Value)
            End If

            ' --- VALIDATION FOR REQUIRED FIELDS ---
            If String.IsNullOrWhiteSpace(modelName) Then
                MessageBox.Show("Model name is required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If width = 0D Then
                MessageBox.Show("Width is required and must be greater than zero.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If depth = 0D Then
                MessageBox.Show("Depth is required and must be greater than zero.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            If height = 0D Then
                MessageBox.Show("Height is required and must be greater than zero.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Gather optional fields as before
            Dim optionalHeight As Decimal? = Nothing
            If dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso row.Cells("OptionalHeight").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()) Then
                optionalHeight = Convert.ToDecimal(row.Cells("OptionalHeight").Value)
            End If
            Dim optionalDepth As Decimal? = Nothing
            If dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso row.Cells("OptionalDepth").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()) Then
                optionalDepth = Convert.ToDecimal(row.Cells("OptionalDepth").Value)
            End If
            Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
            Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
            Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
            Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
            Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
            Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
            Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
            Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
            Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
            Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
            Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
            Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

            ' --- Detect Insert vs Update using DataRowView.RowState ---
            Dim isInsert As Boolean = False
            Dim originalModelName As String = ""
            If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
                Dim drv = DirectCast(row.DataBoundItem, DataRowView)
                originalModelName = Convert.ToString(drv("ModelName"))
                If drv.Row.RowState = DataRowState.Added Then
                    isInsert = True
                End If
            End If

            ' Show confirmation dialog
            Dim msg As String = "You are about to save the following model:" & vbCrLf &
            $"Model Name: {modelName}" & vbCrLf &
            $"Width: {width}" & vbCrLf &
            $"Depth: {depth}" & vbCrLf &
            $"Height: {height}" & vbCrLf &
            $"Optional Height: {optionalHeight}" & vbCrLf &
            $"Optional Depth: {optionalDepth}" & vbCrLf &
            $"Angle Type: {angleType}" & vbCrLf &
            $"Amp Handle Location: {ampHandleLocation}" & vbCrLf &
            $"TAH Width: {tahWidth}" & vbCrLf &
            $"TAH Height: {tahHeight}" & vbCrLf &
            $"SAH Width: {sahWidth}" & vbCrLf &
            $"SAH Height: {sahHeight}" & vbCrLf &
            $"TAH Rear Offset: {tahRearOffset}" & vbCrLf &
            $"SAH Rear Offset: {sahRearOffset}" & vbCrLf &
            $"SAH TopDown Offset: {sahTopDownOffset}" & vbCrLf &
            $"Music Rest Design: {musicRestDesign}" & vbCrLf &
            $"Chart/Template: {chartTemplate}" & vbCrLf &
            $"Notes: {notes}" & vbCrLf &
            "Do you want to save this model?"

            Dim result = MessageBox.Show(msg, "Confirm Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Dim db As New DbConnectionManager()
                If isInsert Then
                    If db.ModelExists(selectedManufacturer, selectedSeries, modelName) Then
                        MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    Try
                        db.InsertModel(selectedManufacturer, selectedSeries, modelName, width, depth, height,
                        optionalHeight, optionalDepth, angleType, ampHandleLocation,
                        tahWidth, tahHeight, sahWidth, sahHeight,
                        tahRearOffset, sahRearOffset, sahTopDownOffset,
                        musicRestDesign, chartTemplate, notes)
                        LoadModelsForSelectedSeries()
                    Catch ex As Exception
                        MessageBox.Show("DB ERROR: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Else
                    Try
                        ' Get modelId from the DataRowView (drv) or DataRow
                        Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
                        Dim modelId As Integer = CInt(drv("PK_ModelId"))
                        db.UpdateModel(modelId, modelName, width, depth, height, optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight, sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chartTemplate, notes)
                        LoadModelsForSelectedSeries()
                    Catch ex As Exception
                        MessageBox.Show("DB ERROR: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub TrySaveNewSeries()
        Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturer)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            ' Require equipment type selection
            Dim equipmentType As String = cmbEquipmentType.Text.Trim()
            If String.IsNullOrEmpty(equipmentType) Then
                MessageBox.Show("Please select or enter an equipment type for this series.", "Equipment Type Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Exit Sub
            End If

            ' Convert equipment type string to integer value
            ' Get the selected EquipmentTypeId from the ComboBox
            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)

            ' Show confirmation message
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                            $"Manufacturer: {selectedManufacturer}" & vbCrLf &
                            $"Series: {enteredSeries}" & vbCrLf &
                            $"Equipment Type: {equipmentType}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(selectedManufacturer, enteredSeries, equipmentTypeValue)
                ' Refresh series list
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturer)
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub

    Private Sub btnSaveSeries_Click(sender As Object, e As EventArgs) Handles btnSaveSeries.Click
        If String.IsNullOrEmpty(cmbManufacturerName.Text.Trim()) OrElse cmbManufacturerName.SelectedIndex = -1 Then
            MessageBox.Show("Please select a manufacturer before saving or updating a series.", "Manufacturer Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        TrySaveOrUpdateSeries()
        LoadSeriesGridForManufacturer()
    End Sub


    Private Sub LoadSeriesGridForManufacturer()
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            dgvSeriesName.DataSource = Nothing
            Return
        End If

        Dim selectedManufacturerId As Integer
        If Not Integer.TryParse(cmbManufacturerName.SelectedValue.ToString(), selectedManufacturerId) Then
            dgvSeriesName.DataSource = Nothing
            Return
        End If

        Dim db As New DbConnectionManager()
        Dim dt = db.GetSeriesAndEquipmentTypeByManufacturer(selectedManufacturerId)
        dgvSeriesName.DataSource = dt
        dgvSeriesName.ClearSelection()

        ' Hide PK_SeriesId column
        If dgvSeriesName.Columns.Contains("PK_SeriesId") Then
            dgvSeriesName.Columns("PK_SeriesId").Visible = False
        End If
    End Sub
    Private Sub dgvSeriesName_SelectionChanged(sender As Object, e As EventArgs) Handles dgvSeriesName.SelectionChanged
        If Not isFormLoaded OrElse dgvSeriesName.SelectedRows.Count = 0 Then Return

        If Not isUserSelectingSeries Then Return
        isUserSelectingSeries = False

        Dim selectedRow = dgvSeriesName.SelectedRows(0)
        Dim selectedSeries As String = Convert.ToString(selectedRow.Cells("SeriesName").Value)
        Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()

        If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(selectedSeries) Then Return

        Dim db As New DbConnectionManager()
        Dim dt = db.GetModelsByManufacturerAndSeries(selectedManufacturer, selectedSeries)
        If dt.Rows.Count = 0 Then
            MessageBox.Show("There are no models associated with this series yet.", "No Models Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
            dgvModelInformation.DataSource = Nothing
        Else
            dgvModelInformation.DataSource = dt
        End If
    End Sub
    Private Sub InitializeModelGridColumns(Optional equipmentTypeName As String = Nothing)
        dgvModelInformation.Columns.Clear()

        ' Always visible columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "ModelName", .HeaderText = "Model Name", .DataPropertyName = "ModelName"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Width", .HeaderText = "Width", .DataPropertyName = "Width"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Depth", .HeaderText = "Depth", .DataPropertyName = "Depth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "Height", .HeaderText = "Height", .DataPropertyName = "Height"})

        ' Add equipment-type-specific columns
        If equipmentTypeName = "Guitar Amplifier" Then
            AddGuitarAmplifierColumns()
        ElseIf equipmentTypeName = "Music Keyboard" Then
            AddMusicKeyboardColumns()
        End If

        ' Always visible for all types
        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "Chart_Template", .HeaderText = "Chart/Template"})
        dgvModelInformation.Columns.Add("Notes", "Notes")

        Dim saveButtonCol As New DataGridViewButtonColumn With {
    .Name = "Save",
    .HeaderText = "Save",
    .Text = "Save",
    .UseColumnTextForButtonValue = True,
    .Width = 60
}
        dgvModelInformation.Columns.Add(saveButtonCol)
        FormatModelGridColumns()
    End Sub

    Private Sub AddGuitarAmplifierColumns()
        Dim db As New DbConnectionManager()

        ' Optional Height/Depth
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalHeight", .HeaderText = "Opt. Height", .DataPropertyName = "OptionalHeight"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "OptionalDepth", .HeaderText = "Opt. Depth", .DataPropertyName = "OptionalDepth"})

        ' AngleType ComboBox
        Dim angleTypeCol As New DataGridViewComboBoxColumn With {
        .Name = "AngleType",
        .HeaderText = "Angle Type",
        .DataPropertyName = "AngleType",
        .Width = 120,
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    }
        angleTypeCol.DataSource = db.GetAllAngleTypesTable()
        angleTypeCol.DisplayMember = "AngleTypeName"
        angleTypeCol.ValueMember = "AngleTypeName"
        angleTypeCol.FlatStyle = FlatStyle.Standard
        angleTypeCol.ReadOnly = False
        dgvModelInformation.Columns.Add(angleTypeCol)

        ' TAH/SAH columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHWidth", .HeaderText = "TAH Width", .DataPropertyName = "TAHWidth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHHeight", .HeaderText = "TAH Height", .DataPropertyName = "TAHHeight"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHWidth", .HeaderText = "SAH Width", .DataPropertyName = "SAHWidth"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHHeight", .HeaderText = "SAH Height", .DataPropertyName = "SAHHeight"})

        ' AmpHandleLocation ComboBox
        Dim ampHandleCol As New DataGridViewComboBoxColumn With {
        .Name = "AmpHandleLocation",
        .HeaderText = "Amp Handle Location",
        .DataPropertyName = "AmpHandleLocation",
        .Width = 125,
        .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
    }
        ampHandleCol.DataSource = db.GetAllAmpHandleLocationsTable()
        ampHandleCol.DisplayMember = "AmpHandleLocationName"
        ampHandleCol.ValueMember = "AmpHandleLocationName"
        ampHandleCol.FlatStyle = FlatStyle.Standard
        ampHandleCol.ReadOnly = False
        dgvModelInformation.Columns.Add(ampHandleCol)

        ' Rear/Offset columns
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "TAHRearOffset", .HeaderText = "TAH Rear Offset", .DataPropertyName = "TAHRearOffset"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHRearOffset", .HeaderText = "SAH Rear Offset", .DataPropertyName = "SAHRearOffset"})
        dgvModelInformation.Columns.Add(New DataGridViewTextBoxColumn With {.Name = "SAHTopDownOffset", .HeaderText = "SAH TopDown Offset", .DataPropertyName = "SAHTopDownOffset"})
    End Sub

    Private Sub AddMusicKeyboardColumns()
        dgvModelInformation.Columns.Add(New DataGridViewCheckBoxColumn With {.Name = "MusicRestDesign", .HeaderText = "Music Rest Design", .DataPropertyName = "MusicRestDesign"})
    End Sub

    'Private Sub UpdateModelGridColumnVisibility(equipmentTypeName As String)
    '    ' No column-adding code here anymore!
    '    ' You can use this for show/hide logic if needed, but not for adding columns.
    'End Sub
    Private Sub FormatModelGridColumns()
        ' Center header text and wrap
        For Each col As DataGridViewColumn In dgvModelInformation.Columns
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            col.HeaderCell.Style.WrapMode = DataGridViewTriState.True
        Next

        dgvModelInformation.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        dgvModelInformation.ColumnHeadersHeight = 44 ' Adjust as needed

        ' Center cell text for specific columns
        Dim centerStyle As New DataGridViewCellStyle()
        centerStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For Each colName In New String() {"Width", "Depth", "Height", "OptionalHeight", "OptionalDepth"}
            If dgvModelInformation.Columns.Contains(colName) Then
                dgvModelInformation.Columns(colName).DefaultCellStyle = centerStyle
            End If
        Next

        ' Set column widths if needed
        Dim widths As New Dictionary(Of String, Integer) From {
        {"ModelName", 150},
        {"Width", 50},
        {"Depth", 50},
        {"Height", 50},
        {"OptionalHeight", 50},
        {"OptionalDepth", 50},
        {"AngleType", 75},
        {"AmpHandleLocation", 125},
        {"TAHWidth", 50},
        {"TAHHeight", 50},
        {"SAHWidth", 50},
        {"SAHHeight", 50},
        {"TAHRearOffset", 80},
        {"SAHRearOffset", 80},
        {"SAHTopDownOffset", 90},
        {"Chart_Template", 90},
        {"Notes", 175}
    }
        For Each kvp In widths
            If dgvModelInformation.Columns.Contains(kvp.Key) Then
                dgvModelInformation.Columns(kvp.Key).Width = kvp.Value
            End If
        Next
    End Sub

    ' Handle when a row begins editing to store original values
    Private Sub dgvModelInformation_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.RowEnter
        Dim row = dgvModelInformation.Rows(e.RowIndex)
        If row.IsNewRow Then Exit Sub

        Dim values As New Dictionary(Of String, Object)
        For Each col As DataGridViewColumn In dgvModelInformation.Columns
            values(col.Name) = row.Cells(col.Name).Value
        Next
        originalRowValues(e.RowIndex) = values
    End Sub

    Private Sub dgvModelInformation_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvModelInformation.RowValidating
        If isLoadingModels Then Return

        Dim row = dgvModelInformation.Rows(e.RowIndex)

        ' Only skip if the row is the placeholder for a new row and is empty
        If row.IsNewRow AndAlso row.Cells.Cast(Of DataGridViewCell).All(Function(c) c.Value Is Nothing OrElse String.IsNullOrWhiteSpace(c.FormattedValue?.ToString())) Then
            Return
        End If

        Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
        Dim selectedSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(selectedSeries) Then
            MessageBox.Show("Please select a series before adding a model.", "Series Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            Return
        End If

        dgvModelInformation.EndEdit()

        ' Gather all values
        Dim modelName As String = If(dgvModelInformation.Columns.Contains("ModelName"), Convert.ToString(row.Cells("ModelName").Value)?.Trim(), Nothing)
        Dim width As Decimal = 0D
        If dgvModelInformation.Columns.Contains("Width") AndAlso row.Cells("Width").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Width").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Width").Value.ToString()) Then
            width = Convert.ToDecimal(row.Cells("Width").Value)
        End If
        Dim depth As Decimal = 0D
        If dgvModelInformation.Columns.Contains("Depth") AndAlso row.Cells("Depth").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Depth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Depth").Value.ToString()) Then
            depth = Convert.ToDecimal(row.Cells("Depth").Value)
        End If
        Dim height As Decimal = 0D
        If dgvModelInformation.Columns.Contains("Height") AndAlso row.Cells("Height").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("Height").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("Height").Value.ToString()) Then
            height = Convert.ToDecimal(row.Cells("Height").Value)
        End If

        Dim optionalHeight As Decimal? = Nothing
        If dgvModelInformation.Columns.Contains("OptionalHeight") AndAlso row.Cells("OptionalHeight").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("OptionalHeight").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalHeight").Value.ToString()) Then
            optionalHeight = Convert.ToDecimal(row.Cells("OptionalHeight").Value)
        End If
        Dim optionalDepth As Decimal? = Nothing
        If dgvModelInformation.Columns.Contains("OptionalDepth") AndAlso row.Cells("OptionalDepth").Value IsNot Nothing AndAlso Not IsDBNull(row.Cells("OptionalDepth").Value) AndAlso Not String.IsNullOrWhiteSpace(row.Cells("OptionalDepth").Value.ToString()) Then
            optionalDepth = Convert.ToDecimal(row.Cells("OptionalDepth").Value)
        End If
        Dim angleType As String = If(dgvModelInformation.Columns.Contains("AngleType"), Convert.ToString(row.Cells("AngleType").Value), Nothing)
        Dim ampHandleLocation As String = If(dgvModelInformation.Columns.Contains("AmpHandleLocation"), Convert.ToString(row.Cells("AmpHandleLocation").Value), Nothing)
        Dim tahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("TAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHWidth").Value)), Convert.ToDecimal(row.Cells("TAHWidth").Value), Nothing)
        Dim tahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("TAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHHeight").Value)), Convert.ToDecimal(row.Cells("TAHHeight").Value), Nothing)
        Dim sahWidth As Decimal? = If(dgvModelInformation.Columns.Contains("SAHWidth") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHWidth").Value)), Convert.ToDecimal(row.Cells("SAHWidth").Value), Nothing)
        Dim sahHeight As Decimal? = If(dgvModelInformation.Columns.Contains("SAHHeight") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHHeight").Value)), Convert.ToDecimal(row.Cells("SAHHeight").Value), Nothing)
        Dim tahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("TAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("TAHRearOffset").Value)), Convert.ToDecimal(row.Cells("TAHRearOffset").Value), Nothing)
        Dim sahRearOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHRearOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHRearOffset").Value)), Convert.ToDecimal(row.Cells("SAHRearOffset").Value), Nothing)
        Dim sahTopDownOffset As Decimal? = If(dgvModelInformation.Columns.Contains("SAHTopDownOffset") AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(row.Cells("SAHTopDownOffset").Value)), Convert.ToDecimal(row.Cells("SAHTopDownOffset").Value), Nothing)
        Dim musicRestDesign As Boolean? = If(dgvModelInformation.Columns.Contains("MusicRestDesign") AndAlso row.Cells("MusicRestDesign").Value IsNot Nothing, Convert.ToBoolean(row.Cells("MusicRestDesign").Value), Nothing)
        Dim chartTemplate As Boolean? = If(dgvModelInformation.Columns.Contains("Chart_Template") AndAlso row.Cells("Chart_Template").Value IsNot Nothing, Convert.ToBoolean(row.Cells("Chart_Template").Value), Nothing)
        Dim notes As String = If(dgvModelInformation.Columns.Contains("Notes"), Convert.ToString(row.Cells("Notes").Value), Nothing)

        If String.IsNullOrEmpty(modelName) Then
            MessageBox.Show("Model name is required.")
            e.Cancel = True
            Return
        End If

        Dim db As New DbConnectionManager()

        ' Get the original model name from the DataRowView
        Dim originalModelName As String = ""
        If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
            originalModelName = Convert.ToString(DirectCast(row.DataBoundItem, DataRowView)("ModelName"))
        End If

        ' Build a dictionary of current values
        Dim currentValues As New Dictionary(Of String, Object) From {
        {"ModelName", modelName},
        {"Width", width},
        {"Depth", depth},
        {"Height", height},
        {"OptionalHeight", optionalHeight},
        {"OptionalDepth", optionalDepth},
        {"AngleType", angleType},
        {"AmpHandleLocation", ampHandleLocation},
        {"TAHWidth", tahWidth},
        {"TAHHeight", tahHeight},
        {"SAHWidth", sahWidth},
        {"SAHHeight", sahHeight},
        {"TAHRearOffset", tahRearOffset},
        {"SAHRearOffset", sahRearOffset},
        {"SAHTopDownOffset", sahTopDownOffset},
        {"MusicRestDesign", musicRestDesign},
        {"Chart_Template", chartTemplate},
        {"Notes", notes}
    }

        ' Check if this is a new row (insert) or update
        Dim isInsert As Boolean = String.IsNullOrEmpty(originalModelName)

        If isInsert Then
            ' Show confirmation dialog for new model
            Dim msg As String = "You are about to add a new model with the following values:" & vbCrLf &
            String.Join(vbCrLf, currentValues.Select(Function(kv) $"{kv.Key}: {If(kv.Value, "")}"))
            Dim result = MessageBox.Show(msg, "Confirm New Model", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then
                e.Cancel = True
                Return
            End If

            If db.ModelExists(selectedManufacturer, selectedSeries, modelName) Then
                MessageBox.Show("A model with this name already exists for the selected manufacturer and series.", "Duplicate Model", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True
                Return
            End If

            db.InsertModel(selectedManufacturer, selectedSeries, modelName, width, depth, height,
            optionalHeight, optionalDepth, angleType, ampHandleLocation,
            tahWidth, tahHeight, sahWidth, sahHeight,
            tahRearOffset, sahRearOffset, sahTopDownOffset,
            musicRestDesign, chartTemplate, notes)
            LoadModelsForSelectedSeries()
        Else
            ' Compare with original values
            Dim originalValues As Dictionary(Of String, Object) = Nothing
            If originalRowValues.TryGetValue(e.RowIndex, originalValues) Then
                Dim changes = currentValues.Where(Function(kv) originalValues.ContainsKey(kv.Key) AndAlso Not ValuesAreEquivalent(originalValues(kv.Key), kv.Value)).ToList()

                If changes.Count > 0 Then
                    Dim msg As String = "You are about to update this model. Changes:" & vbCrLf &
                    String.Join(vbCrLf, changes.Select(Function(kv) $"{kv.Key}: {originalValues(kv.Key)} -> {If(kv.Value, "")}")) &
                    vbCrLf & "Do you want to save these changes?"
                    Dim result = MessageBox.Show(msg, "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If result = DialogResult.No Then
                        ' Revert changes
                        For Each kv In originalValues
                            If dgvModelInformation.Columns.Contains(kv.Key) Then
                                row.Cells(kv.Key).Value = kv.Value
                            End If
                        Next
                        e.Cancel = True
                        Return
                    End If
                Else
                    ' No changes, skip update
                    Return
                End If
            End If

            ' --- CORRECTED UPDATE CALL ---
            If row.DataBoundItem IsNot Nothing AndAlso TypeOf row.DataBoundItem Is DataRowView Then
                Dim drv As DataRowView = DirectCast(row.DataBoundItem, DataRowView)
                Dim modelId As Integer = CInt(drv("PK_ModelId"))
                db.UpdateModel(modelId, modelName, width, depth, height, optionalHeight, optionalDepth, angleType, ampHandleLocation, tahWidth, tahHeight, sahWidth, sahHeight, tahRearOffset, sahRearOffset, sahTopDownOffset, musicRestDesign, chartTemplate, notes)
            End If
        End If
    End Sub
    Private Function ValuesAreEquivalent(val1 As Object, val2 As Object) As Boolean
        ' Treat Nothing, empty string, 0, and False as equivalent if both sides are "empty"
        If IsNothing(val1) AndAlso IsNothing(val2) Then Return True
        If IsNothing(val1) AndAlso (val2 Is "" OrElse val2 Is DBNull.Value) Then Return True
        If IsNothing(val2) AndAlso (val1 Is "" OrElse val1 Is DBNull.Value) Then Return True

        ' Numeric zero equivalence
        If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso IsNumeric(val2) AndAlso Convert.ToDecimal(val2) = 0 Then Return True
        If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso IsNumeric(val1) AndAlso Convert.ToDecimal(val1) = 0 Then Return True

        ' Boolean false equivalence
        If (IsNothing(val1) OrElse val1 Is "" OrElse val1 Is DBNull.Value) AndAlso TypeOf val2 Is Boolean AndAlso CBool(val2) = False Then Return True
        If (IsNothing(val2) OrElse val2 Is "" OrElse val2 Is DBNull.Value) AndAlso TypeOf val1 Is Boolean AndAlso CBool(val1) = False Then Return True

        ' Otherwise, use normal equality
        Return Object.Equals(val1, val2)
    End Function

    Private Sub dgvModelInformation_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvModelInformation.CellEndEdit
        dgvModelInformation.CommitEdit(DataGridViewDataErrorContexts.Commit)
        dgvModelInformation.EndEdit()
    End Sub
    Private Sub dgvModelInformation_Leave(sender As Object, e As EventArgs) Handles dgvModelInformation.Leave
        ' If the current row is the new row and has data, force commit and validation
        If dgvModelInformation.CurrentRow IsNot Nothing AndAlso dgvModelInformation.CurrentRow.IsNewRow Then
            Dim hasData = dgvModelInformation.CurrentRow.Cells.Cast(Of DataGridViewCell) _
            .Any(Function(c) c.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(c.FormattedValue?.ToString()))
            If hasData Then
                ' Force the row to be validated and committed
                dgvModelInformation.CurrentRow.Selected = True
                dgvModelInformation.EndEdit()
                dgvModelInformation.CurrentCell = Nothing
                ' Explicitly call Validate to ensure RowValidating fires
                Me.Validate()
            End If
        ElseIf dgvModelInformation.IsCurrentRowDirty Then
            dgvModelInformation.EndEdit()
        End If
    End Sub




    Private Sub btnUploadModels_Click(sender As Object, e As EventArgs) Handles btnUploadModels.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Excel or CSV Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All Files (*.*)|*.*"
            ofd.Title = "Select Model Information File"
            If ofd.ShowDialog() <> DialogResult.OK Then Return

            Dim filePath = ofd.FileName
            Dim dt As New DataTable()

            If Path.GetExtension(filePath).ToLower() = ".csv" Then
                ' --- CSV Handling ---
                Dim lines = File.ReadAllLines(filePath)
                If lines.Length < 2 Then
                    MessageBox.Show("The selected file does not contain enough data.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                Dim headers = lines(0).Split(","c).Select(Function(h) h.Trim()).ToArray()
                For Each header In headers
                    dt.Columns.Add(header)
                Next
                For i As Integer = 1 To lines.Length - 1
                    Dim values = lines(i).Split(","c).Select(Function(v) v.Trim()).ToArray()
                    If values.Length = headers.Length Then
                        dt.Rows.Add(values)
                    End If
                Next
            ElseIf Path.GetExtension(filePath).ToLower() = ".xlsx" OrElse Path.GetExtension(filePath).ToLower() = ".xls" Then
                ' --- Excel Handling ---
                Dim connStr As String
                If Path.GetExtension(filePath).ToLower() = ".xlsx" Then
                    connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    connStr = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties='Excel 8.0;HDR=YES';"
                End If
                Try
                    Using conn As New OleDbConnection(connStr)
                        conn.Open()
                        Dim dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                        Dim sheetName = dtSchema.Rows(0)("TABLE_NAME").ToString()
                        Using cmd As New OleDbCommand($"SELECT * FROM [{sheetName}]", conn)
                            Using da As New OleDbDataAdapter(cmd)
                                da.Fill(dt)
                            End Using
                        End Using
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error reading Excel file: " & ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End Try
            Else
                MessageBox.Show("Unsupported file type.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' --- Validate required columns ---
            Dim requiredCols = {"ModelName", "Width", "Depth", "Height"}
            For Each col In requiredCols
                If Not dt.Columns.Contains(col) Then
                    MessageBox.Show($"Missing required column: {col}", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            Next

            ' --- Insert each row into the database ---
            Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
            Dim selectedSeries As String = cmbSeriesName.Text.Trim()
            If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(selectedSeries) Then
                MessageBox.Show("Please select a manufacturer and series before uploading.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim db As New DbConnectionManager()
            Dim insertedCount As Integer = 0
            For Each row As DataRow In dt.Rows
                If String.IsNullOrWhiteSpace(row("ModelName").ToString()) Then Continue For
                Dim modelName = row("ModelName").ToString().Trim()
                Dim width = Convert.ToDecimal(row("Width"))
                Dim depth = Convert.ToDecimal(row("Depth"))
                Dim height = Convert.ToDecimal(row("Height"))

                ' Optional fields
                Dim optionalHeight = If(dt.Columns.Contains("OptionalHeight") AndAlso Not IsDBNull(row("OptionalHeight")) AndAlso row("OptionalHeight").ToString() <> "", CType(row("OptionalHeight"), Decimal?), Nothing)
                Dim optionalDepth = If(dt.Columns.Contains("OptionalDepth") AndAlso Not IsDBNull(row("OptionalDepth")) AndAlso row("OptionalDepth").ToString() <> "", CType(row("OptionalDepth"), Decimal?), Nothing)
                Dim angleType = If(dt.Columns.Contains("AngleType"), row("AngleType").ToString(), Nothing)
                Dim ampHandleLocation = If(dt.Columns.Contains("AmpHandleLocation"), row("AmpHandleLocation").ToString(), Nothing)
                Dim tahWidth = If(dt.Columns.Contains("TAHWidth") AndAlso Not IsDBNull(row("TAHWidth")) AndAlso row("TAHWidth").ToString() <> "", CType(row("TAHWidth"), Decimal?), Nothing)
                Dim tahHeight = If(dt.Columns.Contains("TAHHeight") AndAlso Not IsDBNull(row("TAHHeight")) AndAlso row("TAHHeight").ToString() <> "", CType(row("TAHHeight"), Decimal?), Nothing)
                Dim sahWidth = If(dt.Columns.Contains("SAHWidth") AndAlso Not IsDBNull(row("SAHWidth")) AndAlso row("SAHWidth").ToString() <> "", CType(row("SAHWidth"), Decimal?), Nothing)
                Dim sahHeight = If(dt.Columns.Contains("SAHHeight") AndAlso Not IsDBNull(row("SAHHeight")) AndAlso row("SAHHeight").ToString() <> "", CType(row("SAHHeight"), Decimal?), Nothing)
                Dim tahRearOffset = If(dt.Columns.Contains("TAHRearOffset") AndAlso Not IsDBNull(row("TAHRearOffset")) AndAlso row("TAHRearOffset").ToString() <> "", CType(row("TAHRearOffset"), Decimal?), Nothing)
                Dim sahRearOffset = If(dt.Columns.Contains("SAHRearOffset") AndAlso Not IsDBNull(row("SAHRearOffset")) AndAlso row("SAHRearOffset").ToString() <> "", CType(row("SAHRearOffset"), Decimal?), Nothing)
                Dim sahTopDownOffset = If(dt.Columns.Contains("SAHTopDownOffset") AndAlso Not IsDBNull(row("SAHTopDownOffset")) AndAlso row("SAHTopDownOffset").ToString() <> "", CType(row("SAHTopDownOffset"), Decimal?), Nothing)
                Dim musicRestDesign = If(dt.Columns.Contains("MusicRestDesign") AndAlso Not IsDBNull(row("MusicRestDesign")) AndAlso row("MusicRestDesign").ToString() <> "", CBool(row("MusicRestDesign")), CType(Nothing, Boolean?))
                Dim chartTemplate = If(dt.Columns.Contains("Chart_Template") AndAlso Not IsDBNull(row("Chart_Template")) AndAlso row("Chart_Template").ToString() <> "", CBool(row("Chart_Template")), CType(Nothing, Boolean?))
                Dim notes = If(dt.Columns.Contains("Notes"), row("Notes").ToString(), Nothing)

                If db.ModelExists(selectedManufacturer, selectedSeries, modelName) Then Continue For

                Try
                    db.InsertModel(selectedManufacturer, selectedSeries, modelName, width, depth, height,
                        optionalHeight, optionalDepth, angleType, ampHandleLocation,
                        tahWidth, tahHeight, sahWidth, sahHeight,
                        tahRearOffset, sahRearOffset, sahTopDownOffset,
                        musicRestDesign, chartTemplate, notes)
                    insertedCount += 1
                Catch ex As Exception
                    MessageBox.Show($"Error inserting model '{modelName}': {ex.Message}", "Insert Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Next

            LoadModelsForSelectedSeries()
            MessageBox.Show($"{insertedCount} models imported successfully.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub
    Private Sub TrySaveOrUpdateSeries()
        Dim selectedManufacturer As String = cmbManufacturerName.Text.Trim()
        Dim enteredSeries As String = cmbSeriesName.Text.Trim()
        If String.IsNullOrEmpty(selectedManufacturer) OrElse String.IsNullOrEmpty(enteredSeries) Then Exit Sub

        Dim db As New DbConnectionManager()
        Dim existingSeries = db.GetSeriesNamesByManufacturer(selectedManufacturer)
        If Not existingSeries.AsEnumerable().Any(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase)) Then
            ' Require equipment type selection
            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If

            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim msg As String = $"Do you want to save this new series?" & vbCrLf &
                                $"Manufacturer: {selectedManufacturer}" & vbCrLf &
                                $"Series: {enteredSeries}" & vbCrLf &
                                $"Equipment Type: {cmbEquipmentType.Text.Trim()}"
            Dim result = MessageBox.Show(msg, "Save New Series", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                db.InsertSeries(selectedManufacturer, enteredSeries, equipmentTypeValue)
                cmbSeriesName.DataSource = Nothing
                cmbSeriesName.DataSource = db.GetSeriesNamesByManufacturer(selectedManufacturer)
                cmbSeriesName.SelectedItem = enteredSeries
            End If
        Else
            ' Update equipment type for existing series
            If cmbEquipmentType.SelectedValue Is Nothing OrElse Not IsNumeric(cmbEquipmentType.SelectedValue) Then
                MessageBox.Show("Please select a valid equipment type.", "Invalid Equipment Type", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbEquipmentType.Focus()
                Return
            End If
            Dim equipmentTypeValue As Integer = Convert.ToInt32(cmbEquipmentType.SelectedValue)
            Dim seriesId As Integer = CInt(existingSeries.AsEnumerable().First(Function(row) String.Equals(row.Field(Of String)("SeriesName"), enteredSeries, StringComparison.OrdinalIgnoreCase))("PK_SeriesId"))
            db.UpdateSeriesEquipmentType(seriesId, equipmentTypeValue)
            MessageBox.Show("Series equipment type updated.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        SetEquipmentTypeForSelectedSeries()
    End Sub
End Class