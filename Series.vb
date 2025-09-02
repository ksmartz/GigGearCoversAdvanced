Imports System.Reflection

Public Enum EquipmentType
    GuitarAmplifier
    MusicKeyboard
End Enum

Public Class Series
    Public Property SeriesId As Integer
    Public Property ManufacturerId As Integer
    Public Property Name As String ' Unique per Manufacturer
    Public Property EquipmentType As EquipmentType
    Public Property ModelList As New List(Of Model)
End Class
