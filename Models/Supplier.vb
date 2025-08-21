Imports System.ComponentModel.DataAnnotations

Public Class Supplier
    Public Property SupplierId As Integer
    <Required>
    Public Property CompanyName As String
    Public Property Contact1 As String
    Public Property Contact2 As String
    Public Property Phone1 As String
    Public Property Phone2 As String
    Public Property Email1 As String
    Public Property Email2 As String
    Public Property Website As String
    Public Property Address1 As String
    Public Property Address2 As String
    Public Property City As String
    Public Property State As String
    Public Property ZipPostal As String
    Public Overridable Property Fabrics As ICollection(Of Fabric)
End Class