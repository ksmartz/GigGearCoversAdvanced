Imports System.Data.SqlClient

Public Class DbConnectionManager
    Private Shared ReadOnly connectionString As String = "Data Source=MYPC\SQLEXPRESS;Initial Catalog=GigGearCoversDb;Integrated Security=True"

    Public Shared Function GetConnection() As SqlConnection
        Dim conn As New SqlConnection(connectionString)
        conn.Open()
        Return conn
    End Function
End Class