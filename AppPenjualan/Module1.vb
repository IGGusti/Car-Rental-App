Imports System.Data.SqlClient
Module Module1
    Public CONN As SqlConnection
    Public CMD As SqlCommand
    Public DS As New DataSet
    Public DA As SqlDataAdapter
    Public RD As SqlDataReader
    Public LokasiData As String
    Public Sub Koneksi()
        LokasiData = "data source=IGG;initial catalog=DBJNMJUAL;integrated security=true"
        CONN = New SqlConnection(LokasiData)
        If CONN.State = ConnectionState.Closed Then
            CONN.Open()
        End If
    End Sub
End Module