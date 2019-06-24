Imports System.Data.SqlClient
Public Class FormPrintStruk

    Private Sub FormPrintStruk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_PRINTSTRUK where Plat='" & Label17.Text & "' AND TglSewa='" & Label11.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then

        Else
            Call Koneksi()
            Dim simpan As String = "insert into TBL_PRINTSTRUK values ('" & Label17.Text & "','" & Label16.Text & "','" & Label15.Text & "','" & Label14.Text & "','" & Label11.Text & "','" & Label13.Text & "','" & Label18.Text & "','" & Label20.Text & "','" & Label12.Text & "')"
            CMD = New SqlCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
        End If

        Me.PrintForm1.PrintAction = Printing.PrintAction.PrintToPreview
        Me.PrintForm1.Print()
    End Sub
End Class