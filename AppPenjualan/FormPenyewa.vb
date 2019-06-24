Imports System.Data.SqlClient
Public Class FormPenyewa
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select KTPPenyewa as KTPPenyewa,NamaPenyewa as NamaPenyewa,AlamatPenyewa as AlamatPenyewa,TelpPenyewa as TelpPenyewa,TotalSewa as TotalSewa From TBL_PENYEWA", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_PENYEWA")
        DataGridView1.DataSource = DS.Tables("TBL_PENYEWA")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub FormPenyewa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_PENYEWA where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_PENYEWA where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class