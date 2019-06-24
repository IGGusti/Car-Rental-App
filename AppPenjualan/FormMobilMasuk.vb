Imports System.Data.SqlClient
Public Class FormMobilMasuk
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select TglMasuk as TglMasuk,Plat as Plat,NamaMobil as NamaMobil,WarnaMobil as WarnaMobil,HargaSewa as HargaSewa,KTPPenyewa as KTPPenyewa,NamaPenyewa as NamaPenyewa From TBL_MOBILMASUK", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_MOBILMASUK")
        DataGridView1.DataSource = DS.Tables("TBL_MOBILMASUK")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub FormMobilMasuk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_MOBILMASUK where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_MOBILMASUK where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
End Class