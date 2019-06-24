Imports System.Data.SqlClient
Public Class FormMobilKeluar
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select TglSewa as TglSewa,KTPPenyewa as KTPPenyewa,Plat as Plat,NamaMobil as NamaMobil,WarnaMobil as WarnaMobil,HargaSewa as HargaSewa,NamaPenyewa as NamaPenyewa,AlamatPenyewa as AlamatPenyewa,TelpPenyewa as TelpPenyewa,LamaSewa as LamaSewa,TotalSewa as TotalSewa, Jaminan as Jaminan From TBL_MOBILKELUAR", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_MOBILKELUAR")
        DataGridView1.DataSource = DS.Tables("TBL_MOBILKELUAR")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub FormMobilKeluar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_MOBILKELUAR where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_MOBILKELUAR where NamaPenyewa like '%" & TextBox1.Text & "%'", CONN)
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