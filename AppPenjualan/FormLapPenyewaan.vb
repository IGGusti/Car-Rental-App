Imports System.Data.SqlClient
Public Class FormLapPenyewaan
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select TglSewa as TglSewa,TglMasuk as TglMasuk,Plat as Plat,NamaMobil as NamaMobil,KTPPenyewa as KTPPenyewa,NamaPenyewa as NamaPenyewa,LamaSewa as LamaSewa,TotalSewa as TotalSewa From TBL_LAPSEWA", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_LAPSEWA")
        DataGridView1.DataSource = DS.Tables("TBL_LAPSEWA")
        DataGridView1.ReadOnly = True
    End Sub

    Private Sub FormLapPenyewaan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TampilGrid()
        Call Koneksi()
        CMD = New SqlCommand("SELECT SUM(TotalSewa) as TotalSewa FROM TBL_LAPSEWA", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        Label3.Text = RD.Item("TotalSewa")
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_LAPSEWA where NamaMobil like '%" & TextBox1.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_LAPSEWA where NamaMobil like '%" & TextBox1.Text & "%'", CONN)
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