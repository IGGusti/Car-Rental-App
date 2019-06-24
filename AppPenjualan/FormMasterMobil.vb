Imports System.Data.SqlClient
Public Class FormMasterMobil
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select Plat as Plat,NamaMobil as NamaMobil,STNK as STNK,BPKB as BPKB, WarnaMobil as WarnaMobil,HargaMobil as HargaMobil,HargaSewa as HargaSewa From TBL_DETAILMOBIL", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_DETAILMOBIL")
        DataGridView1.DataSource = DS.Tables("TBL_DETAILMOBIL")
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub FormDetailMobil_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""

        Call TampilGrid()
        TextBox1.Focus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_DETAILMOBIL where Plat='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                MsgBox("Plat sudah terpakai, mohon cek ulang ya bosq!!")
            Else
                Call Koneksi()
                Dim simpan As String = "insert into TBL_DETAILMOBIL values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil di Input", MsgBoxStyle.Information, "Information")
                Call TampilGrid()

                'Dim save As String = "insert into TBL_MOBIL values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "')"
                'CMD = New SqlCommand(save, CONN)
                'CMD.ExecuteNonQuery()

                Dim nyimpen As String = "insert into TBL_STOKMOBILREADY values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "')"
                CMD = New SqlCommand(nyimpen, CONN)
                CMD.ExecuteNonQuery()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_STOKMOBILREADY  where Plat='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_DETAILMOBIL where Plat='" & TextBox1.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    Call Koneksi()
                    Dim edit As String = "update TBL_DETAILMOBIL set NamaMobil='" & TextBox2.Text & "',STNK='" & TextBox3.Text & "',BPKB='" & TextBox4.Text & "',WarnaMobil='" & TextBox5.Text & "',HargaMobil='" & TextBox6.Text & "',HargaSewa='" & TextBox7.Text & "' where Plat='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(edit, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Data berhasil di Edit", MsgBoxStyle.Information, "Information")
                    Call TampilGrid()

                    'Dim ganti As String = "update TBL_MOBIL set NamaMobil='" & TextBox2.Text & "',WarnaMobil='" & TextBox5.Text & "',HargaSewa='" & TextBox7.Text & "' where Plat='" & TextBox1.Text & "'"
                    'CMD = New SqlCommand(ganti, CONN)
                    'CMD.ExecuteNonQuery()

                    Dim ubah As String = "update TBL_STOKMOBILREADY set NamaMobil='" & TextBox2.Text & "',WarnaMobil='" & TextBox5.Text & "',HargaSewa='" & TextBox7.Text & "' where Plat='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(ubah, CONN)
                    CMD.ExecuteNonQuery()
                Else
                    MsgBox("Plat tidak bisa diubah, mohon cek ulang bosq!!")
                End If
            Else
                MsgBox("Mobil masih ada di luar, kembalikan dulu dari Tabel Mobil Keluar bosq!!!")
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        CMD = New SqlCommand("Select * from TBL_DETAILMOBIL where Plat='" & DataGridView1.Item(0, i).Value & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Focus()
        Else
            TextBox1.Text = RD.Item("Plat")
            TextBox2.Text = RD.Item("NamaMobil")
            TextBox3.Text = RD.Item("STNK")
            TextBox4.Text = RD.Item("BPKB")
            TextBox5.Text = RD.Item("WarnaMobil")
            TextBox6.Text = RD.Item("HargaMobil")
            TextBox7.Text = RD.Item("HargaSewa")
            TextBox2.Focus()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Klik 2x, pada data yang ingin dihapus bosq!!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_STOKMOBILREADY  where Plat='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_DETAILMOBIL  where Plat='" & TextBox1.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    Call Koneksi()

                    Dim del As String = "delete from TBL_STOKMOBILREADY where Plat='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(del, CONN)
                    CMD.ExecuteNonQuery()

                    Dim hapus As String = "delete from TBL_DETAILMOBIL where Plat='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(hapus, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Data berhasil di Hapus", MsgBoxStyle.Information, "Information")

                    Call TampilGrid()

                    'Dim apus As String = "delete from TBL_MOBIL where Plat='" & TextBox1.Text & "'"
                    'CMD = New SqlCommand(apus, CONN)
                    'CMD.ExecuteNonQuery()

                Else
                    MsgBox("Data Mobil sudah dihapus bosq!!!")
                End If
            Else
                MsgBox("Mobil masih ada di luar, kembalikan dulu dari Tabel Mobil Keluar bosq!!!")
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
        Call FormMenuUtama.TampilGrid()
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_DETAILMOBIL where NamaMobil like '%" & TextBox8.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_DETAILMOBIL where NamaMobil like '%" & TextBox8.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""

    End Sub
End Class