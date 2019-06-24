Imports System.Data.SqlClient
Public Class FormMenuUtama

    Dim argaewa As Integer

    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogoutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        LaporanToolStripMenuItem.Enabled = False
        UtilityToolStripMenuItem.Enabled = False
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        GroupBox4.Enabled = False
        Button2.Enabled = False
        Button1.Enabled = True
    End Sub

    Private Sub FormMenuUtama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call Terkunci()
        Call TampilGrid()
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub

    Private Sub LoginToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        FormLogin.ShowDialog()
    End Sub

    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Dim Pilih = MessageBox.Show("Logout", "Yakin, mau logout??", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Pilih = Windows.Forms.DialogResult.Yes Then
            Call Terkunci()
        End If
    End Sub

    Private Sub AdminToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdminToolStripMenuItem.Click
        FormMasterAdmin.ShowDialog()
    End Sub

    Private Sub DataKaryawanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataKaryawanToolStripMenuItem.Click
        FormMasterKaryawan.ShowDialog()
    End Sub

    Private Sub DataMobiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataMobiToolStripMenuItem.Click
        FormMasterMobil.ShowDialog()
    End Sub

    Private Sub MobilToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MobilToolStripMenuItem.Click
        FormMobil.ShowDialog()
    End Sub

    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select Plat as Plat,NamaMobil as NamaMobil,WarnaMobil as WarnaMobil,HargaSewa as HargaSewa From TBL_STOKMOBILREADY", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_STOKMOBILREADY")
        DataGridView1.DataSource = DS.Tables("TBL_STOKMOBILREADY")
        DataGridView1.ReadOnly = True

        Call Koneksi()
        DA = New SqlDataAdapter("select Plat as Plat,NamaMobil as NamaMobil,WarnaMobil as WarnaMobil,HargaSewa as HargaSewa,TglSewa as TglSewa,KTPPenyewa as KTPPenyewa,NamaPenyewa as NamaPenyewa,AlamatPenyewa as AlamatPenyewa,TelpPenyewa as TelpPenyewa,LamaSewa as LamaSewa,TotalSewa as TotalSewa, Jaminan as Jaminan From TBL_MOBILKELUAR", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_MOBILKELUAR")
        DataGridView2.DataSource = DS.Tables("TBL_MOBILKELUAR")
        DataGridView2.ReadOnly = True
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_STOKMOBILREADY where NamaMobil like '%" & TextBox1.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_STOKMOBILREADY where NamaMobil like '%" & TextBox1.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FormLogin.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Pilih = MessageBox.Show("Logout", "Yakin, mau logout??", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If Pilih = Windows.Forms.DialogResult.Yes Then
            Label22.Text = ""
            Call Terkunci()
        End If
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        Label9.Text = "Plat"
        Label10.Text = "Mobil"
        Label11.Text = "Warna"
        Label12.Text = "HargaSewa"
    End Sub
    Private Sub MobilMasukToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MobilMasukToolStripMenuItem.Click
        FormMobilMasuk.ShowDialog()
    End Sub

    Private Sub MobilKeluarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MobilKeluarToolStripMenuItem.Click
        FormMobilKeluar.ShowDialog()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        CMD = New SqlCommand("Select * from TBL_STOKMOBILREADY where Plat='" & DataGridView1.Item(0, i).Value & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Focus()
        Else
            Label9.Text = RD.Item("Plat")
            Label10.Text = RD.Item("NamaMobil")
            Label11.Text = RD.Item("WarnaMobil")
            Label12.Text = RD.Item("HargaSewa")
            TextBox2.Focus()
        End If
    End Sub
    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView2.CurrentRow.Index
        CMD = New SqlCommand("Select * from TBL_MOBILKELUAR where Plat='" & DataGridView2.Item(0, i).Value & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox9.Text = "Plat"
            Button5.Focus()
        Else
            TextBox9.Text = RD.Item("Plat")
            TextBox10.Text = RD.Item("NamaMobil")
            TextBox11.Text = RD.Item("WarnaMobil")
            argaewa = RD.Item("HargaSewa")
            TextBox13.Text = RD.Item("KTPPenyewa")
            TextBox14.Text = RD.Item("NamaPenyewa")
            Button5.Focus()
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        ElseIf Label9.Text = "Plat" Or Label10.Text = "Mobil" Or Label11.Text = "Warna" Or Label12.Text = "HargaSewa" Then
            MsgBox("Klik 2X Nama Mobil yang ingin dipilih pada Stok Mobil ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_STOKMOBILREADY where Plat='" & Label9.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_LAPSEWA where Plat='" & Label9.Text & "' AND TglSewa='" & TextBox6.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    MsgBox("Tanggalnya sepertinya salah, mohon cek ulang bosq!!")
                Else
                    Call Koneksi()
                    Dim simpan As String = "insert into TBL_MOBILKELUAR values ('" & Label9.Text & "','" & Label10.Text & "','" & Label11.Text & "','" & Label12.Text & "','" & TextBox6.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "','" & Label14.Text & "','" & TextBoxPrint.Text & "')"
                    CMD = New SqlCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Mobil berhasil di Sewa", MsgBoxStyle.Information, "Information")
                    Call TampilGrid()

                    Call Koneksi()
                    CMD = New SqlCommand("select * from TBL_PENYEWA where KTPPenyewa='" & TextBox2.Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    If RD.HasRows Then
                        Dim totalsebelum As Integer
                        Dim totalsewa As Integer

                        totalsebelum = RD.Item("TotalSewa")
                        totalsewa = Val(totalsebelum) + Val(Label14.Text)

                        Call Koneksi()
                        Dim edit As String = "update TBL_PENYEWA set TotalSewa='" & totalsewa & "' where KTPPenyewa='" & TextBox2.Text & "'"
                        CMD = New SqlCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    Else
                        Call Koneksi()
                        Dim save As String = "insert into TBL_PENYEWA values ('" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & Label14.Text & "')"
                        CMD = New SqlCommand(save, CONN)
                        CMD.ExecuteNonQuery()
                    End If

                    Call Koneksi()
                    Dim del As String = "delete from TBL_STOKMOBILREADY where Plat='" & Label9.Text & "'"
                    CMD = New SqlCommand(del, CONN)
                    CMD.ExecuteNonQuery()

                    Call Koneksi()
                    Dim nyimpen As String = "insert into TBL_LAPSEWA values ('" & TextBox6.Text & "','" & "Belum Masuk" & "','" & Label9.Text & "','" & Label10.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox7.Text & "','" & Label14.Text & "')"
                    CMD = New SqlCommand(nyimpen, CONN)
                    CMD.ExecuteNonQuery()

                    Call TampilGrid()

                    FormPrintStruk.Label11.Text = TextBox6.Text
                    FormPrintStruk.Label12.Text = TextBox3.Text
                    FormPrintStruk.Label17.Text = Label9.Text
                    FormPrintStruk.Label16.Text = Label10.Text
                    FormPrintStruk.Label15.Text = Label11.Text
                    FormPrintStruk.Label14.Text = Label12.Text
                    FormPrintStruk.Label13.Text = TextBox7.Text
                    FormPrintStruk.Label18.Text = Label14.Text
                    FormPrintStruk.Label20.Text = TextBoxPrint.Text

                    FormPrintStruk.ShowDialog()
                End If
            Else
                MsgBox("Mobil sudah disewa, mohon cek ulang bosq!!")
            End If
        End If
    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer

        x = Val(TextBox7.Text)
        y = Val(Label12.Text)
        z = x * y
        Label14.Text = z

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBoxPrint.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        Label9.Text = "Plat"
        Label10.Text = "Mobil"
        Label11.Text = "Warna"
        Label12.Text = "Harga Sewa"
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox9.Text = "" Or TextBox10.Text = "" Or TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox13.Text = "" Or TextBox14.Text = "" Then
            MsgBox("Klik 2X Nama Mobil yang ingin dipilih pada Mobil Keluar ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_MOBILKELUAR where Plat='" & TextBox9.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_MOBILMASUK where Plat='" & TextBox9.Text & "' AND TglMasuk='" & TextBox12.Text & "' AND KTPPenyewa='" & TextBox13.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    MsgBox("Tanggalnya sepertinya salah, mohon cek ulang bosq!!")
                Else
                    Call Koneksi()
                    Dim simpan As String = "insert into TBL_MOBILMASUK values ('" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "','" & argaewa & "','" & TextBox12.Text & "','" & TextBox13.Text & "','" & TextBox14.Text & "')"
                    CMD = New SqlCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Mobil berhasil di Kembalikan", MsgBoxStyle.Information, "Information")

                    Call Koneksi()
                    CMD = New SqlCommand("select * from TBL_MOBILKELUAR where Plat='" & TextBox9.Text & "' and KTPPenyewa='" & TextBox13.Text & "'", CONN)
                    RD = CMD.ExecuteReader
                    RD.Read()
                    If RD.HasRows Then
                        Call Koneksi()
                        Dim edit As String = "update TBL_LAPSEWA set TglMasuk='" & TextBox12.Text & "' where Plat='" & TextBox9.Text & "' and KTPPenyewa='" & TextBox13.Text & "'"
                        CMD = New SqlCommand(edit, CONN)
                        CMD.ExecuteNonQuery()
                    End If

                    Call Koneksi()
                    Dim del As String = "delete from TBL_MOBILKELUAR where Plat='" & TextBox9.Text & "'"
                    CMD = New SqlCommand(del, CONN)
                    CMD.ExecuteNonQuery()

                    Call Koneksi()
                    Dim nyimpen As String = "insert into TBL_STOKMOBILREADY values ('" & TextBox9.Text & "','" & TextBox10.Text & "','" & TextBox11.Text & "','" & argaewa & "')"
                    CMD = New SqlCommand(nyimpen, CONN)
                    CMD.ExecuteNonQuery()
                    Call TampilGrid()






                End If
            Else
                MsgBox("Mobil sudah dikembalikan, mohon cek ulang bosq!!")
            End If
        End If
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_MOBILKELUAR where NamaPenyewa like '%" & TextBox8.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_MOBILKELUAR where NamaPenyewa like '%" & TextBox8.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView2.DataSource = DS.Tables("ketemu")
            DataGridView2.ReadOnly = True
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
    End Sub

    Private Sub LapPenyewaanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LapPenyewaanToolStripMenuItem.Click
        FormLapPenyewaan.ShowDialog()
    End Sub

    Private Sub GantiPasswordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        FormGantiPassword.ShowDialog()
    End Sub

    Private Sub PenyewaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PenyewaToolStripMenuItem.Click
        FormPenyewa.ShowDialog()
    End Sub

    Private Sub TextBox12_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.ValueChanged

    End Sub

End Class