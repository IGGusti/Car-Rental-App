Imports System.Data.SqlClient
Public Class FormMasterKaryawan
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select KTPKaryawan as KTP,NamaKaryawan as Nama,AlamatKaryawan as Alamat,TelpKaryawan as Telp From TBL_KARYAWAN", CONN)
        DS = New DataSet
        DA.Fill(DS, "TBL_KARYAWAN")
        DataGridView1.DataSource = DS.Tables("TBL_KARYAWAN")
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                MsgBox("KTP Karyawan sudah terpakai, mohon cek ulang bosq!!")
            Else
                Call Koneksi()

                Dim save As String = "insert into TBL_ADMIN values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & "12345" & "','" & "USER" & "')"
                CMD = New SqlCommand(save, CONN)
                CMD.ExecuteNonQuery()

                Dim simpan As String = "insert into TBL_KARYAWAN values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil di Input", MsgBoxStyle.Information, "Information")

                Call TampilGrid()

            End If
        End If
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        CMD = New SqlCommand("Select * from TBL_KARYAWAN where KTPKaryawan='" & DataGridView1.Item(0, i).Value & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Focus()
        Else
            TextBox1.Text = RD.Item("KTPKaryawan")
            TextBox2.Text = RD.Item("NamaKaryawan")
            TextBox3.Text = RD.Item("AlamatKaryawan")
            TextBox4.Text = RD.Item("TelpKaryawan")
            TextBox2.Focus()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()

                Dim ubah As String = "update TBL_ADMIN set NamaAdmin='" & TextBox2.Text & "',LevelAdmin='" & "USER" & "' where KodeAdmin='" & TextBox1.Text & "'"
                CMD = New SqlCommand(ubah, CONN)
                CMD.ExecuteNonQuery()

                Dim edit As String = "update TBL_KARYAWAN set NamaKaryawan='" & TextBox2.Text & "',AlamatKaryawan='" & TextBox3.Text & "',TelpKaryawan='" & TextBox4.Text & "' where KTPKaryawan='" & TextBox1.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil di Edit", MsgBoxStyle.Information, "Information")

                Call TampilGrid()

            Else
                MsgBox("Belum ada data KTP Karyawan, mohon cek ulang bosq!!!")
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Klik 2x, pada data yang ingin dihapus bosq!!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()

                Dim hapus As String = "delete from TBL_KARYAWAN where KTPKaryawan='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
                CMD.ExecuteNonQuery()

                Dim apus As String = "delete from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'"
                CMD = New SqlCommand(apus, CONN)
                CMD.ExecuteNonQuery()

                MsgBox("Data berhasil di Hapus", MsgBoxStyle.Information, "Information")

                Call TampilGrid()
            Else
                MsgBox("Data sudah dihapus bosq!!!")
            End If
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_KARYAWAN where NamaKaryawan like '%" & TextBox5.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_KARYAWAN where NamaKaryawan like '%" & TextBox5.Text & "%'", CONN)
            DS = New DataSet
            DA.Fill(DS, "ketemu")
            DataGridView1.DataSource = DS.Tables("ketemu")
            DataGridView1.ReadOnly = True
        End If
    End Sub
    Private Sub FormMasterDataKaryawan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

        Call TampilGrid()
        TextBox1.Focus()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub
End Class