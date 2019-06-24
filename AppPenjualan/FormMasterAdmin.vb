Imports System.Data.SqlClient
Public Class FormMasterAdmin
    'Untuk Memunculkan di ComboBox
    Sub TampilStatus()
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Items.Add("USER")
        ComboBox1.SelectedItem = "ADMIN"
    End Sub
    'Untuk Menampilkan Data di DataGridView
    Sub TampilGrid()
        Call Koneksi()
        DA = New SqlDataAdapter("select kodeAdmin as Kode,NamaAdmin as Nama,PasswordAdmin as Password,LevelAdmin as Level From tbl_admin", CONN)
        DS = New DataSet
        DA.Fill(DS, "tbl_Admin")
        DataGridView1.DataSource = DS.Tables("tbl_Admin")
        DataGridView1.ReadOnly = True
    End Sub
    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.RowIndex > DataGridView1.Rows.Count() OrElse e.RowIndex < 0 OrElse DataGridView1.Rows(e.RowIndex) Is Nothing OrElse DataGridView1.Rows(e.RowIndex).DataBoundItem Is Nothing Then
            e.Value = Nothing
        ElseIf e.ColumnIndex = 2 Then
            e.Value = "*****"
        End If
    End Sub
    Private Sub FormMasterAdmin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""

        Call TampilStatus()
        Call TampilGrid()
        TextBox1.Focus()
        TextBox3.PasswordChar = "*"
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                MsgBox("Kode Admin sudah terpakai, mohon cek ulang bosq!!")
            Else
                Call Koneksi()
                Dim simpan As String = "insert into TBL_ADMIN values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
                CMD = New SqlCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil di Input", MsgBoxStyle.Information, "Information")

                Call TampilGrid()

                If ComboBox1.Text = "USER" Then
                    Call Koneksi()
                    Dim simpann As String = "insert into TBL_KARYAWAN values ('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                    CMD = New SqlCommand(simpann, CONN)
                    CMD.ExecuteNonQuery()
                End If
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Call Koneksi()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        CMD = New SqlCommand("Select * from TBL_ADMIN where KodeAdmin='" & DataGridView1.Item(0, i).Value & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Focus()
        Else
            TextBox1.Text = RD.Item("KodeAdmin")
            TextBox2.Text = RD.Item("NamaAdmin")
            TextBox3.Text = RD.Item("PasswordAdmin")
            ComboBox1.Text = RD.Item("LevelAdmin")
            TextBox2.Focus()
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                Dim edit As String = "update TBL_ADMIN set NamaAdmin='" & TextBox2.Text & "',PasswordAdmin='" & TextBox3.Text & "',LevelAdmin='" & ComboBox1.Text & "' where KodeAdmin='" & TextBox1.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data berhasil di Edit", MsgBoxStyle.Information, "Information")
                Call TampilGrid()

                If ComboBox1.Text = "USER" Then
                    Dim editt As String = "update TBL_KARYAWAN set NamaAdmin='" & TextBox2.Text & "' where KodeAdmin='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(editt, CONN)
                    CMD.ExecuteNonQuery()
                End If
            Else
                MsgBox("Belum ada data Kode Admin, mohon cek ulang bosq!!!")
            End If
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Klik 2x, pada data yang ingin dihapus bosq!!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                If ComboBox1.Text = "USER" Then
                    Dim hapuss As String = "delete from TBL_KARYAWAN where KodeAdmin='" & TextBox1.Text & "'"
                    CMD = New SqlCommand(hapuss, CONN)
                    CMD.ExecuteNonQuery()
                End If

                Call Koneksi()
                Dim hapus As String = "delete from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "'"
                CMD = New SqlCommand(hapus, CONN)
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
    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        Call Koneksi()
        CMD = New SqlCommand("select * from TBL_ADMIN where NamaAdmin like '%" & TextBox4.Text & "%'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Call Koneksi()
            DA = New SqlDataAdapter("select * from TBL_ADMIN where NamaAdmin like '%" & TextBox4.Text & "%'", CONN)
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
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class