Imports System.Data.SqlClient
Public Class FormLogin

    Dim berhasil As String

    'Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
    '   If e.KeyChar = Chr(100) Then TextBox2.Focus()
    'End Sub
    'Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
    '  If e.KeyChar = Chr(100) Then Button1.Focus()
    'End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data login belum lengkap, lengkapi dulu ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "' and PasswordAdmin='" & TextBox2.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & TextBox1.Text & "' and PasswordAdmin='" & TextBox2.Text & "' and LevelAdmin='USER'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                If RD.HasRows Then
                    FormMenuUtama.Label22.Text = "Welcome, " & RD.Item("NamaAdmin")
                    FormMenuUtama.Button1.Enabled = False
                    FormMenuUtama.LoginToolStripMenuItem.Enabled = False
                    FormMenuUtama.LogoutToolStripMenuItem.Enabled = True
                    FormMenuUtama.KeluarToolStripMenuItem.Enabled = True
                    FormMenuUtama.MasterToolStripMenuItem.Enabled = True
                    FormMenuUtama.AdminToolStripMenuItem.Enabled = False
                    FormMenuUtama.DataKaryawanToolStripMenuItem.Enabled = False
                    FormMenuUtama.DataMobiToolStripMenuItem.Enabled = False
                    FormMenuUtama.TransaksiToolStripMenuItem.Enabled = True
                    FormMenuUtama.LaporanToolStripMenuItem.Enabled = True
                    FormMenuUtama.UtilityToolStripMenuItem.Enabled = True
                    FormMenuUtama.GroupBox1.Enabled = True
                    FormMenuUtama.GroupBox2.Enabled = True
                    FormMenuUtama.GroupBox3.Enabled = True
                    FormMenuUtama.GroupBox4.Enabled = True
                    FormMenuUtama.Button2.Enabled = True
                Else
                    FormMenuUtama.Label22.Text = "Welcome, Admin"
                    Call Terbuka()
                End If
                FormGantiPassword.Label3.Text = TextBox1.Text
                Me.Close()
            Else
                MsgBox("Kode Admin atau Password salah, ingat-ingat lagi ya bosq!!")
            End If
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    Sub Terbuka()
        FormMenuUtama.Button1.Enabled = False
        FormMenuUtama.MasterToolStripMenuItem.Enabled = True
        FormMenuUtama.LogoutToolStripMenuItem.Enabled = True
        FormMenuUtama.TransaksiToolStripMenuItem.Enabled = True
        FormMenuUtama.LaporanToolStripMenuItem.Enabled = True
        FormMenuUtama.UtilityToolStripMenuItem.Enabled = True
        FormMenuUtama.LoginToolStripMenuItem.Enabled = False
        FormMenuUtama.GroupBox1.Enabled = True
        FormMenuUtama.GroupBox2.Enabled = True
        FormMenuUtama.GroupBox3.Enabled = True
        FormMenuUtama.GroupBox4.Enabled = True
        FormMenuUtama.Button2.Enabled = True
    End Sub
    Private Sub FormLogin_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        TextBox1.Focus()
    End Sub
    Private Sub FormLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox2.PasswordChar = "*"
    End Sub

End Class