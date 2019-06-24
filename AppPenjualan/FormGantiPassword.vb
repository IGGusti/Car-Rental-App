Imports System.Data.SqlClient
Public Class FormGantiPassword

    'Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
    '   If e.KeyChar = Chr(100) Then TextBox2.Focus()
    'End Sub
    'Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
    '   If e.KeyChar = Chr(100) Then Button1.Focus()
    'End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap, mohon diisi semua ya bosq!!")
            Exit Sub
        Else
            Call Koneksi()
            CMD = New SqlCommand("select * from TBL_ADMIN where KodeAdmin='" & Label3.Text & "' and PasswordAdmin='" & TextBox1.Text & "' and LevelAdmin='USER'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Call Koneksi()
                Dim edit As String = "update TBL_ADMIN set PasswordAdmin='" & TextBox2.Text & "' where KodeAdmin='" & Label3.Text & "'"
                CMD = New SqlCommand(edit, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Password berhasil diganti ya bosq!!", MsgBoxStyle.Information, "Information")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.PasswordChar = "*"
                TextBox2.PasswordChar = "*"
            Else
                MsgBox("Password Lama salah, ingat-ingat lagi ya bosq!!")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox1.PasswordChar = "*"
                TextBox2.PasswordChar = "*"
            End If
        End If

    End Sub

    Private Sub FormGantiPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox1.PasswordChar = "*"
        TextBox2.PasswordChar = "*"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class