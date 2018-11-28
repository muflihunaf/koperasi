Imports System.Data.SqlClient
Public Class FormLogin

    Private Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub login()
        koneksi()
        CMD = New SqlCommand("select * from tbl_pengguna where nm_login = '" & TextBox1.Text & "' and pass_login = '" & TextBox2.Text & "'", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            FormMain.WindowState = FormWindowState.Maximized
            FormMain.TextBox1.Text = Me.TextBox1.Text
            FormMain.TextBox2.Text = RD.Item("kd_pengguna")
            FormMain.Show()
            Me.Visible = False

        Else
            MsgBox("User tidak Terdaftar")
        End If
    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        kosong()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Kurang lengkap")
            kosong()
        Else
            login()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        kosong()
    End Sub
End Class
