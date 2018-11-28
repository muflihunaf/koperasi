Imports System.Data.SqlClient
Public Class FormPengguna
    Private Sub awal()
        koneksi()
        kosong()
        kode()
        tampil("tbl_pengguna", DataGridView1)

    End Sub

    Private Sub isidata(ByVal i As Integer)
        TextBox1.Text = DataGridView1.Rows(i).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(i).Cells(1).Value
        TextBox3.Text = DataGridView1.Rows(i).Cells(2).Value
        TextBox4.Text = DataGridView1.Rows(i).Cells(3).Value
        ComboBox1.Text = DataGridView1.Rows(i).Cells(4).Value
    End Sub

    Private Sub kosong()
        koneksi()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.Items.Add("")
        TextBox1.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True

    End Sub

    Private Sub simpan()
        koneksi()
        str = "insert into tbl_pengguna values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
            tampil("tbl_pengguna", DataGridView1)
        End If
    End Sub
    Private Sub hapus()
        koneksi()
        str = "delete from tbl_pengguna where kd_pengguna = '" & TextBox1.Text & "'"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
            tampil("tbl_pengguna", DataGridView1)
        End If
    End Sub
    Private Sub ubah()
        koneksi()
        str = "update tbl_pengguna set nm_pengguna = '" & TextBox2.Text & "',nm_login = '" & TextBox3.Text & "',pass_login = '" & TextBox4.Text & "',level = '" & ComboBox1.Text & "' "
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
            tampil("tbl_pengguna", DataGridView1)
        End If
    End Sub

    Private Sub kode()
        koneksi()
        CMD = New SqlCommand("select * from tbl_pengguna order by kd_pengguna desc", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "USR" + "001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_pengguna").ToString, 5, 2)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "USR00" & TextBox1.Text & ""
            Else
                TextBox1.Text = "USR0" & TextBox1.Text & ""
            End If

        End If

    End Sub

    Private Sub FormPengguna_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        kosong()
        kode()
        tampil("tbl_pengguna", DataGridView1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum Lengkap")
        Else
            simpan()
        End If

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Data Belum Lengkap")
        Else
            ubah()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data Terlebih Dahulu")
        Else
            If MsgBox("Anda Yakin? ", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                hapus()
            End If

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub
End Class