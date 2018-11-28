Imports System.Data.SqlClient
Public Class FormUnitKerja
    Private Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        tampil("tbl_unit_kerja", DataGridView1)
    End Sub

    Private Sub isidata(ByVal i As Integer)
        TextBox1.Text = DataGridView1.Rows(i).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(i).Cells(1).Value
    End Sub
    Private Sub kode()
        koneksi()
        CMD = New SqlCommand("select * from tbl_unit_kerja order by kd_unit_kerja desc", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "UNT" + "001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_unit_kerja").ToString, 5, 2)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "UNT00" & TextBox1.Text & ""
            Else
                TextBox1.Text = "UNT0" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Private Sub simpan()
        koneksi()
        str = "insert into tbl_unit_kerja values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
        If query(str) Then
            MsgBox("Berhasil Simpan Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub ubah()
        koneksi()
        str = "update tbl_unit_kerja set unit_kerja = '" & TextBox2.Text & "' where kd_unit_kerja = '" & TextBox1.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Ubah Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub hapus()
        koneksi()
        str = "delete from tbl_unit_kerja where kd_unit_kerja = '" & TextBox1.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Hapus Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub FormUnitKerja_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        kosong()
        kode()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("Data Kurang Lengkap")
        Else
            simpan()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Or TextBox1.Text = "" Then
            MsgBox("Data Kurang Lengkap")
        Else
            ubah()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data Dulu")

        Else
            If MsgBox("Anda Yakin?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                hapus()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub
End Class