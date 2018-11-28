Imports System.Data.SqlClient
Public Class FormJenisBarang
    Private Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.ReadOnly = True
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        tampil("tbl_jenis_barang", DataGridView1)
    End Sub

    Private Sub isidata(ByVal i As Integer)
        TextBox1.Text = DataGridView1.Rows(i).Cells(0).Value
        TextBox2.Text = DataGridView1.Rows(i).Cells(1).Value
        Button1.Enabled = False
        Button4.Enabled = False
        Button3.Enabled = True
        Button2.Enabled = True
    End Sub
    Private Sub kode()
        koneksi()
        CMD = New SqlCommand("select * from tbl_jenis_barang order by kd_jenis_brg desc", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "JNS" + "001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_jenis_brg").ToString, 5, 2)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "JNS00" & TextBox1.Text & ""
            Else
                TextBox1.Text = "JNS0" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Private Sub FormJenisBarang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        kosong()
        kode()
    End Sub
    Private Sub simpan()
        str = "insert into tbl_jenis_barang values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
        If query(str) Then
            MsgBox("Berhasil Simpan Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub ubah()
        str = "update tbl_jenis_barang set jenis_brg ='" & TextBox2.Text & "' where kd_jenis_brg = '" & TextBox1.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Update Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub hapus()
        str = "delete from tbl_jenis_barang where kd_jenis_brg = '" & TextBox1.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Simpan Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("Simpan ")
        Else
            simpan()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data Belum Lengkap")
        Else
            ubah()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("pilih Data Terlebih dahulu")
        Else
            If MsgBox("Anda Yakin Hapus Data Ini?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                hapus()
            End If

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub
End Class