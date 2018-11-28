Imports System.Data.SqlClient
Public Class FormBarang

    Private Sub kosong()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox9.ReadOnly = True
        TextBox2.ReadOnly = True
        TextBox3.Focus()
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        tampil("tbl_barang", DataGridView1)

    End Sub
    Private Sub kode()
        koneksi()
        CMD = New SqlCommand("select * from tbl_barang order by kd_barang desc", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Text = "BRG" + "001"
        Else
            TextBox2.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_barang").ToString, 5, 2)) + 1
            If Len(TextBox2.Text) = 1 Then
                TextBox2.Text = "BRG00" & TextBox2.Text & ""
            Else
                TextBox2.Text = "BRG0" & TextBox2.Text & ""
            End If
        End If
    End Sub

    Private Sub isidata(ByVal i As Integer)
        TextBox2.Text = DataGridView1.Rows(i).Cells(0).Value
        TextBox3.Text = DataGridView1.Rows(i).Cells(2).Value
        TextBox4.Text = DataGridView1.Rows(i).Cells(3).Value
        TextBox5.Text = DataGridView1.Rows(i).Cells(4).Value
        TextBox9.Text = DataGridView1.Rows(i).Cells(1).Value
        TextBox7.Text = DataGridView1.Rows(i).Cells(5).Value
        TextBox8.Text = DataGridView1.Rows(i).Cells(6).Value

    End Sub
    Private Sub simpan()
        str = "insert into tbl_barang values('" & TextBox2.Text & "','" & TextBox9.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
        End If
    End Sub
    Private Sub hapus()
        str = "delete from tbl_barang where kd_barang = '" & TextBox2.Text & "'"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
        End If
    End Sub
    Private Sub ubah()
        str = "update tbl_barang set kd_jenis_brg = '" & TextBox9.Text & "',nm_barang = '" & TextBox3.Text & "',hrg_beli= '" & TextBox4.Text & "',hrg_jual= '" & TextBox5.Text & "',stok= '" & TextBox7.Text & "',keterangan= '" & TextBox8.Text & "' where kd_barang = '" & TextBox2.Text & "'"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
        End If
    End Sub
    Private Sub FormBarang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        kosong()
        kode()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("data Belum Lengkap")
        Else
            simpan()
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        koneksi()
        RD.Close()
        CMD = New SqlCommand("select * from tbl_jenis_barang where jenis_brg = '" & TextBox6.Text & "'", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox6.Text = RD.Item("jenis_brg")
            TextBox9.Text = RD.Item("kd_jenis_brg")
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
        RD.Close()
        CMD = New SqlCommand("select * from tbl_jenis_barang where kd_jenis_brg = '" & TextBox9.Text & "'", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox6.Text = RD.Item("jenis_brg")
        End If
        Button4.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MsgBox("Pilih Data Terlebih Dahulu")
        Else
            If MsgBox("Apa Anda Yakin?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                hapus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("data Belum Lengkap")
        Else
            ubah()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = True
        Button4.Enabled = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        koneksi()
        DA = New SqlDataAdapter("select * from tbl_barang where nm_barang like  '%" & TextBox1.Text & "%'", con)
        Dim ds As New DataSet
        DA.Fill(ds)
        Dim dt As New DataTable
        For Each dt In ds.Tables
            DataGridView1.DataSource = dt
        Next

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub
End Class