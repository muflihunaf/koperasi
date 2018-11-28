Imports System.Data.SqlClient
Public Class FormAnggota

    Private Sub kosong()
        koneksi()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox2.ReadOnly = True
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Text = Format(Now, "dd/MM/yyyy")
        TextBox8.ReadOnly = True
        TextBox9.Enabled = False
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("Laki-Laki")
        ComboBox1.Items.Add("Perempuan")
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = True
        tampil("tbl_anggota", DataGridView1)
    End Sub
    Private Sub kode()
        koneksi()
        CMD = New SqlCommand("select * from tbl_anggota order by kd_anggota desc", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox2.Text = "ANG" + "001"
        Else
            TextBox2.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_anggota").ToString, 5, 2)) + 1
            If Len(TextBox2.Text) = 1 Then
                TextBox2.Text = "ANG00" & TextBox2.Text & ""
            Else
                TextBox2.Text = "ANG0" & TextBox2.Text & ""
            End If
        End If
    End Sub
    Private Sub isidata(ByVal i As Integer)
        TextBox2.Text = DataGridView1.Rows(i).Cells(0).Value
        TextBox9.Text = DataGridView1.Rows(i).Cells(1).Value
        TextBox4.Text = DataGridView1.Rows(i).Cells(2).Value
        TextBox5.Text = DataGridView1.Rows(i).Cells(3).Value
        TextBox6.Text = DataGridView1.Rows(i).Cells(4).Value
        DateTimePicker1.Text = DataGridView1.Rows(i).Cells(5).Value
        ComboBox1.Text = DataGridView1.Rows(i).Cells(6).Value
        TextBox7.Text = DataGridView1.Rows(i).Cells(7).Value
        TextBox8.Text = DataGridView1.Rows(i).Cells(8).Value
        str = "select * from tbl_unit_kerja where kd_unit_kerja = '" & TextBox9.Text & "'"
        RD.Close()
        CMD = New SqlCommand(str, con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox3.Text = RD.Item("unit_kerja")
        End If
    End Sub
    Private Sub simpan()
        koneksi()
        str = "insert into tbl_anggota values('" & TextBox2.Text & "','" & TextBox9.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & DateTimePicker1.Text & "','" & ComboBox1.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
        If query(str) Then
            MsgBox("Berhasil")
            kosong()
            kode()
        End If
    End Sub
    Private Sub ubah()
        koneksi()
        str = "update tbl_anggota set kd_unit_kerja = '" & TextBox9.Text & "',npp = '" & TextBox4.Text & "',nm_anggota = '" & TextBox5.Text & "',tmp_lahir = '" & TextBox6.Text & "', tgl_lahir = '" & DateTimePicker1.Text & "',jenis_kelamin = '" & ComboBox1.Text & "', alamat = '" & TextBox7.Text & "' where kd_anggota = '" & TextBox2.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Ubah Data")
            kosong()
            kode()
        End If
    End Sub
    Private Sub hapus()
        koneksi()
        str = "delete from tbl_anggota where kd_anggota = '" & TextBox2.Text & "'"
        If query(str) Then
            MsgBox("Berhasil Hapus Data")
            kosong()
            kode()
        End If
    End Sub

    Private Sub FormAnggota_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        kosong()
        kode()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or ComboBox1.Text = "" Or DateTimePicker1.Text = "" Then
            MsgBox("Data Belum Lengkap")
        Else
            simpan()
        End If

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        koneksi()
        str = "select * from tbl_unit_kerja where unit_kerja = '" & TextBox3.Text & "'"
        CMD = New SqlCommand(str, con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            TextBox3.Text = RD.Item("unit_kerja")
            TextBox9.Text = RD.Item("kd_unit_kerja")
        End If

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
        Button2.Enabled = True
        Button3.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MsgBox("Pilih Data Terlebih Dahulu")
        Else
            hapus()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Or ComboBox1.Text = "" Or DateTimePicker1.Text = "" Then
            MsgBox("Data Belum Lengkap")
        Else
            ubah()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        koneksi()
        DA = New SqlDataAdapter("select * from tbl_anggota where nm_anggota like '%" & TextBox1.Text & "%'", con)
        Dim ds As New DataSet
        DA.Fill(ds)
        Dim dt As New DataTable
        For Each dt In ds.Tables
            DataGridView1.DataSource = dt
        Next

    End Sub
End Class