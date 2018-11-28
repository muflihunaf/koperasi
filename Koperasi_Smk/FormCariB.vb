Imports System.Data.SqlClient
Public Class FormCariB
    Private Sub FormCariB_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()
        FormBorderStyle = FormBorderStyle.None
        tampil("tbl_barang", DataGridView1)
    End Sub

    Private Sub isidata(ByVal i As Integer)
        FormTransaksiUmum.TextBox1.Text = DataGridView1.Rows(i).Cells(0).Value
        FormTransaksiUmum.TextBox2.Text = DataGridView1.Rows(i).Cells(2).Value
        FormTransaksiUmum.TextBox3.Text = DataGridView1.Rows(i).Cells(4).Value
        FormTransaksiUmum.TextBox11.Text = DataGridView1.Rows(i).Cells(5).Value
        FormTransaksiUmum.Enabled = True
        Me.Visible = False

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        koneksi()
        DA = New SqlDataAdapter("select * from tbl_barang where nm_barang like '%" & TextBox1.Text & "%'", con)
        Dim ds As New DataSet
        DA.Fill(ds)
        Dim dt As DataTable
        For Each dt In ds.Tables
            DataGridView1.DataSource = dt
        Next
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        isidata(e.RowIndex)
    End Sub

End Class