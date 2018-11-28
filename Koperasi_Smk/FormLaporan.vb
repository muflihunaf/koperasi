Imports System.Data.SqlClient
Public Class FormLaporan
    Dim tinggi As Double
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        koneksi()
        Dim ds As New DataSet
        DA = New SqlDataAdapter("select * from trans_penjualan_umum where tgl_trans_umum  between '" & DateTimePicker1.Value & "' and '" & DateTimePicker2.Value & "' order by kd_penjualan_umum desc ", con)
        DA.Fill(ds)
        Dim dt As DataTable
        For Each dt In ds.Tables
            DataGridView1.DataSource = dt
        Next
    End Sub

    Private Sub FormLaporan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        koneksi()

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawString("Laporan Penjualan", New Drawing.Font("Arial", 20), Brushes.Black, 4, tinggi)
        tinggi += 50
        e.Graphics.DrawString("Kode Penjualan", New Drawing.Font("Arial", 10), Brushes.Black, 30, tinggi)
        e.Graphics.DrawString("Tanggal Penjualan", New Drawing.Font("Arial", 10), Brushes.Black, 200, tinggi)
        e.Graphics.DrawString("Total Umum", New Drawing.Font("Arial", 10), Brushes.Black, 340, tinggi)
        tinggi += 10
        For baris As Integer = 0 To DataGridView1.RowCount - 2
            tinggi += 15
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(0).Value.ToString, New Drawing.Font("Arial", 8), Brushes.Black, 30, tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(1).Value.ToString, New Drawing.Font("Arial", 8), Brushes.Black, 200, tinggi)
            e.Graphics.DrawString(DataGridView1.Rows(baris).Cells(2).Value.ToString, New Drawing.Font("Arial", 8), Brushes.Black, 340, tinggi)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub
End Class