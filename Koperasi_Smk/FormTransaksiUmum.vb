Imports System.Data.SqlClient
Public Class FormTransaksiUmum
    Private Sub kosongA()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox11.Clear()
        RichTextBox1.Clear()
        RichTextBox2.Clear()
        RichTextBox3.Clear()
    End Sub
    Private Sub umum()
        koneksi()
        str = "select * from trans_penjualan_umum order by kd_penjualan_umum desc"
        CMD = New SqlCommand(str, con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox8.Text = "TRU" + "001"
        Else
            TextBox8.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_penjualan_umum").ToString, 5, 2)) + 1
            If Len(TextBox8.Text) = 1 Then
                TextBox8.Text = "TRU00" & TextBox8.Text & ""
            Else
                TextBox8.Text = "TRU0" & TextBox8.Text & ""
            End If
        End If
    End Sub

    Private Sub rinci()
        koneksi()
        str = "select * from penjualan_rinci_umum order by kd_rinci_umum desc"
        CMD = New SqlCommand(str, con)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox12.Text = "UR" + "001"
        Else
            TextBox12.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_rinci_umum").ToString, 4, 2)) + 1
            If Len(TextBox12.Text) = 1 Then
                TextBox12.Text = "UR00" & TextBox12.Text & ""
            Else
                TextBox12.Text = "UR0" & TextBox12.Text & ""
            End If
        End If
    End Sub
    Private Sub FormTransaksiUmum_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormBorderStyle = FormBorderStyle.None
        kosongA()
        umum()
        rinci()
        TextBox6.Text = Format(Now, "dd/MM/yyyy")
        TextBox5.Text = FormMain.TextBox1.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormCariB.Show()
        Me.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        koneksi()
        CMD = New SqlCommand("select * from trans_penjualan_umum where kd_penjualan_umum='" & TextBox8.Text & "'", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            Dim strb As String
            strb = "insert into penjualan_rinci_umum values('" & TextBox8.Text & "','" & TextBox1.Text & "','" & TextBox12.Text & "','" & FormMain.TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & Val(TextBox3.Text) * Val(TextBox4.Text) & "')"
            If query(strb) Then
                Dim strstok As String
                strstok = "update tbl_barang set stok = '" & Val(TextBox11.Text) - Val(TextBox4.Text) & "' where kd_barang = '" & TextBox1.Text & "'"
                If query(strstok) Then
                    kosongA()
                    DA = New SqlDataAdapter("select * from penjualan_rinci_umum where kd_penjualan_umum = '" & TextBox8.Text & "' ", con)
                    Dim ds As New DataSet
                    Dim dt As DataTable
                    DA.Fill(ds)
                    For Each dt In ds.Tables
                        DataGridView1.DataSource = dt
                    Next
                End If
            End If
        Else
            str = "insert into trans_penjualan_umum values('" & TextBox8.Text & "','" & Format(Now, "yyyy/MM/dd") & "','" & RichTextBox1.Text & "')"
            If query(str) Then
                Dim strb As String
                strb = "insert into penjualan_rinci_umum values('" & TextBox8.Text & "','" & TextBox1.Text & "','" & TextBox12.Text & "','" & FormMain.TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & Val(TextBox3.Text) * Val(TextBox4.Text) & "')"
                If query(strb) Then
                    kosongA()

                    Dim strstok As String
                    strstok = "update tbl_barang set stok = '" & Val(TextBox11.Text) - Val(TextBox4.Text) & "' where kd_barang = '" & TextBox1.Text & "'"
                    If query(strstok) Then
                        DA = New SqlDataAdapter("select * from penjualan_rinci_umum where kd_penjualan_umum = '" & TextBox8.Text & "' ", con)
                        Dim ds As New DataSet
                        Dim dt As DataTable
                        DA.Fill(ds)
                        For Each dt In ds.Tables
                            DataGridView1.DataSource = dt
                        Next
                    End If
                End If
            End If

        End If
        rinci()
        koneksi()
        CMD = New SqlCommand("select sum(sub_total_umum) as total from penjualan_rinci_umum where kd_penjualan_umum = '" & TextBox8.Text & "'", con)
        RD = CMD.ExecuteReader
        RD.Read()
        If RD.HasRows Then
            RichTextBox1.Text = RD.Item("total")
        Else
            RichTextBox1.Text = 0
        End If

    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged
        RichTextBox3.Text = Val(RichTextBox2.Text) - Val(RichTextBox1.Text)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        str = "update trans_penjualan_umum set total_umum = '" & RichTextBox1.Text & "' where kd_penjualan_umum = '" & TextBox8.Text & "'"
        If query(str) Then
            MsgBox("berhasil")
            kosongA()
            umum()
            rinci()
            RichTextBox1.Text = "0"
            RichTextBox2.Clear()
            RichTextBox3.Clear()
            koneksi()
            DA = New SqlDataAdapter("select * from penjualan_rinci_umum where kd_penjualan_umum = '" & TextBox8.Text & "'", con)
            Dim ds As New DataSet
            Dim dt As New DataTable
            DA.Fill(ds)
            For Each dt In ds.Tables
                DataGridView1.DataSource = dt
            Next

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Visible = False
    End Sub
End Class