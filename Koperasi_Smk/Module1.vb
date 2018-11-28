Imports System.Data.SqlClient
Imports System.Data
Module Module1
    Public DA As SqlDataAdapter
    Public RD As SqlDataReader
    Public con As SqlConnection
    Public CMD As SqlCommand
    Public str As String
    Public Sub koneksi()
        str = "data source=Muflihun; initial catalog=koperasi; integrated security=true"
        con = New SqlConnection(str)
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
    End Sub

    Public Function query(ByVal sql As String)
        koneksi()
        Try
            CMD = New SqlCommand
            CMD.Connection = con
            CMD.CommandType = CommandType.Text
            CMD.CommandText = sql
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            con.Close()
            Return True
        Catch ex As Exception
            MsgBox("terjadi kesalahan")
            CMD.Dispose()
            con.Close()
            Return False
        Finally
            CMD.Dispose()
            con.Close()
        End Try

    End Function

    Public Sub tampil(ByVal tbl As String, ByVal dg As DataGridView)
        koneksi()

        Dim ds As New DataSet
        DA = New SqlDataAdapter("select * from " & tbl, con)
        DA.Fill(ds)
        Dim dt As New DataTable
        For Each dt In ds.Tables
            dg.DataSource = dt
        Next


    End Sub

End Module
