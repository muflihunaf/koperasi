Imports System.Data.SqlClient
Public Class FormMain
    Private Sub AnggotaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnggotaToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormAnggota.MdiParent = Me
        Me.Panel1.Controls.Add(FormAnggota)
        FormAnggota.WindowState = FormWindowState.Maximized
        FormAnggota.Show()
    End Sub

    Private Sub PenggunaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PenggunaToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormPengguna.MdiParent = Me
        Me.Panel1.Controls.Add(FormPengguna)
        FormPengguna.WindowState = FormWindowState.Maximized
        FormPengguna.Show()

    End Sub



    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub UnitKerjaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnitKerjaToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormUnitKerja.MdiParent = Me
        Me.Panel1.Controls.Add(FormUnitKerja)
        FormUnitKerja.WindowState = FormWindowState.Maximized
        FormUnitKerja.Show()
    End Sub

    Private Sub JenisBarangToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles JenisBarangToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormJenisBarang.MdiParent = Me
        Me.Panel1.Controls.Add(FormJenisBarang)
        FormJenisBarang.WindowState = FormWindowState.Maximized
        FormJenisBarang.Show()
    End Sub

    Private Sub BarangToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BarangToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormBarang.MdiParent = Me
        Me.Panel1.Controls.Add(FormBarang)
        FormBarang.WindowState = FormWindowState.Maximized
        FormBarang.Show()
    End Sub

    Private Sub PenjualanUmumToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PenjualanUmumToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormTransaksiUmum.MdiParent = Me
        Me.Panel1.Controls.Add(FormTransaksiUmum)
        FormTransaksiUmum.WindowState = FormWindowState.Maximized
        FormTransaksiUmum.Show()
    End Sub

    Private Sub PenjualanAnggotaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PenjualanAnggotaToolStripMenuItem.Click
        Me.Panel1.Controls.Clear()
        FormTransaksiAnggota.MdiParent = Me
        Panel1.Controls.Add(FormTransaksiAnggota)
        FormTransaksiAnggota.WindowState = FormWindowState.Maximized
        FormTransaksiAnggota.Show()
    End Sub

    Private Sub LaporanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaporanToolStripMenuItem.Click
        FormLaporan.Show()
    End Sub
End Class