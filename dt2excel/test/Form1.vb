Public Class Form1

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim _d2e As jp.polestar.io.dt2excel = New jp.polestar.io.dt2excel

        _d2e.ToDataTable("アカウント台帳.xlsx", "Sheet1", 15, 5)

    End Sub
End Class
