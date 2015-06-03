Public Class UserControl1

    Private Sub UserControl1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim _d2t As jp.polestar.io.dt2excel = New jp.polestar.io.dt2excel
        _d2t.ToDataTable("アカウント台帳.xlsx", "Sheet1", 10)

    End Sub
End Class
