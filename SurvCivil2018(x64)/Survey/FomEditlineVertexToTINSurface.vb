Public Class FomEditlineVertexToTINSurface
    Public boolShow As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
    End Sub
    Private Sub FomEditlineVertexToTINSurface_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    '==================================================================================
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        boolShow = True
    End Sub
    '===============================================================================
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        boolShow = False
    End Sub
End Class