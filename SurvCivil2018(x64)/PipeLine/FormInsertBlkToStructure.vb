Public Class FormInsertBlkToStructure
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayBlk As String() = {""}
        Dim numBlk As Integer = AcadFunction.FuncReadListBlockDef(arrayBlk)
        ' Добавить код инициализации после вызова InitializeComponent().
        ComboBox1.DataSource = arrayBlk
    End Sub
    Private Sub FormInsertBlkToStructure_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class