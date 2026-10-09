Public Class FormOverDrawPlan
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayFormat As String() = {"Полный кадастровый номер", "Краткий кадастровый номер"}
        ' Добавить код инициализации после вызова InitializeComponent().
        ComboBox1.DataSource = arrayFormat

        Dim arrayTextStyles As String() = {""}
        Dim boolfindtextstyle As Integer = AcadFunction.FuncReadListTextStyle(arrayTextStyles)
        ComboBox2.DataSource = arrayTextStyles
        Dim boolFindStuleText As Boolean = MathFunction.FuncFindValueToArray(arrayTextStyles, "CADASTRAL_OBJECT")
        If boolFindStuleText = True Then
            ComboBox2.Text = "CADASTRAL_OBJECT"
        End If
    End Sub
    Private Sub FuncOverDrawPlan_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class