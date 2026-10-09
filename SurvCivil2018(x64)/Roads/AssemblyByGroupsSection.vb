Public Class AssemblyByGroupsSection
    Public SelectName As String = ""
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SelectName = ListBox1.SelectedItem.ToString
        Me.Hide()
    End Sub
End Class