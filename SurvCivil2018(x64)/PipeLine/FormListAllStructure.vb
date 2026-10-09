Public Class FormListAllStructure
    Public boolButton As Boolean = False
    Public ind As Integer = 0
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolButton = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolButton = True
        Me.Hide()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim indrow As Integer = e.RowIndex
        Dim indcol As Integer = e.ColumnIndex
        If indrow > -1 And indcol = 0 Then
            Dim valf As Integer = DataGridView1.Rows(indrow).Cells(indcol).Value
            If valf = 1 Then
                Dim TempVAl As String = DataGridView1.Rows(indrow).Cells(3).Value
                If IsNothing(TempVAl) = True OrElse Len(TempVAl.Trim) = 0 Then
                    DataGridView1.Rows(indrow).Cells(3).Value = ind + 1
                    ind += 1
                End If
            ElseIf valf = 0 Then
                DataGridView1.Rows(indrow).Cells(3).Value = ""
            End If
        End If
    End Sub
End Class