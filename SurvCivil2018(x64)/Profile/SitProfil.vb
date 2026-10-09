Public Class SitProfil
    Inherits System.Windows.Forms.Form
    Public stbool As Boolean = False
    'начальные установки при инициализации формы
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        Dim ListLayer As String() = Nothing
        Dim LayerCount As Integer = 0 'SurvFunction.ListLayerGeo2(ListLayer)
        Dim ListLayer2 As String() = Nothing
        ReDim Preserve ListLayer2(0)
        ListLayer2(0) = "PROF-Продольный профиль (элементы плана)"
        For i As Integer = 1 To LayerCount - 1
            ReDim Preserve ListLayer2(i)
            ListLayer2(i) = ListLayer(i - 1)
        Next i
        ComboBox1.DataSource = ListLayer2
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        stbool = False
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        stbool = True
        Me.Hide()
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        Label7.Text = "Ширина захвата: " & (NumericUpDown2.Value * NumericUpDown3.Value * NumericUpDown4.Value) / 100
    End Sub


    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        Label7.Text = "Ширина захвата: " & (NumericUpDown2.Value * NumericUpDown3.Value * NumericUpDown4.Value) / 100
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        Label7.Text = "Ширина захвата: " & (NumericUpDown2.Value * NumericUpDown3.Value * NumericUpDown4.Value) / 100
    End Sub

    Private Sub SitProfil_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class