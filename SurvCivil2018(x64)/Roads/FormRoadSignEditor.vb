Public Class FormRoadSignEditor
    Public boolShow As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayMLeaderStyle As String() = {}
        Dim countMleader As Integer = AcadFunction.FuncReadListMLeaderStyle(arrayMLeaderStyle)
        ComboBox1.DataSource = arrayMLeaderStyle
        Dim arraylayers As String() = {"А-П_Дорожный_знак"}
        Dim countLayer As Integer = AcadFunction.FuncReadLayersDWG(arraylayers)
        ComboBox2.DataSource = arraylayers
        PictureBox1.SizeMode = Windows.Forms.PictureBoxSizeMode.Zoom
    End Sub
    Private Sub FormRoadSignEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'полилиния
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            ComboBox1.Enabled = False
        End If
    End Sub
    'нет
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            ComboBox1.Enabled = False
        End If
    End Sub
    'мультивыноска
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            ComboBox1.Enabled = True
        End If
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
End Class