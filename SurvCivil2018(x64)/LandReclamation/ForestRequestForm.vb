Imports System.Windows.Forms
Public Class ForestRequestForm
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
    End Sub
    Public Property ResultData As Dictionary(Of String, String)

    Private Sub ForestRequestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' === Лесной участок ===
        SetupGrid(dgvForestPart, {
            ("Action", "updateForestStead"),
            ("numberForestPart", ""),
            ("cadNumberForestPart", "")
        })

        ' === ОКС ===
        SetupGrid(dgvOKS, {
            ("Action", "oksBorderForestry"),
            ("numberOKS", ""),
            ("cadNumberOKS", ""),
            ("numberForestry", ""),
            ("numberLandPlot", "")
        })

        ' === Лесничество ===
        SetupGrid(dgvForestry, {
            ("Action", "borderUnchanged"),
            ("numberForestry", ""),
            ("registryNumber", "")
        })

        ' === ЗОУИТ ===
        SetupGrid(dgvSpecialZone, {
            ("Action", "zouitPutInfo"),
            ("registryNumber", "")
        })
    End Sub

    ''' <summary>
    ''' Заполняет DataGridView парами (поле, значение по умолчанию).
    ''' Первая колонка — только для чтения.
    ''' </summary>
    Private Sub SetupGrid(grid As DataGridView, fields As (String, String)())
        Dim dt As New DataTable()
        dt.Columns.Add("Поле", GetType(String))
        dt.Columns.Add("Значение", GetType(String))
        For Each f In fields
            dt.Rows.Add(f.Item1, f.Item2)
        Next
        grid.DataSource = dt
        grid.Columns(0).ReadOnly = True
        grid.Columns(0).Width = 200
        grid.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        ResultData = New Dictionary(Of String, String)()

        ' Для каждого блока: если хотя бы одно поле заполнено — считаем блок активным
        Dim forestPartFilled = CollectGrid(dgvForestPart, "ForestPart")
        Dim oksFilled = CollectGrid(dgvOKS, "OKS")
        Dim forestryFilled = CollectGrid(dgvForestry, "Forestry")
        Dim specialZoneFilled = CollectGrid(dgvSpecialZone, "SpecialZone")

        ' Валидация: хотя бы один блок должен быть заполнен
        If Not forestPartFilled AndAlso Not oksFilled AndAlso Not forestryFilled AndAlso Not specialZoneFilled Then
            MessageBox.Show("Заполните хотя бы один блок.", "Ошибка",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Валидация: в активном блоке обязательные поля не пустые
        Dim errMsg As String = Nothing

        If forestPartFilled Then
            errMsg = CheckRequired(dgvForestPart, {"Action", "numberForestPart", "cadNumberForestPart"})
            If errMsg IsNot Nothing Then
                MessageBox.Show("Лесной участок: " & errMsg, "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        If oksFilled Then
            errMsg = CheckRequired(dgvOKS, {"Action", "numberOKS", "cadNumberOKS"})
            If errMsg IsNot Nothing Then
                MessageBox.Show("ОКС: " & errMsg, "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        If forestryFilled Then
            errMsg = CheckRequired(dgvForestry, {"Action", "numberForestry", "registryNumber"})
            If errMsg IsNot Nothing Then
                MessageBox.Show("Лесничество: " & errMsg, "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        If specialZoneFilled Then
            errMsg = CheckRequired(dgvSpecialZone, {"Action", "registryNumber"})
            If errMsg IsNot Nothing Then
                MessageBox.Show("ЗОУИТ: " & errMsg, "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ''' <summary>
    ''' Собирает значения из grid в ResultData с префиксом блока.
    ''' Возвращает True, если хотя бы одно поле заполнено.
    ''' </summary>
    Private Function CollectGrid(grid As DataGridView, blockName As String) As Boolean
        Dim anyFilled As Boolean = False
        For Each row As DataGridViewRow In grid.Rows
            If row.IsNewRow Then Continue For
            Dim key = blockName & "." & CStr(row.Cells(0).Value)
            Dim value = If(row.Cells(1).Value Is Nothing, "", CStr(row.Cells(1).Value))
            ResultData(key) = value
            If Not String.IsNullOrWhiteSpace(value) Then anyFilled = True
        Next
        Return anyFilled
    End Function

    ''' <summary>
    ''' Проверяет, что все обязательные поля заполнены.
    ''' Возвращает текст ошибки или Nothing.
    ''' </summary>
    Private Function CheckRequired(grid As DataGridView, requiredFields As String()) As String
        Dim missing As New List(Of String)()
        For Each row As DataGridViewRow In grid.Rows
            If row.IsNewRow Then Continue For
            Dim key = CStr(row.Cells(0).Value)
            If requiredFields.Contains(key) Then
                Dim value = If(row.Cells(1).Value Is Nothing, "", CStr(row.Cells(1).Value))
                If String.IsNullOrWhiteSpace(value) Then missing.Add(key)
            End If
        Next
        If missing.Count > 0 Then Return "не заполнены поля: " & String.Join(", ", missing)
        Return Nothing
    End Function

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class