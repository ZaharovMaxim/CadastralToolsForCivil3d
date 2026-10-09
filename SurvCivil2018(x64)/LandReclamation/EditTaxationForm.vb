Imports System.Windows.Forms
Imports System.IO
Public Class EditTaxationForm
    Private Const SchemasPath As String = "k:\Civil3DToolsUtility\NET\Schemas\"
    Private Function LoadEnumFromXsd(xsdFile As String, typeName As String) As List(Of ComboItem)
        Dim result As New List(Of ComboItem)
        Dim doc As New Xml.XmlDocument()
        doc.Load(xsdFile)

        Dim ns As New Xml.XmlNamespaceManager(doc.NameTable)
        ns.AddNamespace("xs", "http://www.w3.org/2001/XMLSchema")

        Dim nodes = doc.SelectNodes(
        $"//xs:simpleType[@name='{typeName}']/xs:restriction/xs:enumeration", ns)

        For Each node As Xml.XmlNode In nodes
            Dim code As String = node.Attributes("value").Value
            Dim docNode = node.SelectSingleNode("xs:annotation/xs:documentation", ns)
            Dim descr As String = If(docNode IsNot Nothing, docNode.InnerText.Trim(), code)
            result.Add(New ComboItem(code, descr))
        Next

        Return result
    End Function


    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub EditTaxationForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Участковые лесничества
        Dim subforestries = LoadEnumFromXsd(
            IO.Path.Combine(SchemasPath, "common3.4\dDistrictForestriesKindTypes-2.1.5.xsd"),
            "DistrictForestriesKindEType")

        cmbSubforestry.DataSource = subforestries
        cmbSubforestry.DisplayMember = "Description"
        cmbSubforestry.ValueMember = "Code"

        ' Целевое назначение лесов
        Dim purposes = LoadEnumFromXsd(
            IO.Path.Combine(SchemasPath, "common3.4\dPurposeForestsKindTypes-2.1.2.xsd"),
            "PurposeForestsKindEType")

        Dim colPurpose = CType(dgvPieces.Columns("PurposeForests"), DataGridViewComboBoxColumn)
        colPurpose.DataSource = purposes
        colPurpose.DisplayMember = "Description"
        colPurpose.ValueMember = "Code"

        ' Особо защитные участки
        Dim protectedForests = LoadEnumFromXsd(
            IO.Path.Combine(SchemasPath, "common3.4\dNameProtectedForestsKindTypes-2.1.5.xsd"),
            "NameProtectedForestsKindEType")

        Dim colProtected = CType(dgvPieces.Columns("NameProtectedForests"), DataGridViewComboBoxColumn)
        colProtected.DataSource = protectedForests
        colProtected.DisplayMember = "Description"
        colProtected.ValueMember = "Code"

        ' Начальное состояние
        UpdateModeControls()
    End Sub

    Private Sub rbFgislk_CheckedChanged(sender As Object, e As EventArgs) _
        Handles rbFgislk.CheckedChanged, rbLu.CheckedChanged
        UpdateModeControls()
    End Sub

    Private Sub UpdateModeControls()
        Dim useFgislk As Boolean = rbFgislk.Checked

        txtFgislkParent.Enabled = useFgislk

        cmbSubforestry.Enabled = Not useFgislk
        txtNameTract.Enabled = Not useFgislk
        txtLuQuarter.Enabled = Not useFgislk
        txtLuTaxation.Enabled = Not useFgislk
    End Sub
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim data As EditTaxationData

        Try
            data = CollectData()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Проверьте ввод",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        Dim builder As New EditTaxationXmlBuilder()
        Dim xml As String = builder.Build(data)

        Using sfd As New SaveFileDialog()
            sfd.Filter = "XML files (*.xml)|*.xml"
            sfd.FileName = "editTaxation.xml"
            If sfd.ShowDialog() = DialogResult.OK Then
                System.IO.File.WriteAllText(sfd.FileName, xml, System.Text.Encoding.UTF8)
                MessageBox.Show("XML сохранён: " & sfd.FileName)
            End If
        End Using
    End Sub

    Private Function CollectData() As EditTaxationData
        Dim data As New EditTaxationData()
        data.UseFgislk = rbFgislk.Checked

        If data.UseFgislk Then
            data.FgislkParentRegistrationNo = txtFgislkParent.Text.Trim()
            If data.FgislkParentRegistrationNo = "" Then
                Throw New Exception("Укажите учётный номер ФГИС ЛК.")
            End If
        Else
            If cmbSubforestry.SelectedItem Is Nothing Then
                Throw New Exception("Выберите участковое лесничество.")
            End If
            data.SubforestryCode = CType(cmbSubforestry.SelectedItem, ComboItem).Code
            data.NameTract = txtNameTract.Text.Trim()
            data.LuForestQuarterNumber = txtLuQuarter.Text.Trim()
            data.LuForestTaxationNumber = txtLuTaxation.Text.Trim()

            If data.LuForestQuarterNumber = "" OrElse data.LuForestTaxationNumber = "" Then
                Throw New Exception("Укажите номер квартала и номер выдела.")
            End If
        End If

        For Each row As DataGridViewRow In dgvPieces.Rows
            If row.IsNewRow Then Continue For

            Dim p As New NewTaxationPiece()
            p.LuForestTaxationNumber = GetCellString(row, "LuForestTaxationNumber")
            p.ObjectGuid = GetCellString(row, "ObjectGuid")

            If p.LuForestTaxationNumber = "" Then
                Throw New Exception("В новом выделе не указан лесоустроительный номер.")
            End If

            ' objectGuid обязателен по схеме
            If p.ObjectGuid = "" Then
                p.ObjectGuid = Guid.NewGuid().ToString()
            End If

            Dim areaText = GetCellString(row, "Area")
            If areaText <> "" Then
                Dim a As Decimal
                If Decimal.TryParse(areaText, Globalization.NumberStyles.Any,
                                    Globalization.CultureInfo.InvariantCulture, a) Then
                    p.Area = a
                Else
                    Throw New Exception($"Не удалось разобрать площадь: {areaText}")
                End If
            End If

            Dim evText = GetCellString(row, "Events")
            If evText <> "" Then
                Dim b As Boolean
                If Boolean.TryParse(evText, b) Then
                    p.Events = b
                Else
                    Throw New Exception($"events должно быть true/false: {evText}")
                End If
            End If

            p.PurposeForests = GetCellString(row, "PurposeForests")
            p.NameProtectedForests = GetCellString(row, "NameProtectedForests")

            Dim remText = GetCellString(row, "RemoveProtectedForests")
            If remText <> "" Then
                Dim b As Boolean
                If Boolean.TryParse(remText, b) Then
                    p.RemoveProtectedForests = b
                Else
                    Throw New Exception($"removeProtectedForests должно быть true/false: {remText}")
                End If
            End If

            If Not String.IsNullOrEmpty(p.NameProtectedForests) AndAlso
               p.RemoveProtectedForests.HasValue Then
                Throw New Exception(
                    "Нельзя одновременно задавать nameProtectedForests и removeProtectedForests.")
            End If

            data.Pieces.Add(p)
        Next

        If data.Pieces.Count = 0 Then
            Throw New Exception("Добавьте хотя бы один новый выдел.")
        End If

        Return data
    End Function

    Private Function GetCellString(row As DataGridViewRow, colName As String) As String
        Dim v = row.Cells(colName).Value
        If v Is Nothing Then Return ""
        Return v.ToString().Trim()
    End Function


    Private Sub btnAddPiece_Click(sender As Object, e As EventArgs) Handles btnAddPiece.Click
        dgvPieces.Rows.Add()
    End Sub

    Private Sub btnRemovePiece_Click(sender As Object, e As EventArgs) Handles btnRemovePiece.Click
        If dgvPieces.SelectedRows.Count = 0 Then Return
        For Each r As DataGridViewRow In dgvPieces.SelectedRows
            If Not r.IsNewRow Then dgvPieces.Rows.Remove(r)
        Next
    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class