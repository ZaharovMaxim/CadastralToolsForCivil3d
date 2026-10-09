Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Public Class FormCreateRoadLayers
    Public boolOk As Boolean = False
    Public ArrayLayers As String() = Nothing
    Public ArrayPSTable As String() = {""}
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ' Добавить код инициализации после вызова InitializeComponent().
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument
        Dim rng As Random = New Random()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
            'забираем в массив все поверхности
            Dim idSurfCollection As ObjectIdCollection = CivilDoc.GetSurfaceIds()
            If idSurfCollection.Count > 0 Then
                Dim ArraySurf As String() = Nothing
                Dim countArraySurf As Integer = 0
                For Each id As ObjectId In idSurfCollection
                    Dim surf As TinSurface = acTrans.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve ArraySurf(countArraySurf)
                    ArraySurf(countArraySurf) = surf.Name
                    countArraySurf += 1
                Next
                If IsArray(ArraySurf) = True Then
                    ComboBox1.DataSource = ArraySurf
                End If
            End If
            'забираем в массив все слои
            Dim nameLayerSolid3d As String = "0"
            Dim boolFindArray As Integer = AcadFunction.FuncReadLayersDWG(ArrayLayers, nameLayerSolid3d)
            Dim cmbbox As DataGridViewComboBoxCell = DataGridView1.Rows(0).Cells(1)
            If IsArray(ArrayLayers) = True Then
                cmbbox.DataSource = ArrayLayers
                cmbbox.Value = nameLayerSolid3d
            End If
            DataGridView1.Rows(0).Cells(2).Value = rng.Next(1, 10)
            DataGridView1.Rows(0).Cells(3).Value = "0"
            DataGridView1.Rows(0).Cells(4).Value = "0"
            'забираем в массив все таблицы PS
            Dim countPS As Integer = FuncPropertySet.FuncReturnArrayNamesDictGWG(ArrayPSTable)
            If IsArray(ArrayPSTable) = True Then
                Dim cmbbox2 As DataGridViewComboBoxCell = DataGridView1.Rows(0).Cells(5)
                cmbbox2.DataSource = ArrayPSTable
            End If
        End Using
    End Sub
    'OR
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolOk = True
        Me.Hide()
    End Sub
    'ОТМЕНА
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolOk = False
        Me.Hide()
    End Sub
    'автоматическое заполнение
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If DataGridView1.Rows.Count > 1 Then
            Dim indRow As Integer = e.RowIndex
            Dim cmbboxLayer As DataGridViewComboBoxCell = DataGridView1.Rows(indRow).Cells(1)
            Dim cmbboxPS As DataGridViewComboBoxCell = DataGridView1.Rows(indRow).Cells(5)
            Dim rng As Random = New Random()
            If IsNothing(cmbboxLayer.Value) Then
                If indRow > 0 Then
                    'слои
                    Dim cmbboxPrev As DataGridViewComboBoxCell = DataGridView1.Rows(indRow - 1).Cells(1)
                    cmbboxLayer.DataSource = cmbboxPrev.DataSource
                    cmbboxLayer.Value = cmbboxPrev.Value
                    'таблицы
                    Dim cmbboxPrev1 As DataGridViewComboBoxCell = DataGridView1.Rows(indRow - 1).Cells(5)
                    cmbboxPS.DataSource = cmbboxPrev1.DataSource
                    cmbboxPS.Value = cmbboxPrev1.Value
                End If
                DataGridView1.Rows(indRow).Cells(2).Value = rng.Next(1, 10)
                DataGridView1.Rows(indRow).Cells(3).Value = "0"
                DataGridView1.Rows(indRow).Cells(4).Value = "0"
            End If
        End If
    End Sub
    'контекстное меню Удалить строку
    Private Sub УдалитьСтрокуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьСтрокуToolStripMenuItem.Click
        If DataGridView1.Rows.Count > 1 Then
            Dim col As DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
            Dim colselect As DataGridViewCell = col.Item(0)
            Dim rowInd As Integer = colselect.RowIndex
            DataGridView1.Rows.RemoveAt(rowInd)
        End If
    End Sub
    'контекстное меню Вставить строку
    Private Sub ВставитьСтрокуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВставитьСтрокуToolStripMenuItem.Click
        If DataGridView1.Rows.Count > 2 Then
            Dim col As DataGridViewSelectedCellCollection = DataGridView1.SelectedCells
            Dim colselect As DataGridViewCell = col.Item(0)
            Dim rowInd As Integer = colselect.RowIndex
            DataGridView1.Rows.Insert(rowInd)
        End If
    End Sub

End Class