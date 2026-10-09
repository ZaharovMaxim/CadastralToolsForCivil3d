Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Public Class MarksOnTopCoating
    Inherits System.Windows.Forms.Form
    Public boolProf As Boolean = False
    Public indLeft As Integer = 1
    Public indRight As Integer = 1
    Private mouseLocationRow As Integer = 0
    Private mouseLocationCol As Integer = 0
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
    End Sub

    'НАЖАТА КНОПКА ОК
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolProf = True
        'Me.Hide()
    End Sub
    'НАЖАТА КНОПКА ОТМЕНА
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolProf = False
        Me.Hide()
    End Sub
    'выбрать пикет начала и конца трассы на трассе
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox6.Text
        'получаем доступ к трассе
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim Align As Alignment = Nothing
        Using tr
            Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
            Dim ArrayAlign As String() = {""}
            Dim CountAlign As Integer = 0
            For Each id As ObjectId In DocColectionAlignment
                Dim objAlign As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                Align = objAlign
                Dim nameAlign1 As String = Align.Name
                If NameAlign Like nameAlign1 Then
                    Exit For
                End If
            Next
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'пикет начала и конца трассы
            Dim StartPk As Double = Align.StartingStation
            Dim EndPk As Double = Align.EndingStation
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку начала трассы: "
            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
            Dim PKVedStartAlign As Double = 0
            Dim OfsetVedStartAlign As Double = 0
            If pPtRes.Status = PromptStatus.None Then
                StartPk = Align.StartingStation
            ElseIf pPtRes.Status = PromptStatus.Cancel Then
                StartPk = Align.StartingStation
            Else
                Try
                    Align.StationOffset(ptStart.X, ptStart.Y, PKVedStartAlign, OfsetVedStartAlign)
                    StartPk = PKVedStartAlign
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                End Try
            End If
            'запрашиваем точку конца трассы
            pPtOpts = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку конца трассы: "
            pPtRes = m_Editor.GetPoint(pPtOpts)
            Dim ptEnd As Point3d = pPtRes.Value 'координаты точки конца трассы
            Dim PKVedEndAlign As Double = 0
            If pPtRes.Status = PromptStatus.None Then
                EndPk = Align.EndingStation
            ElseIf pPtRes.Status = PromptStatus.Cancel Then
                EndPk = Align.EndingStation
            Else
                Try
                    Align.StationOffset(ptEnd.X, ptEnd.Y, PKVedEndAlign, OfsetVedStartAlign)
                    EndPk = PKVedEndAlign
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                End Try
            End If
            'форматируем пикет
            Try
                TextBox1.Text = Align.GetStationStringWithEquations(StartPk)
                TextBox1.Tag = StartPk
            Catch ex As System.ArgumentException
                TextBox1.Text = StartPk
                TextBox1.Tag = StartPk
            End Try
            Try
                TextBox2.Text = Align.GetStationStringWithEquations(EndPk)
                TextBox2.Tag = EndPk
            Catch ex As System.ArgumentException
                TextBox2.Text = EndPk
                TextBox2.Tag = EndPk
            End Try
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'указать пикет на профиле
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox6.Text
        Dim NameProf As String = ComboBox7.Text
        'получаем доступ к трассе
        Dim StartPK As Double = 0
        Dim EndPK As Double = 0
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim Align As Alignment = Nothing
        Dim Prof As SchemeOfRouteInPlan = Nothing
        Using tr
            Dim AlignCol As ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCol.Count > 0 Then
                For Each id As ObjectId In AlignCol
                    Align = tr.GetObject(id, OpenMode.ForRead)
                    If Align.Name Like NameAlign Then
                        Exit For
                    Else
                        Align = Nothing
                    End If
                Next
            End If
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите вид профиля: ")
            resBlk = m_Editor.GetEntity(opt)
            If Not (resBlk.Status = PromptStatus.None Or resBlk.Status = PromptStatus.Cancel Or resBlk.Status = PromptStatus.Error) Then 'нет выбора
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim objProf As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf objProf Is ProfileView Then
                    Dim ProfView As ProfileView = objProf
                    '=================================================================================
                    'делаем запрос на начальную и конечные точки
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.AllowNone = True
                    pPtOpts.Message = vbLf & "Укажите точку начала трассы на виде профиля: "
                    Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
                    Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
                    Dim PKStartAlign As Double = 0 'начальный пикет
                    Dim ElevPK As Double = 0
                    If pPtRes.Status = PromptStatus.None Then
                        StartPK = ProfView.StationStart
                    ElseIf pPtRes.Status = PromptStatus.Cancel Then
                        StartPK = ProfView.StationStart
                    Else
                        Try
                            ProfView.FindStationAndElevationAtXY(ptStart.X, ptStart.Y, PKStartAlign, ElevPK)
                            StartPK = PKStartAlign
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            MsgBox("Не верно указана точка начала!!!")
                        End Try
                    End If
                    'запрашиваем точку конца трассы
                    pPtOpts = New PromptPointOptions("")
                    pPtOpts.AllowNone = True
                    pPtOpts.Message = vbLf & "Укажите точку конца трассы на виде профиля: "
                    pPtRes = m_Editor.GetPoint(pPtOpts)
                    Dim ptEnd As Point3d = pPtRes.Value 'координаты точки конца трассы
                    Dim PKEndAlign As Double = 0
                    If pPtRes.Status = PromptStatus.None Then
                        EndPK = Align.EndingStation
                    ElseIf pPtRes.Status = PromptStatus.Cancel Then
                        EndPK = Align.EndingStation
                    Else
                        Try
                            ProfView.FindStationAndElevationAtXY(ptEnd.X, ptEnd.Y, PKEndAlign, ElevPK)
                            EndPK = PKEndAlign
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            MsgBox("Не верно указана точка начала!!!")
                            Exit Sub
                        End Try
                    End If
                    'форматируем пикет
                    If IsNothing(Align) = False Then
                        Try
                            TextBox1.Text = Align.GetStationStringWithEquations(StartPK)
                            TextBox1.Tag = StartPK
                        Catch ex As System.ArgumentException
                            TextBox1.Text = StartPK
                            TextBox1.Tag = StartPK
                        End Try
                        Try
                            TextBox2.Text = Align.GetStationStringWithEquations(EndPK)
                            TextBox2.Tag = EndPK
                        Catch ex As System.ArgumentException
                            TextBox2.Text = EndPK
                            TextBox2.Tag = EndPK
                        End Try
                    Else
                        TextBox1.Text = StartPK
                        TextBox1.Tag = StartPK
                        TextBox2.Text = EndPK
                        TextBox2.Tag = EndPK
                    End If
                Else
                    MsgBox("Выбранный объект должен иметь тип ВИД ПРОФИЛЯ!!!")
                End If
            End If

        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор ячейки (кода)
    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Try
            If DataGridView1.RowCount > 2 Then
                Dim ColInd As Integer = e.ColumnIndex
                Dim rowInd As Integer = e.RowIndex
                '=============================================================================================
                'код для сторны лево
                If ColInd = 0 Then
                    Dim boolChLeft As Integer = Val(DataGridView1.Rows(rowInd).Cells(ColInd).Value) 'смотрим нажата ли область
                    If boolChLeft = 1 Then
                        'смотрим обозначение
                        Dim NembT As String = DataGridView1.Rows(rowInd).Cells(3).Value 'смотрим нажата ли область
                        If NembT Like "" Then
                            DataGridView1.Rows(rowInd).Cells(3).Value = "Т" & indLeft 'смотрим нажата ли область
                            indLeft += 1
                            Me.Hide()
                            Me.Show()
                        End If
                    Else
                        DataGridView1.Rows(rowInd).Cells(3).Value = "" 'смотрим нажата ли область
                        Me.Hide()
                        Me.Show()
                    End If
                End If
                '=================================================================================================
                'код для стороны право
                If ColInd = 1 Then
                    Dim boolChRight As Integer = Val(DataGridView1.Rows(rowInd).Cells(ColInd).Value) 'смотрим нажата ли область
                    If boolChRight = 1 Then
                        'смотрим обозначение
                        Dim NembT As String = DataGridView1.Rows(rowInd).Cells(4).Value 'смотрим нажата ли область
                        If NembT Like "" Then
                            DataGridView1.Rows(rowInd).Cells(4).Value = "Т" & indRight 'смотрим нажата ли область
                            indRight += 1
                            Me.Hide()
                            Me.Show()
                        End If
                    Else
                        'смотрим обозначение
                        DataGridView1.Rows(rowInd).Cells(4).Value = "" 'смотрим нажата ли область
                        Me.Hide()
                        Me.Show()
                    End If
                End If
            End If
        Catch ex As ArgumentOutOfRangeException
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'реактор на клик мыши по ячейке
    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        mouseLocationRow = e.RowIndex
        mouseLocationCol = e.ColumnIndex
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т1
    Private Sub T1ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles Т1ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T1"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т2
    Private Sub T2ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T2ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T2"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т3
    Private Sub T3ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T3ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T3"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т4
    Private Sub T4ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T4ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T4"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т5
    Private Sub T5ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T5ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T5"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т6
    Private Sub T6ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T6ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T6"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т7
    Private Sub T7ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T7ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T7"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т8
    Private Sub T8ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T8ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T8"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т9
    Private Sub T9ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T9ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T9"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ставим код Т10
    Private Sub T10ToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles T10ToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = "T10"
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 1
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'стереть код
    Private Sub НетToolStripMenuItem_MouseEnter(sender As Object, e As EventArgs) Handles НетToolStripMenuItem.MouseEnter
        If mouseLocationCol > 2 And mouseLocationRow >= 0 Then
            DataGridView1.Rows(mouseLocationRow).Cells(mouseLocationCol).Value = ""
            If mouseLocationCol = 3 Then
                DataGridView1.Rows(mouseLocationRow).Cells(0).Value = 0
            ElseIf mouseLocationCol = 4 Then
                DataGridView1.Rows(mouseLocationRow).Cells(1).Value = 0
            End If
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox2.Enabled = True
        Else
            ComboBox2.Enabled = False
        End If
    End Sub


    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            ComboBox3.Enabled = True
        Else
            ComboBox3.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class