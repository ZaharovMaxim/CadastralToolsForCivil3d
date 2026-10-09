Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Public Class ProjProfilePropXYZD
    Inherits System.Windows.Forms.Form
    Public boolProf As Boolean = False
    Public AlignLine As Autodesk.Civil.DatabaseServices.Alignment = Nothing
    Public ProfLine As Autodesk.Civil.DatabaseServices.Profile = Nothing
    Public ProfStart As Double = 0
    Public ProfEnd As Double = 0
    Public Chag As Integer = 0
    'начальные установки при инициализации формы
    Public Sub New()
        InitializeComponent()
        'загружаем все профиля в форму ComboBox
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
            Dim ArrayAlign As String() = {""}
            Dim CountAlign As Integer = 0
            'ищем все трассы чертежа, первую делаем по умолчанию
            For Each id As ObjectId In DocColectionAlignment
                Dim objAlign As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                Dim Align As Alignment = objAlign
                Dim nameAlign As String = Align.Name
                ReDim Preserve ArrayAlign(CountAlign)
                ArrayAlign(CountAlign) = nameAlign
                If CountAlign = 0 Then
                    AlignLine = objAlign
                    ProfStart = Align.StartingStation
                    ProfEnd = Align.EndingStation
                    'форматируем пикет
                    Try
                        TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
                    Catch ex As System.ArgumentException
                        TextBox1.Text = ProfStart
                    End Try
                    Try
                        TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
                    Catch ex As System.ArgumentException
                        TextBox2.Text = ProfEnd
                    End Try
                End If
                CountAlign += 1
            Next
            'ищем профиля
            Dim nameAlign1 As String = ArrayAlign(0)
            Dim ArrayProf As String() = {""}
            Dim CountProf As Integer = 0
            If nameAlign1.Trim.Length > 0 Then
                Dim DocColectionAlignment1 As ObjectIdCollection = CivilDoc.GetAlignmentIds
                For Each id As ObjectId In DocColectionAlignment1
                    Dim objAlign1 As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                    Dim Align1 As Alignment = objAlign1
                    Dim nameAlign As String = Align1.Name
                    If nameAlign Like nameAlign1 Then
                        Dim ProfIDColl As ObjectIdCollection = Align1.GetProfileIds
                        For Each id1 As ObjectId In ProfIDColl
                            Dim objProf As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(id1, OpenMode.ForRead)
                            Dim NameProf As String = objProf.Name
                            ReDim Preserve ArrayProf(CountProf)
                            ArrayProf(CountProf) = NameProf
                            If CountProf = 0 Then
                                ProfLine = objProf
                            End If
                            CountProf += 1
                        Next
                    End If
                Next
            End If
            ComboBox1.DataSource = ArrayAlign
            ComboBox2.DataSource = ArrayProf
            tr.Commit()
        End Using
        'находим все директории с файлами Excel
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTables As String = "d:\AutodeskCivilTools\TemplateTables\Roads\Ведомости рабочих отметок продольного профиля\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "TemplateTables\Roads\Ведомости рабочих отметок продольного профиля\") Then
            PutchTables = parentPatch2 & "TemplateTables\Roads\Ведомости рабочих отметок продольного профиля\"
        End If
        ComboBox3.Tag = PutchTables
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор траассы из списка
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim ArrayProf As String() = {""}
        Dim CountProf As Integer = 0
        ComboBox2.DataSource = ArrayProf
        'загружаем все профиля в форму ComboBox
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
            Dim ArrayAlign As String() = {""}
            Dim CountAlign As Integer = 0
            Dim Align As Alignment = Nothing
            For Each id As ObjectId In DocColectionAlignment
                Dim objAlign As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                Align = objAlign
                Dim nameAlign As String = Align.Name
                If nameAlign Like ComboBox1.Text Then
                    AlignLine = objAlign
                    Exit For
                End If
            Next
            ProfStart = Align.StartingStation
            ProfEnd = Align.EndingStation
            'форматируем пикет
            'форматируем пикет
            Try
                TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
            Catch ex As System.ArgumentException
                TextBox1.Text = ProfStart
            End Try
            Try
                TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
            Catch ex As System.ArgumentException
                TextBox2.Text = ProfEnd
            End Try
            'ищем профиля
            Dim ProfIDColl As ObjectIdCollection = Align.GetProfileIds
            For Each id1 As ObjectId In ProfIDColl
                Dim objProf As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(id1, OpenMode.ForRead)
                Dim NameProf As String = objProf.Name
                ReDim Preserve ArrayProf(CountProf)
                ArrayProf(CountProf) = NameProf
                If CountProf = 0 Then
                    ProfLine = objProf
                End If
                CountProf += 1
            Next
            ComboBox2.DataSource = ArrayProf
            tr.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор трассы
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите трассу: ")
            resBlk = m_Editor.GetEntity(opt)
            If Not (resBlk.Status = PromptStatus.None Or resBlk.Status = PromptStatus.Cancel Or resBlk.Status = PromptStatus.Error) Then 'нет выбора
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim obj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf obj Is Alignment Then
                    Dim Align As Alignment = obj
                    AlignLine = obj
                    ComboBox1.Text = Align.Name
                    Dim ArrayProf As String() = {""}
                    Dim CountProf As Integer = 0
                    ComboBox2.DataSource = ArrayProf
                    Dim ProfIDColl As ObjectIdCollection = Align.GetProfileIds
                    For Each id1 As ObjectId In ProfIDColl
                        Dim objProf As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(id1, OpenMode.ForRead)
                        Dim NameProf As String = objProf.Name
                        ReDim Preserve ArrayProf(CountProf)
                        ArrayProf(CountProf) = NameProf
                        If CountProf = 0 Then
                            ProfLine = objProf
                        End If
                        CountProf += 1
                    Next
                    ProfStart = Align.StartingStation
                    ProfEnd = Align.EndingStation
                    'форматируем пикет
                    Try
                        TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
                    Catch ex As System.ArgumentException
                        TextBox1.Text = ProfStart
                    End Try
                    Try
                        TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
                    Catch ex As System.ArgumentException
                        TextBox2.Text = ProfEnd
                    End Try
                    ComboBox2.DataSource = ArrayProf
                Else
                    MsgBox("Выбранный объект должен иметь тип ТРАССА!!!")
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор линии профиля
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите линию профиля: ")
            resBlk = m_Editor.GetEntity(opt)
            If Not (resBlk.Status = PromptStatus.None Or resBlk.Status = PromptStatus.Cancel Or resBlk.Status = PromptStatus.Error) Then 'нет выбора
                Dim pvId As ObjectId = resBlk.ObjectId
                Dim objProf As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)
                If TypeOf (objProf) Is Autodesk.Civil.DatabaseServices.Profile Then
                    Dim Prof As Autodesk.Civil.DatabaseServices.Profile = objProf
                    'смотрим какой трассе он принадлежит
                    Dim IDAlign As ObjectId = Prof.AlignmentId
                    Dim objAlign As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(IDAlign, OpenMode.ForRead)
                    Dim Align As Alignment = objAlign
                    AlignLine = objAlign
                    ComboBox1.Text = Align.Name
                    ProfStart = Align.StartingStation
                    ProfEnd = Align.EndingStation
                    'форматируем пикет
                    Try
                        TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
                    Catch ex As System.ArgumentException
                        TextBox1.Text = ProfStart
                    End Try
                    Try
                        TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
                    Catch ex As System.ArgumentException
                        TextBox2.Text = ProfEnd
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'заново заполняем массив с профилями
                    Dim ArrayProf As String() = {""}
                    Dim CountProf As Integer = 0
                    ComboBox2.DataSource = ArrayProf
                    Dim ProfIDColl As ObjectIdCollection = Align.GetProfileIds
                    For Each id1 As ObjectId In ProfIDColl
                        Dim objProf1 As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(id1, OpenMode.ForRead)
                        Dim NameProf As String = objProf.Name
                        ReDim Preserve ArrayProf(CountProf)
                        ArrayProf(CountProf) = NameProf
                        If CountProf = 0 Then
                            ProfLine = objProf
                        End If
                        CountProf += 1
                    Next
                    ComboBox2.DataSource = ArrayProf
                    ComboBox2.Text = Prof.Name
                    tr.Commit()
                Else
                    MsgBox("Выбранный объект долженн иметь тип ЛИНИЯ ПРОФИЛЯ!!!")
                End If
            End If


        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolProf = True
        Chag = NumericUpDown1.Value
        If ComboBox2.Text = "" Then ProfLine = Nothing
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolProf = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'указать на трассе
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox1.Text
        'получаем доступ к трассе
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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

            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку начала трассы: "
            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
            Dim PKVedStartAlign As Double = 0
            Dim OfsetVedStartAlign As Double = 0
            If pPtRes.Status = PromptStatus.None Then
                TextBox1.Text = Align.StartingStation
            Else
                Try
                    Align.StationOffset(ptStart.X, ptStart.Y, PKVedStartAlign, OfsetVedStartAlign)
                    ProfStart = PKVedStartAlign
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    Exit Sub
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
                TextBox2.Text = Align.EndingStation
            Else
                Try
                    Align.StationOffset(ptEnd.X, ptEnd.Y, PKVedEndAlign, OfsetVedStartAlign)
                    ProfEnd = PKVedEndAlign
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    Exit Sub
                End Try
            End If
            'форматируем пикет
            Try
                TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
            Catch ex As System.ArgumentException
                TextBox1.Text = ProfStart
            End Try
            Try
                TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
            Catch ex As System.ArgumentException
                TextBox2.Text = ProfEnd
            End Try
        End Using
    End Sub
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        'загружаем все профиля в форму ComboBox
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
            Dim NameAlignText As String = ComboBox1.Text
            Dim Align As Alignment = Nothing
            For Each id As ObjectId In DocColectionAlignment
                Dim objAlign As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                Align = objAlign
                Dim nameAlign As String = Align.Name
                If nameAlign Like ComboBox1.Text Then
                    AlignLine = objAlign
                    Exit For
                End If
            Next
            'ищем профиля
            Dim NameProfText As String = ComboBox2.Text
            Dim ProfIDColl As ObjectIdCollection = Align.GetProfileIds
            For Each id1 As ObjectId In ProfIDColl
                Dim objProf As Autodesk.Civil.DatabaseServices.Profile = tr.GetObject(id1, OpenMode.ForRead)
                Dim NameProf As String = objProf.Name
                If NameProf Like NameProfText Then
                    ProfLine = objProf
                    Exit For
                End If
            Next
            tr.Commit()
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'указать на профиле
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox1.Text
        Dim NameProf As String = ComboBox2.Text
        'получаем доступ к трассе
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim Align As Alignment = Nothing
        Dim Prof As SchemeOfRouteInPlan = Nothing
        Using tr
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
                        TextBox1.Text = Align.StartingStation
                    Else
                        Try
                            ProfView.FindStationAndElevationAtXY(ptStart.X, ptStart.Y, PKStartAlign, ElevPK)
                            ProfStart = PKStartAlign
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            MsgBox("Не верно указана точка начала!!!")
                            Exit Sub
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
                        TextBox2.Text = Align.EndingStation
                    Else
                        Try
                            ProfView.FindStationAndElevationAtXY(ptEnd.X, ptEnd.Y, PKEndAlign, ElevPK)
                            ProfEnd = PKEndAlign
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            MsgBox("Не верно указана точка начала!!!")
                            Exit Sub
                        End Try
                    End If
                    'форматируем пикет
                    Try
                        TextBox1.Text = Align.GetStationStringWithEquations(ProfStart)
                    Catch ex As System.ArgumentException
                        TextBox1.Text = ProfStart
                    End Try
                    Try
                        TextBox2.Text = Align.GetStationStringWithEquations(ProfEnd)
                    Catch ex As System.ArgumentException
                        TextBox2.Text = ProfEnd
                    End Try
                    tr.Commit()
                Else
                    MsgBox("Выбранный объект должен иметь тип ВИД ПРОФИЛЯ!!!")
                End If
            End If

        End Using
    End Sub

    Private Sub ProjProfilePropXYZD_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub
End Class