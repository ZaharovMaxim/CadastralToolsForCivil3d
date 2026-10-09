Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports System.IO

Public Class RoadPlane
    Public boolStatus As Boolean = False 'нажата кнопка OK
    Public parentPatch2 As String = "d:\AutodeskCivilTools\"
    Public PatchFileTemplate As String = "" 'абсолютный путь к файлу-шаблону
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        parentPatch2 = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            ComboBox8.Tag = PutchTables
            Try
                For Each fileDWG As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileDWG)
                        If CountArrayTemplates = 0 Then
                            ComboBox8.Tag = PutchTables
                            PatchFileTemplate = fileDWG
                        End If
                        CountArrayTemplates += 1
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox8.DataSource = ArrayTemplates
        End If
        ' Добавить код инициализации после вызова InitializeComponent().
        'формат угловых величин
        Dim StrGradFormat As String() = ({"Десятичные градусы", "Град.мин.сек", "Грады", "Радианы"})
        ComboBox1.DataSource = StrGradFormat
        ComboBox1.SelectedIndex = 1
        'формат длин линий
        Dim StrPKFormat As String() = ({"Пикет+", "Десятичное число", "Километр+"})
        ComboBox2.DataSource = StrPKFormat
        'форма таблицы
        Dim NameTables As String() = {"Трасса автомобильной дороги"}
        Dim NameStylesTable As String() = Nothing
        Dim boolReadStyleTbl As Boolean = AcadFunction.FuncReadNameObjectAllDestDWG(PatchFileTemplate, NameTables, NameStylesTable, "Таблица")
        If IsNothing(NameStylesTable) = False Then
            If IsArray(NameStylesTable) = True Then
                ComboBox4.DataSource = NameStylesTable
            End If
        End If
        'формат
        Dim StrFormat As String() = ({"Таблица Excel", "Таблица AutoCAD"})
        ComboBox10.DataSource = StrFormat
        'формат
        Dim StrVidDoc As String() = ({"Проектная документация", "Рабочая документация"})
        ComboBox3.DataSource = StrVidDoc
        'читаем все трассы с чертежа
        Dim ArrayNameAlign As String() = Nothing
        Dim CountArrayNameAlign As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
        If AlignCollection.Count > 0 Then
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                For Each idAlign As ObjectId In AlignCollection
                    Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                    ReDim Preserve ArrayNameAlign(CountArrayNameAlign)
                    ArrayNameAlign(CountArrayNameAlign) = Align.Name
                    If CountArrayNameAlign = 0 Then
                        Try
                            Dim startPK As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4)
                            Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                            TextBox2.Text = startPKstr
                            TextBox2.Tag = startPK
                        Catch ex As System.ArgumentException
                            TextBox2.Text = Math.Round(Align.StartingStation, 4)
                            TextBox2.Tag = Align.StartingStation
                        End Try
                        Try
                            Dim EndPK As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4)
                            Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                            TextBox3.Text = EndPKStr
                            TextBox3.Tag = EndPK
                        Catch ex As System.ArgumentException
                            TextBox3.Text = Math.Round(Align.EndingStation, 4)
                            TextBox3.Tag = Align.EndingStation
                        End Try
                    End If
                    CountArrayNameAlign += 1
                Next
            End Using
        End If
        If IsArray(ArrayNameAlign) = True Then
            ComboBox5.DataSource = ArrayNameAlign
        End If
    End Sub
    '==========================================================================================
    'нажата кнопка ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolStatus = True
        PatchFileTemplate = ComboBox8.Tag & ComboBox8.Text
        Me.Hide()
    End Sub
    '========================================================================================
    'указать точку начала трассы
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox5.Text
        If NameAlign.Trim.Length > 0 Then
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCollection.Count > 0 Then
                Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                Using tr
                    For Each idAlign As ObjectId In AlignCollection
                        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                        If Align.Name Like NameAlign Then
                            Dim PKStartAlign As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4) 'пикет начала трассы
                            Dim OfsetStartAlign As Double = 0 'смещение точки от трассы
                            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                            pPtOpts.AllowNone = True
                            pPtOpts.Message = vbLf & "Укажите точку начала трассы"
                            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
                            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
                            If pPtRes.Status = PromptStatus.None Or pPtRes.Status = PromptStatus.Cancel Then 'нет выбора
                                Try
                                    Dim startPK As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4)
                                    Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                                    TextBox2.Text = startPKstr
                                    TextBox2.Tag = startPK
                                Catch ex As System.ArgumentException
                                    TextBox2.Text = Math.Round(Align.StartingStation, 4)
                                    TextBox2.Tag = Align.StartingStation
                                End Try
                            Else
                                Try
                                    Align.StationOffset(ptStart.X, ptStart.Y, PKStartAlign, OfsetStartAlign)
                                    Try
                                        Dim startPK As Double = MathFunction.FuncTrimDigitCall(PKStartAlign, 4)
                                        Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                                        TextBox2.Text = startPKstr
                                        TextBox2.Tag = startPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox2.Text = Math.Round(Align.StartingStation, 4)
                                        TextBox2.Tag = Align.StartingStation
                                    End Try
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    MsgBox("Не верно указана точка начала!!!")
                                    Try
                                        Dim startPK As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4)
                                        Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                                        TextBox2.Text = startPKstr
                                        TextBox2.Tag = startPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox2.Text = Math.Round(Align.StartingStation, 4)
                                        TextBox2.Tag = Align.StartingStation
                                    End Try
                                    Exit Sub
                                Catch ex As ArgumentException
                                    MsgBox("Не верно указана точка начала!!!")
                                    Try
                                        Dim startPK As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4)
                                        Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                                        TextBox2.Text = startPKstr
                                        TextBox2.Tag = startPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox2.Text = Math.Round(Align.StartingStation, 4)
                                        TextBox2.Tag = Align.StartingStation
                                    End Try
                                    Exit Sub
                                End Try
                            End If
                        End If
                    Next
                End Using
            End If
        End If
    End Sub
    '========================================================================================
    'указать точку конца трассы
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox5.Text
        If NameAlign.Trim.Length > 0 Then
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCollection.Count > 0 Then
                For Each idAlign As ObjectId In AlignCollection
                    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                    Using tr
                        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                        If Align.Name Like NameAlign Then
                            Dim PKEndAlign As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4) 'пикет конца трассы
                            Dim OfsetEndAlign As Double = 0 'смещение точки от трассы
                            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                            pPtOpts.AllowNone = True
                            pPtOpts.Message = vbLf & "Укажите точку конца трассы"
                            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
                            Dim ptEnd As Point3d = pPtRes.Value 'координаты точки начала трассы
                            If pPtRes.Status = PromptStatus.None Or pPtRes.Status = PromptStatus.Cancel Then 'нет выбора
                                Try
                                    Dim EndPK As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4)
                                    Dim endPKstr As String = Align.GetStationStringWithEquations(EndPK)
                                    TextBox3.Text = endPKstr
                                    TextBox3.Tag = EndPK
                                Catch ex As System.ArgumentException
                                    TextBox3.Text = Math.Round(Align.EndingStation, 4)
                                    TextBox3.Tag = Align.EndingStation
                                End Try
                            Else
                                Try
                                    Align.StationOffset(ptEnd.X, ptEnd.Y, PKEndAlign, OfsetEndAlign)
                                    Try
                                        Dim EndPK As Double = MathFunction.FuncTrimDigitCall(PKEndAlign, 4)
                                        Dim EndPKstr As String = Align.GetStationStringWithEquations(EndPK)
                                        TextBox3.Text = EndPKstr
                                        TextBox3.Tag = EndPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox3.Text = Math.Round(Align.EndingStation, 4)
                                        TextBox3.Tag = Align.EndingStation
                                    End Try
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    MsgBox("Не верно указана точка начала!!!")
                                    Try
                                        Dim EndPK As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4)
                                        Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                                        TextBox3.Text = EndPKStr
                                        TextBox3.Tag = EndPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox3.Text = Math.Round(Align.EndingStation, 4)
                                        TextBox3.Tag = Align.EndingStation
                                    End Try
                                    Exit Sub
                                Catch ex As ArgumentException
                                    MsgBox("Не верно указана точка начала!!!")
                                    Try
                                        Dim EndPK As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4)
                                        Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                                        TextBox3.Text = EndPKStr
                                        TextBox3.Tag = EndPK
                                    Catch ex1 As System.ArgumentException
                                        TextBox3.Text = Math.Round(Align.EndingStation, 4)
                                        TextBox3.Tag = Align.EndingStation
                                    End Try
                                    Exit Sub
                                End Try
                            End If
                        End If
                    End Using

                Next
            End If
        End If
    End Sub
    '========================================================================================
    'пользователь выбрал трассу из списка
    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim NameAlign As String = ComboBox5.Text
        If NameAlign.Trim.Length > 0 Then
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCollection.Count > 0 Then
                Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                Using tr
                    For Each idAlign As ObjectId In AlignCollection
                        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                        If Align.Name Like NameAlign Then
                            Try
                                Dim startPK As Double = MathFunction.FuncTrimDigitCall(Align.StartingStation, 4)
                                Dim startPKstr As String = Align.GetStationStringWithEquations(startPK)
                                TextBox2.Text = startPKstr
                                TextBox2.Tag = startPK
                            Catch ex As System.ArgumentException
                                TextBox2.Text = Math.Round(Align.StartingStation, 4)
                                TextBox2.Tag = Align.StartingStation
                            End Try
                            Try
                                Dim EndPK As Double = MathFunction.FuncTrimDigitFloo(Align.EndingStation, 4)
                                Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                                TextBox3.Text = EndPKStr
                                TextBox3.Tag = EndPK
                            Catch ex As System.ArgumentException
                                TextBox3.Text = Math.Round(Align.EndingStation, 4)
                                TextBox3.Tag = Align.EndingStation
                            End Try
                        End If
                    Next
                End Using
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка отмены
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'опция СХЕМА РАЗБИВКИ КРИВЫХ В ПЛАНЕ
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'формат
        If RadioButton1.Checked = True Then
            Dim StrFormat As String() = ({"Таблица Excel"})
            ComboBox10.DataSource = StrFormat
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'опция ВЕЛОМОСТЬ УГЛОВ ПОВОРОТА ПРЯМЫХ И КРИВЫХ
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'формат
        Dim StrFormat As String() = ({"Таблица Excel", "Таблица AutoCAD"})
        ComboBox10.DataSource = StrFormat
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор (excel или Autocad)
    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        If ComboBox10.Text Like "Таблица Excel" Then
            'читаем шаблоны условных знаков
            Dim PutchTablesExcel As String = parentPatch2 & "\TemplateTables\Roads\Ведомости углов поворота\"
            Dim ArrayTemplatesExcel As String() = Nothing
            Dim CountArrayTemplatesExcel As Integer = 0
            If IO.Directory.Exists(PutchTablesExcel) = True Then
                ComboBox10.Tag = PutchTablesExcel
                Try
                    For Each fileDWG As String In Directory.GetFiles(PutchTablesExcel)
                        'читаем расширение
                        Dim ext As String = Path.GetExtension(fileDWG)
                        If ext Like ".xlsx" Then
                            ReDim Preserve ArrayTemplatesExcel(CountArrayTemplatesExcel)
                            ArrayTemplatesExcel(CountArrayTemplatesExcel) = Path.GetFileName(fileDWG)
                            CountArrayTemplatesExcel += 1
                        End If
                    Next
                Catch ex As IOException
                Catch ex As UnauthorizedAccessException
                Catch ex As ArgumentException
                End Try
            End If
            If IsArray(ArrayTemplatesExcel) = True Then
                ComboBox4.DataSource = ArrayTemplatesExcel
            End If
        Else
            'форма таблицы
            Dim NameTables As String() = {"Трасса автомобильной дороги"}
            Dim NameStylesTable As String() = Nothing
            Dim boolReadStyleTbl As Boolean = AcadFunction.FuncReadNameObjectAllDestDWG(PatchFileTemplate, NameTables, NameStylesTable, "Таблица")
            If IsNothing(NameStylesTable) = False Then
                If IsArray(NameStylesTable) = True Then
                    ComboBox4.DataSource = NameStylesTable
                End If
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub

    Private Sub RoadPlane_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged

    End Sub
End Class