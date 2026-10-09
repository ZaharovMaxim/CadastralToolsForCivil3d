Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Public Class SplitFigure
    Public boolStatus As Boolean = False 'нажата кнопка OK
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
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        ''читаем шаблоны условных знаков
        'If IO.Directory.Exists(PutchTables) = True Then
        '    ComboBox2.Tag = PutchTables
        '    Try
        '        For Each fileDWG As String In Directory.GetFiles(PutchTables)
        '            'читаем расширение
        '            Dim ext As String = Path.GetExtension(fileDWG)
        '            If ext Like ".dwg" Then
        '                ReDim Preserve ArrayTemplates(CountArrayTemplates)
        '                ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileDWG)
        '                If CountArrayTemplates = 0 Then
        '                    PatchFileTemplate = fileDWG
        '                End If
        '                CountArrayTemplates += 1
        '            End If
        '        Next
        '    Catch ex As IOException
        '    Catch ex As UnauthorizedAccessException
        '    Catch ex As ArgumentException
        '    End Try
        'End If
        'If IsArray(ArrayTemplates) = True Then
        '    ComboBox2.DataSource = ArrayTemplates
        'End If
        TextBox3.Text = "FIGURE_CORIDOR"
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
                            Dim startPKstr As String = Align.GetStationStringWithEquations(Align.StartingStation)
                            TextBox1.Text = startPKstr
                            TextBox1.Tag = Align.StartingStation
                        Catch ex As System.ArgumentException
                            TextBox1.Text = Math.Round(Align.StartingStation, 4)
                            TextBox1.Tag = Align.StartingStation
                        End Try
                        Try
                            Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                            TextBox2.Text = EndPKStr
                            TextBox2.Tag = Align.EndingStation
                        Catch ex As System.ArgumentException
                            TextBox2.Text = Math.Round(Align.EndingStation, 4)
                            TextBox2.Tag = Align.EndingStation
                        End Try
                    End If
                    CountArrayNameAlign += 1
                Next
            End Using
        End If
        If IsArray(ArrayNameAlign) = True Then
            ComboBox1.DataSource = ArrayNameAlign
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        boolStatus = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Отмена
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'указать начало трассы
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox1.Text
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
                            If pPtRes.Status = PromptStatus.OK Then 'нет выбора
                                Try
                                    Align.StationOffset(ptStart.X, ptStart.Y, PKStartAlign, OfsetStartAlign)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Dim IDPline As ObjectId = Align.GetPolyline
                                    If IDPline.IsNull = False Then
                                        Dim Pline As Polyline = tr.GetObject(IDPline, OpenMode.ForRead)
                                        Dim PtSelect As Point3d = Pline.GetClosestPointTo(ptStart, True)
                                        Dim Dist As Double = Pline.GetDistAtPoint(PtSelect) + Align.StartingStation
                                        TextBox1.Text = FuncFormatZN.FuncFormatPK(Dist)
                                        TextBox1.Tag = Dist
                                    End If
                                    Exit Sub
                                End Try
                                Try
                                    Dim startPKstr As String = Align.GetStationStringWithEquations(PKStartAlign)
                                    TextBox1.Text = startPKstr
                                    TextBox1.Tag = PKStartAlign
                                Catch ex1 As System.ArgumentException
                                    TextBox1.Text = Align.StartingStation
                                    TextBox1.Tag = Align.StartingStation
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim NameAlign As String = ComboBox1.Text
        If NameAlign.Trim.Length > 0 Then
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCollection.Count > 0 Then
                For Each idAlign As ObjectId In AlignCollection
                    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                    Using tr
                        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                        If Align.Name Like NameAlign Then
                            Dim PKEndAlign As Double = Align.EndingStation 'пикет конца трассы
                            Dim OfsetEndAlign As Double = 0 'смещение точки от трассы
                            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                            pPtOpts.AllowNone = True
                            pPtOpts.Message = vbLf & "Укажите точку конца трассы"
                            Dim pPtRes As PromptPointResult = m_Editor.GetPoint(pPtOpts)
                            Dim ptEnd As Point3d = pPtRes.Value 'координаты точки начала трассы
                            If pPtRes.Status = PromptStatus.OK Then 'нет выбора
                                Try
                                    Align.StationOffset(ptEnd.X, ptEnd.Y, PKEndAlign, OfsetEndAlign)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Dim IDPline As ObjectId = Align.GetPolyline
                                    If IDPline.IsNull = False Then
                                        Dim Pline As Polyline = tr.GetObject(IDPline, OpenMode.ForRead)
                                        Dim PtSelect As Point3d = Pline.GetClosestPointTo(ptEnd, True)
                                        Dim Dist As Double = Pline.GetDistAtPoint(PtSelect) + Align.StartingStation
                                        TextBox1.Text = FuncFormatZN.FuncFormatPK(Dist)
                                        TextBox1.Tag = Dist
                                    End If
                                    Exit Sub
                                End Try
                                Try
                                    Dim EndPKstr As String = Align.GetStationStringWithEquations(PKEndAlign)
                                    TextBox2.Text = EndPKstr
                                    TextBox2.Tag = PKEndAlign
                                Catch ex As System.ArithmeticException
                                    TextBox2.Text = PKEndAlign
                                    TextBox2.Tag = PKEndAlign
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
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim NameAlign As String = ComboBox1.Text
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
                                Dim startPKstr As String = Align.GetStationStringWithEquations(Align.StartingStation)
                                TextBox1.Text = startPKstr
                                TextBox1.Tag = Align.StartingStation
                            Catch ex As System.ArgumentException
                                TextBox1.Text = Math.Round(Align.StartingStation, 4)
                                TextBox1.Tag = Align.StartingStation
                            End Try
                            Try
                                Dim EndPKStr As String = Align.GetStationStringWithEquations(Align.EndingStation)
                                TextBox2.Text = EndPKStr
                                TextBox2.Tag = Align.EndingStation
                            Catch ex As System.ArgumentException
                                TextBox2.Text = Math.Round(Align.EndingStation, 4)
                                TextBox2.Tag = Align.EndingStation
                            End Try
                        End If
                    Next
                End Using
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'раздел солидов по расстояниям
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            NumericUpDown1.Value = 1000
            Label1.Text = "Расстояние, м"
        End If
    End Sub
    'раздел солидов по таблице Excel
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            NumericUpDown1.Value = 1
            Label1.Text = "Номер столбца"
        End If
    End Sub

    Private Sub SplitFigure_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class