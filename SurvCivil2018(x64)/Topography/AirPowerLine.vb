Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Public Class AirPowerLine
    Public boolShow As Boolean = False
    Public PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName

        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        Dim PutchTemplateFile As String = ""
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            Try
                For Each fileDWG As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileDWG)
                        CountArrayTemplates += 1
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        Else
            MsgBox("Файл шаблона не найден. Макро превано!!!")
            Exit Sub
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox1.DataSource = ArrayTemplates
            ComboBox1.Tag = PutchTables
        End If
        Dim AnnoScale As AnnotationScale = acCurDb.Cannoscale
        Dim MDraw As Double = (annoScale.PaperUnits / annoScale.Scale) * 1000
        NumericUpDown1.Value = Int(MDraw)
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    'ok
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        ComboBox1.Tag = PutchTables
        Me.Hide()
    End Sub

    Private Sub AirPowerLine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class