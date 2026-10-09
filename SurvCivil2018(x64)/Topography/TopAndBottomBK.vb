Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Public Class TopAndBottomBK
    Public boolShow As Boolean = False
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
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            ComboBox1.Tag = PutchTables
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
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox1.DataSource = ArrayTemplates
        End If


    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TopAndBottomBK_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class