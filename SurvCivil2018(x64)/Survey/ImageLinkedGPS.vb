Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Civil.ApplicationServices

Public Class ImageLinkedGPS
    Public boolShowDialog As Boolean = False
    Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
    Public Sub New()
        InitializeComponent()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") = True Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            ComboBox2.Tag = PutchTables
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
            ComboBox2.DataSource = ArrayTemplates
        End If
        Dim CodeCoordSys As String = CivilDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
        If CodeCoordSys.Trim.Length > 0 Then
            Dim ArrayCode As String() = {CodeCoordSys}
            ComboBox1.DataSource = ArrayCode
        End If

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбрать папку с файлами
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim opFile As New System.Windows.Forms.OpenFileDialog()
        opFile.InitialDirectory = "c:\"
        opFile.Filter = "Файлы изображений (*.JPG)|*.JPG"
        opFile.RestoreDirectory = True
        opFile.Multiselect = True
        opFile.ShowDialog()
        Dim fFile1 As String() = opFile.FileNames
        Dim CountFiles As Integer = fFile1.Length
        If CountFiles > 0 Then
            TextBox1.Text = Path.GetDirectoryName(fFile1(0))
            TextBox1.Tag = fFile1
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'каталог для экспорта изображений
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Text = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите каталог для копирования изображений", 16384)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'КНОПКА OK
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        boolShowDialog = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка ОТМЕНА
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolShowDialog = False
        Me.Hide()
    End Sub

    Private Sub ImageLinkedGPS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class