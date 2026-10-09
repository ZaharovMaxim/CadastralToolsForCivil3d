Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Drawing
Imports System.IO
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Windows.OpenFileDialog

Public Class FormListProjectParcels
    Public boolShow As Boolean = False
    Public Sub New()
        InitializeComponent()
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам таблиц
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
            If IO.Directory.Exists(PutcTableDoc) Then
                For Each fileDwg As String In Directory.GetFiles(PutcTableDoc)
                    Try
                        Dim File1 As String = Path.GetFileName(fileDwg)
                        Dim ext As String = Path.GetExtension(fileDwg)
                        If ext Like ".docx" Then
                            ReDim Preserve ArrayTemplates(CountArrayTemplates)
                            ArrayTemplates(CountArrayTemplates) = File1
                            CountArrayTemplates += 1
                        End If
                    Catch ex As ArgumentException
                    End Try
                Next
                If IsArray(ArrayTemplates) = True Then
                    ComboBox1.DataSource = ArrayTemplates
                    ComboBox1.Tag = PutcTableDoc
                End If
            End If
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Директория с файлами-шаблонами не найдена!!!")
                Exit Sub
            End If
        End If
    End Sub
    'OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
    'Cancel
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор файла
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim NameFiles As String() = {""}
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        If RadioButton1.Checked = True Then
            Dim fileDialog1 As OpenFileDialog = New OpenFileDialog("Файл Microsoft Word", "", "docx", "docx", OpenFileDialogFlags.AllowMultiple)
            Dim FileRez As String = fileDialog1.ShowDialog()
            NameFiles = fileDialog1.GetFilenames
        ElseIf RadioButton2.Checked = True Then
            Dim fileDialog1 As OpenFileDialog = New OpenFileDialog("Файл Microsoft Excel", "", "xlsx", "xlsx", OpenFileDialogFlags.AllowMultiple)
            Dim FileRez As String = fileDialog1.ShowDialog()
            NameFiles = fileDialog1.GetFilenames
        End If
        If IsNothing(NameFiles) = True Then Exit Sub
        For i As Integer = 0 To NameFiles.Length - 1
            Try
                Dim File1 As String = Path.GetFileName(NameFiles(i))
                Dim ext As String = Path.GetExtension(NameFiles(i))
                If ext Like ".docx" Then
                    ReDim Preserve ArrayTemplates(CountArrayTemplates)
                    ArrayTemplates(CountArrayTemplates) = File1
                    CountArrayTemplates += 1
                End If
                ComboBox1.Tag = Path.GetDirectoryName(NameFiles(i))
            Catch ex As ArgumentException
            End Try
            If IsArray(ArrayTemplates) = True Then
                ComboBox1.DataSource = ArrayTemplates
            End If
        Next i
    End Sub
End Class