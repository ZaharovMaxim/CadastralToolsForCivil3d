Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Public Class FormReportCatalogCoordinates
    Public boolRp As Boolean = False
    Public arrayDoc As String() = Nothing
    Public WApp As Word.Application = Nothing
    Public WDoc As Word.Document = Nothing
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        'Dim ArrayExl As String() = {"Все участки в один файл", "Каждый участок в отдельный лист", "Каждый участок в отдельный файл"}
        Dim ArrayExl As String() = {""}
        ComboBox1.DataSource = ArrayExl

        Dim ArrayTablePS As String() = Nothing
        Dim boolReadPs As Boolean = FuncPropertySet.FuncReturnArrayNamesDictGWG(ArrayTablePS)
        If IsArray(ArrayTablePS) = True Then
            ComboBox3.DataSource = ArrayTablePS
        End If
        arrayDoc = {"Новый файл docx"}
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim countArrayDoc As Integer = 1
        Dim WordObj As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application()
        Try
            WordObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
        Catch ex As COMException
        End Try
        If IsNothing(WordObj) = False Then
            For i As Integer = 0 To WordObj.Windows.Count - 1
                Dim idx As Object = i + 1
                Dim WinObj As Word.Window = WordObj.Windows.Item(idx)
                ReDim Preserve arrayDoc(countArrayDoc)
                arrayDoc(countArrayDoc) = WinObj.Document.Name
                countArrayDoc += 1
            Next
        End If
        ReDim Preserve arrayDoc(countArrayDoc)
        arrayDoc(countArrayDoc) = "Выбрать файл...."
        ComboBox1.DataSource = arrayDoc

        Dim ArrayPref As String() = {"00:00:0000000:0000:ЗУ", "Участок №", "Публ.сервитут №", "Сервитут №", "00:00:0000000:0000/чзу", "Часть участка №", "№", ""}
        ComboBox4.DataSource = ArrayPref
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolRp = False
        Me.Hide()
    End Sub
    'ок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolRp = True
        Me.Hide()
    End Sub
    'выбран отчет в автокад
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Dim ArrayAutocad As String() = {"Таблица координат", "Участок в отдельный файл"}
            ComboBox1.DataSource = ArrayAutocad
        End If
    End Sub
    'выбран отчет в Excel
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Dim ArrayExl As String() = {"Все участки в один файл", "Каждый участок в отдельный лист", "Каждый участок в отдельный файл", "Для проекта межевания", "МСК-WGS84"}
            ComboBox1.DataSource = ArrayExl
        End If
    End Sub
    'выбран отчет в MapInfo
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            Dim ArrayMapInfo As String() = {"Все участки в один файл", "Каждый участок в отдельный файл", "Согласно приказа 433/пр"}
            ComboBox1.DataSource = ArrayMapInfo
        End If
    End Sub
    'выбран отчет в Технокад
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            Dim ArrayTCad As String() = {"Уточняемые координаты", "Исходные координаты"}
            ComboBox1.DataSource = ArrayTCad
        End If
    End Sub
    'выбран отчет в MSWord
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If IsArray(arrayDoc) = True Then
            ComboBox1.DataSource = arrayDoc
        Else
            arrayDoc = {"Новый файл docx", "Выбрать файл...."}
            ComboBox1.DataSource = arrayDoc
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text Like "Выбрать файл...." Then
            Dim FileNameGeo As String = ""
            OpenFileDialog1.Filter = "Все|*.*"
            OpenFileDialog1.Title = "Выбор файла"
            OpenFileDialog1.Multiselect = False
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FileNameGeo = OpenFileDialog1.FileName
            End If
            If FileNameGeo.Trim.Length > 0 Then
                If File.Exists(FileNameGeo) = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Word документ
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(FileNameGeo)
                    WApp.Visible = True
                End If
            End If
        End If
    End Sub

    Private Sub FormReportCatalogCoordinates_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class