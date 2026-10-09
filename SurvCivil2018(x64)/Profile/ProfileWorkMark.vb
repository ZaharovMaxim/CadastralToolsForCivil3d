Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class ProfileWorkMark
    Inherits System.Windows.Forms.Form
    Public boolProf As Boolean = False
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        'загружаем все профиля в форму ComboBox
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'наносим слои
            Dim ArrayLayer As String() = Nothing
            Dim IDAddLayer As ObjectId = AcadFunction.FuncAddLayer("А-ПП_Отметки_рабочие", 10,, 20)
            Dim FindLayer As Integer = AcadFunction.FuncReadLayersDWG(ArrayLayer)
            ComboBox6.DataSource = ArrayLayer
            For i As Integer = 0 To ArrayLayer.Length - 1
                Dim LayerText As String = ArrayLayer(i)
                If LayerText Like "А-ПП_Отметки_рабочие" Then
                    ComboBox6.Text = "А-ПП_Отметки_рабочие"
                End If
            Next
            'наносим стили текста
            Dim ArrayTextStyle As String() = Nothing
            Dim FindTextStyle As Integer = AcadFunction.FuncReadListTextStyle(ArrayTextStyle)
            ComboBox4.DataSource = ArrayTextStyle
            ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            For i As Integer = 0 To ArrayTextStyle.Length - 1
                Dim StyleText As String = ArrayTextStyle(i)
                If StyleText Like "ESKDU" Then
                    ComboBox4.Text = "ESKDU"
                End If
            Next
            'точность
            Dim HFormat As String() = {"0", "0.0", "0.00", "0.000", "0.0000"}
            ComboBox5.DataSource = HFormat
            ComboBox5.Text = "0.00"
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка отменить
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolProf = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка ок
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        boolProf = True
        Me.Hide()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub ProfileWorkMark_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class