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
Public Class ProjProfileWorkingMark
    Inherits System.Windows.Forms.Form
    Public boolProf As Boolean = False
    Public ViewProfileObj As ProfileView = Nothing
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        'загружаем все профиля в форму ComboBox
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = acDoc.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'выбираем вид профиля
            Dim selectWiewProfileResult As PromptEntityResult = m_Editor.GetEntity(vbCrLf + "Выберите вид профиля:")
            If selectWiewProfileResult.Status = PromptStatus.OK Then
                Dim dbObj As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(selectWiewProfileResult.ObjectId, OpenMode.ForRead)
                ViewProfileObj = TryCast(dbObj, Autodesk.Civil.DatabaseServices.ProfileView)
                If IsNothing(ViewProfileObj) = True Then
                    MsgBox("Ошибка. Необходимо выбрать ВИД ПРОФИЛЯ. Макро прервано!!!")
                    Exit Sub
                End If
            Else
                    MsgBox("Ошибка выбора данных. Макро прервано!!!")
                Exit Sub
            End If
            Dim idAlign As ObjectId = ViewProfileObj.AlignmentId
            Dim Align As Alignment = tr.GetObject(idAlign, OpenMode.ForRead)
            Dim objProfileColl As ObjectIdCollection = Align.GetProfileIds
            If objProfileColl.Count > 1 Then
                Dim ArrayRedProfileName As String() = Nothing
                Dim ArrayBlaskProfileName As String() = Nothing
                Dim ArrayDopProfileName As String() = Nothing
                Dim CountArrayProfileName As Integer = 0
                ReDim Preserve ArrayDopProfileName(CountArrayProfileName)
                ArrayDopProfileName(CountArrayProfileName) = ""
                CountArrayProfileName += 1
                For Each id As ObjectId In objProfileColl
                    Dim ObjProfile As Profile = tr.GetObject(id, OpenMode.ForRead)
                    If ObjProfile.Length > 0 Then
                        ReDim Preserve ArrayDopProfileName(CountArrayProfileName)
                        ArrayDopProfileName(CountArrayProfileName) = ObjProfile.Name

                        ReDim Preserve ArrayRedProfileName(CountArrayProfileName - 1)
                        ArrayRedProfileName(CountArrayProfileName - 1) = ObjProfile.Name

                        ReDim Preserve ArrayBlaskProfileName(CountArrayProfileName - 1)
                        ArrayBlaskProfileName(CountArrayProfileName - 1) = ObjProfile.Name
                        CountArrayProfileName += 1
                    End If
                Next
                If ArrayBlaskProfileName.Length > 0 Then
                    ComboBox2.DataSource = ArrayRedProfileName
                    ComboBox2.SelectedItem = 0
                    ComboBox3.DataSource = ArrayBlaskProfileName
                    ComboBox3.SelectedItem = 1
                    ComboBox7.DataSource = ArrayDopProfileName
                Else
                    MsgBox("На виде профиля содержится менее 2 линии профиля. Макро прервано!!!")
                    Exit Sub
                End If
            Else
                MsgBox("На виде профиля содержится менее 2 линии профиля. Макро прервано!!!")
                Exit Sub
            End If
            '=========================================================================================================================
            'стили меток
            Dim LabelColl As Styles.LabelStyleCollection = CivilDoc.Styles.LabelStyles.ProfileViewLabelStyles.StationElevationLabelStyles
            If LabelColl.Count > 0 Then
                Dim ArrayLabel As String() = Nothing
                Dim CountArrayLabel As Integer = 0
                For Each id As ObjectId In LabelColl
                    Dim LabelEl As Styles.LabelStyle = tr.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve ArrayLabel(CountArrayLabel)
                    ArrayLabel(CountArrayLabel) = LabelEl.Name
                    CountArrayLabel += 1
                Next
                If IsArray(ArrayLabel) = True Then
                    ComboBox6.DataSource = ArrayLabel
                End If
            End If
            '=========================================================================================================================
            'стили маркера
            Dim MarkerColl As Styles.MarkerStyleCollection = CivilDoc.Styles.MarkerStyles
            If MarkerColl.Count > 0 Then
                Dim ArrayMarker As String() = Nothing
                Dim CountArrayMarker As Integer = 0
                For Each id As ObjectId In MarkerColl
                    Dim MarkerEl As Styles.MarkerStyle = tr.GetObject(id, OpenMode.ForRead)
                    ReDim Preserve ArrayMarker(CountArrayMarker)
                    ArrayMarker(CountArrayMarker) = MarkerEl.Name
                    CountArrayMarker += 1
                Next
                If IsArray(ArrayMarker) = True Then
                    ComboBox4.DataSource = ArrayMarker
                End If
            End If

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

    Private Sub ProjProfileWorkingMark_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'справка
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        System.Diagnostics.Process.Start("https://youtu.be/OyNLxopARBc")
    End Sub
End Class