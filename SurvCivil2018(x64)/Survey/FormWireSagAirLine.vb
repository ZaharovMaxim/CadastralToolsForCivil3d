Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms
Public Class FormWireSagAirLine
    Public PointWire As Point2d = Nothing
    Public boolFormShow As Boolean = False
    Public Sub New()
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim idColSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
        Dim ArrayNameSurf As String() = Nothing
        Dim countArrayNameSurf As Integer = 0
        If idColSurf.Count > 0 Then
            Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                For Each idSurf As ObjectId In idColSurf
                    Dim Surf As TinSurface = acTrans.GetObject(idSurf, OpenMode.ForRead)
                    ReDim Preserve ArrayNameSurf(countArrayNameSurf)
                    ArrayNameSurf(countArrayNameSurf) = Surf.Name
                    countArrayNameSurf += 1
                Next
            End Using
            If IsArray(ArrayNameSurf) = True Then
                ComboBox1.DataSource = ArrayNameSurf
            Else
                RadioButton2.Checked = True
            End If
        Else
            RadioButton2.Checked = True
        End If
    End Sub
    '=======================================================================================
    'выбор объекта (текст, МТекст, Блок)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim resBlk0 As PromptEntityResult = Nothing
            Dim opt0 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите Текст, МТекст или Блок: ")
            resBlk0 = acDocEd.GetEntity(opt0)
            Dim pvId0 As ObjectId = resBlk0.ObjectId
            Dim obj0 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId0, OpenMode.ForWrite)
            If TypeOf obj0 Is Autodesk.AutoCAD.DatabaseServices.DBText Then
                Dim acText As DBText = obj0
                Dim valText As String = acText.TextString
                If IsNumeric(valText) = True Then
                    NumericUpDown1.Value = valText
                Else
                    If acText.Position.Z <> 0 Then
                        NumericUpDown1.Value = acText.Position.Z
                        RadioButton2.Checked = True
                    End If
                End If
                Dim ext As Extents3d = acText.GeometricExtents
                PointWire = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
            ElseIf TypeOf obj0 Is Autodesk.AutoCAD.DatabaseServices.MText Then
                Dim acText As MText = obj0
                Dim valText As String = acText.Contents
                If IsNumeric(valText) = True Then
                    NumericUpDown1.Value = valText
                Else
                    If acText.Location.Z <> 0 Then
                        NumericUpDown1.Value = acText.Location.Z
                        RadioButton2.Checked = True
                    End If
                End If
                Dim ext As Extents3d = acText.GeometricExtents
                PointWire = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
            ElseIf TypeOf obj0 Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                Dim acblk As BlockReference = obj0
                Dim ArrayZn As String(,) = Nothing
                Dim boolFindAttrib As Boolean = AcadFunction.FuncReadAllBlockAttr(acblk, ArrayZn)
                Dim WireZn As Double = 0
                If IsArray(ArrayZn) = True Then
                    For i As Integer = 0 To ArrayZn.GetUpperBound(1)
                        Dim TempZn As String = ArrayZn(1, i)
                        If IsNumeric(TempZn) = True Then
                            WireZn = TempZn
                            Exit For
                        End If
                    Next
                End If
                If WireZn = 0 Then
                    NumericUpDown1.Value = acblk.Position.Z
                    RadioButton2.Checked = True
                Else
                    NumericUpDown1.Value = WireZn
                End If
                PointWire = acblk.Position.Convert2d(New Plane)
            ElseIf TypeOf obj0 Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                Dim acPTCoGo As CogoPoint = obj0
                NumericUpDown1.Value = acPTCoGo.Elevation
                RadioButton2.Checked = True
                PointWire = acPTCoGo.Location.Convert2d(New Plane)
            End If
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolFormShow = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolFormShow = True
        Me.Hide()
    End Sub

    Private Sub FormWireSagAirLine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class