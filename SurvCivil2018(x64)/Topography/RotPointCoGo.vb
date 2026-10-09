Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Public Class RotPointCoGo
    Public boolRotate As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolRotate = False
        Me.Hide()
    End Sub
    'нажата кнопка ОК
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        boolRotate = True
        Me.Hide()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            NumericUpDown2.Value = 0
            NumericUpDown3.Value = 0
        End If
    End Sub
    'выбор для поиска точек по максимальному расстоянию
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = docBase.TransactionManager.StartTransaction()
            '============================================================================
            'находим максимальное расстояние между точками
            Dim Lenght As Double = 0
            Dim resPoint As PromptPointResult = Nothing
            Dim optPoint As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние до COGO точки (первая точка): ")
            resPoint = ed.GetPoint(optPoint)
            Dim Point1 As Point3d = New Point3d(-1, -1, -1)
            If resPoint.Status = PromptStatus.OK Then
                Point1 = resPoint.Value
            Else
                GoTo line1
            End If

            Dim optPoint1 As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние до COGO точки (вторая точка): ")
            resPoint = ed.GetPoint(optPoint1)
            Dim Point2 As Point3d = New Point3d(-1, -1, -1)
            If resPoint.Status = PromptStatus.OK Then
                Point2 = resPoint.Value
            Else
                GoTo line1
            End If

            Dim pt1 As Point2d = Point1.Convert2d(New Plane)
            Dim pt2 As Point2d = Point2.Convert2d(New Plane)
            Lenght = pt1.GetVectorTo(pt2).Length
            Lenght = Math.Round(Lenght, 2)
line1:
            NumericUpDown1.Value = Lenght
        End Using
        locker.Dispose()
    End Sub
    'смещение от точки
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = docBase.TransactionManager.StartTransaction()
            Dim dx As Double = 0
            Dim dy As Double = 0
            Dim SelectCOGO As PromptEntityResult = ed.GetEntity(vbLf & "Выберите COGO точку для образца:  ")
            If SelectCOGO.Status = PromptStatus.OK Then
                Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectCOGO.ObjectId, OpenMode.ForRead)
                If TypeOf acEnt2 Is CogoPoint Then
                    Dim selectCogoTemp As CogoPoint = acEnt2
                    dx = selectCogoTemp.LabelLocation.X - selectCogoTemp.Easting
                    dy = selectCogoTemp.LabelLocation.Y - selectCogoTemp.Northing
                    NumericUpDown2.Value = dx
                    NumericUpDown3.Value = dy
                    Button2.Tag = selectCogoTemp.PointNumber
                End If
            End If
        End Using
        locker.Dispose()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class