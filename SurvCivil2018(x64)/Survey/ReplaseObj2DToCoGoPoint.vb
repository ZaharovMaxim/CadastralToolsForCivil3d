Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Public Class ReplaseObj2DToCoGoPoint
    Inherits System.Windows.Forms.Form
    Public FormLoad1 As Boolean = False
    Public SelectionSetObj As SelectionSet = Nothing
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim keyRaw As String() = Nothing
        Dim CountCode As Integer = 0
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Dim pdKeySetIds As ObjectIdCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acDoc.Database).SearchOrder
            For i As Integer = 0 To pdKeySetIds.Count - 1
                Dim keySetId As ObjectId = pdKeySetIds(i)
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(keySetId, OpenMode.ForRead)
                Dim DeskKeySet As PointDescriptionKeySet = obj
                Dim pdKeyIds As ObjectIdCollection = DeskKeySet.GetPointDescriptionKeyIds
                For j As Integer = 0 To pdKeyIds.Count - 1
                    Dim keyId As ObjectId = pdKeyIds(j)
                    Dim objkey As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(keyId, OpenMode.ForRead)
                    Dim DeskKey As PointDescriptionKey = objkey
                    Dim Name1 As String = DeskKey.Code
                    ReDim Preserve keyRaw(CountCode)
                    keyRaw(CountCode) = Name1
                    CountCode += 1
                Next j
            Next i
        End Using
        ComboBox1.DataSource = keyRaw
        NumericUpDown4.Value = FuncSurvey.FuncFindMaxNumberCoGoPoint + 1
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim SelectPline As PromptEntityResult = ed.GetEntity("Выберите объект (блок, точка, окружность):  ")
            If SelectPline.Status = PromptStatus.OK Then
                Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(SelectPline.ObjectId, OpenMode.ForWrite)
                'забираем в набор блоки согласно образца
                If TypeOf acEnt1 Is BlockReference Then
                    Dim blkEnt As BlockReference = acEnt1
                    TextBox1.Text = blkEnt.Name
                ElseIf TypeOf acEnt1 Is DBPoint Then
                    Dim point1 As DBPoint = acEnt1
                    TextBox1.Text = "Точка AutoCAD"
                ElseIf TypeOf acEnt1 Is Circle Then
                    Dim cir1 As DBPoint = acEnt1
                    TextBox1.Text = "Окружность AutoCAD"
                Else
                    MsgBox("Тип объекта выбран неверно. Выберите блок, точку, либо окружность!!!")
                End If
            End If
            tr.Commit()
        End Using
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim resPoint As PromptPointResult = Nothing
            Dim optPoint As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние от объекта до текста (первая точка): ")
            resPoint = ed.GetPoint(optPoint)
            Dim Point1 As Point3d = resPoint.Value

            Dim optPoint1 As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние от объекта до текста (вторая точка): ")
            resPoint = ed.GetPoint(optPoint)
            Dim Point2 As Point3d = resPoint.Value

            Dim Lenght As Double = Math.Sqrt((Point1.X - Point2.X) ^ 2 + (Point1.Y - Point2.Y) ^ 2)
            NumericUpDown1.Value = Math.Abs(Lenght)
        End Using
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FormLoad1 = False
        Me.Hide()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        FormLoad1 = True
        Me.Hide()
    End Sub


    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        Dim ListLayer As String() = Nothing
        If CheckBox6.Checked = True Then
            Dim CountLayer As Integer = AcadFunction.FuncReadLayersDWG(ListLayer)
            ComboBox2.Enabled = True
            ComboBox2.DataSource = ListLayer
        Else
            ComboBox2.Enabled = False
            Erase ListLayer
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class