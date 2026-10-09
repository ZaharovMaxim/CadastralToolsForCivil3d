Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Public Class ReplaseObj3DToCoGoPoint
    Inherits System.Windows.Forms.Form
    Public FormLoad1 As Boolean = False
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
        ReDim keyRaw(0)
        keyRaw(0) = ""
        CountCode += 1
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
        ComboBox2.DataSource = keyRaw
        NumericUpDown3.Value = FuncSurvey.FuncFindMaxNumberCoGoPoint + 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormLoad1 = False
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FormLoad1 = True
        Me.Hide()
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        Dim ListLayer As String() = Nothing
        If CheckBox6.Checked = True Then
            Dim CountLayer As Integer = AcadFunction.FuncReadLayersDWG(ListLayer)
            ComboBox1.Enabled = True
            ComboBox1.DataSource = ListLayer
        Else
            ComboBox1.Enabled = False
            Erase ListLayer
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            NumericUpDown1.Enabled = True
            NumericUpDown2.Enabled = True
        Else
            NumericUpDown1.Enabled = False
            NumericUpDown2.Enabled = False
        End If
    End Sub

    Private Sub ReplaseObj3DToCoGoPoint_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class