Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices.Styles
Public Class FormInsertPortalWall
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim arrayBlk As String() = {""}
            Dim numBlk As Integer = AcadFunction.FuncReadListBlockDef(arrayBlk)
            ' Добавить код инициализации после вызова InitializeComponent().
            ComboBox1.DataSource = arrayBlk
            ' Добавить код инициализации после вызова InitializeComponent().
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'поверхность для колодцев
            Dim arrayNameSurf As String() = {""}
            Dim countArrayNameSurf As Integer = 1
            Dim SurfCool As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If SurfCool.Count > 0 Then
                For Each idSurf As ObjectId In SurfCool
                    Dim Surf As Autodesk.Civil.DatabaseServices.TinSurface = idSurf.GetObject(OpenMode.ForRead)
                    ReDim Preserve arrayNameSurf(countArrayNameSurf)
                    arrayNameSurf(countArrayNameSurf) = Surf.Name
                    countArrayNameSurf += 1
                Next
            End If
            ComboBox2.DataSource = arrayNameSurf
        End Using

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub FormInsertPortalWall_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class