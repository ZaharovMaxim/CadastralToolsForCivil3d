Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices.Styles

Public Class FormDrawPipeNetwork
    Public formShow As Boolean = False
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
            Dim arrayPartList As String() = {""}
            Dim countArrayPartList As Integer = 1
            Dim partLictCol As PartsListCollection = CivilDoc.Styles.PartsListSet
            If partLictCol.Count > 0 Then
                For Each idPartList As ObjectId In partLictCol
                    Dim Plist As PartsList = idPartList.GetObject(OpenMode.ForRead)
                    ReDim Preserve arrayPartList(countArrayPartList)
                    arrayPartList(countArrayPartList) = Plist.Name
                    countArrayPartList += 1
                Next
            End If
            ComboBox2.DataSource = arrayPartList
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'трасса
            Dim arrayNameAlign As String() = {""}
            Dim countArrayNameAlign As Integer = 1
            Dim AlignCool As ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCool.Count > 0 Then
                For Each idAlign As ObjectId In AlignCool
                    Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                    ReDim Preserve arrayNameAlign(countArrayNameAlign)
                    arrayNameAlign(countArrayNameAlign) = Align.Name
                    countArrayNameAlign += 1
                Next
            End If
            ComboBox7.DataSource = arrayNameAlign

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
            ComboBox8.DataSource = arrayNameSurf
            Dim arrayNameSurf2 As String() = {""}
            Dim countArrayNameSurf2 As Integer = 1
            Dim SurfCool2 As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If SurfCool2.Count > 0 Then
                For Each idSurf2 As ObjectId In SurfCool2
                    Dim Surf2 As Autodesk.Civil.DatabaseServices.TinSurface = idSurf2.GetObject(OpenMode.ForRead)
                    ReDim Preserve arrayNameSurf2(countArrayNameSurf2)
                    arrayNameSurf2(countArrayNameSurf2) = Surf2.Name
                    countArrayNameSurf2 += 1
                Next
            End If
            ComboBox9.DataSource = arrayNameSurf2
        End Using
    End Sub
    Private Sub FormDrawPipeNetwork_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        formShow = True
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        formShow = False
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
    'выбор категории
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database


        Dim kategory As String = ComboBox2.Text
        Dim arrayPartStruct As String() = {""}
        Dim arrayPartPipe As String() = {""}
        Dim countArrayPipe As Integer = 1
        Dim countArrayStruct As Integer = 1
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim oPartsListId As ObjectId = CivilDoc.Styles.PartsListSet.Item(kategory) 'Standard PartsList
            Dim oPartsList As PartsList = tr.GetObject(oPartsListId, OpenMode.ForWrite)
            Dim kategoryPipe As ObjectIdCollection = oPartsList.GetPartFamilyIdsByDomain(Autodesk.Civil.DatabaseServices.DomainType.Pipe)
            If kategoryPipe.Count > 0 Then
                For i As Integer = 0 To kategoryPipe.Count - 1
                    Dim IDoidPipe As ObjectId = kategoryPipe.Item(i)
                    Dim opfPipe As PartFamily = tr.GetObject(IDoidPipe, OpenMode.ForWrite)
                    ReDim Preserve arrayPartPipe(countArrayPipe)
                    arrayPartPipe(countArrayPipe) = opfPipe.Name
                    countArrayPipe += 1
                Next
            End If
            ComboBox3.DataSource = arrayPartPipe
            Dim kategoryStruct As ObjectIdCollection = oPartsList.GetPartFamilyIdsByDomain(Autodesk.Civil.DatabaseServices.DomainType.Structure)
            If kategoryStruct.Count > 0 Then
                For i As Integer = 0 To kategoryStruct.Count - 1
                    Dim IDoidStruct As ObjectId = kategoryStruct.Item(i)
                    Dim opfStructure As PartFamily = tr.GetObject(IDoidStruct, OpenMode.ForWrite)
                    ReDim Preserve arrayPartStruct(countArrayStruct)
                    arrayPartStruct(countArrayStruct) = opfStructure.Name
                    countArrayStruct += 1
                Next
            End If
            ComboBox5.DataSource = arrayPartStruct
        End Using


    End Sub
    'трубы выбор размера

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged





    End Sub

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim kategory As String = ComboBox2.Text
        Dim kategoryPipe As String = ComboBox3.Text
        Dim arrayPartPipe As String() = {""}
        Dim countArrayPipe As Integer = 1
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim NamePartsList As String = kategory
            Dim oPartsListId As ObjectId = CivilDoc.Styles.PartsListSet.Item(NamePartsList) 'Standard PartsList
            Dim oPartsList As PartsList = tr.GetObject(oPartsListId, OpenMode.ForWrite)

            Dim NameoldPipe As String = kategoryPipe
            Dim IDoidPipe As ObjectId = oPartsList.Item(NameoldPipe)
            Dim opfPipe As PartFamily = tr.GetObject(IDoidPipe, OpenMode.ForWrite)

            For i As Integer = 0 To opfPipe.PartSizeCount - 1
                Dim idPSizePipe As ObjectId = opfPipe.Item(i)
                Dim PSizePipe As PartSize = tr.GetObject(idPSizePipe, OpenMode.ForWrite)
                ReDim Preserve arrayPartPipe(countArrayPipe)
                arrayPartPipe(countArrayPipe) = PSizePipe.Name
                countArrayPipe += 1
            Next i
            ComboBox4.DataSource = arrayPartPipe
        End Using
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim kategory As String = ComboBox2.Text
        Dim kategoryStruct As String = ComboBox5.Text
        Dim arrayPartStruct As String() = {""}
        Dim countArrayStruct As Integer = 1
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim NamePartsList As String = kategory
            Dim oPartsListId As ObjectId = CivilDoc.Styles.PartsListSet.Item(NamePartsList) 'Standard PartsList
            Dim oPartsList As PartsList = tr.GetObject(oPartsListId, OpenMode.ForWrite)

            Dim NameoldStruct As String = kategoryStruct
            Dim IDoidStruct As ObjectId = oPartsList.Item(NameoldStruct)
            Dim opfStruct As PartFamily = tr.GetObject(IDoidStruct, OpenMode.ForWrite)

            For i As Integer = 0 To opfStruct.PartSizeCount - 1
                Dim idPSizeStruct As ObjectId = opfStruct.Item(i)
                Dim PSizeStruct As PartSize = tr.GetObject(idPSizeStruct, OpenMode.ForWrite)
                ReDim Preserve arrayPartStruct(countArrayStruct)
                arrayPartStruct(countArrayStruct) = PSizeStruct.Name
                countArrayStruct += 1
            Next i
            ComboBox6.DataSource = arrayPartStruct
        End Using
    End Sub
End Class