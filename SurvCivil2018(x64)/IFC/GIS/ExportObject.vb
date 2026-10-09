Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.EditorInput
Imports System.Runtime.InteropServices
Imports Autodesk.AEC.Modeler
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.BoundaryRepresentation
Public Class ExportObject
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==============================================================================================================================
    'создание слоев дорожной одежды (нижний слой с выступом)
    <CommandMethod("ExportAlignToGIS")>
    Public Sub ExportAlignToGIS()
        ' Get the current document and database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim putchTemplate As String = "k:\Civil3DToolsUtility\FileResources\Sample\Стандарт_ADSK.dwg"
            Dim category As String = "Проектирование автомобильных дорог"
            Dim nameTable As String = "Трасса автомобильной дороги"
            Dim nameDeskAlign As String = "Ось трассы"
            'Dim NameTypeline As String = ""
            'Dim LayerLine As String = "0"
            'Dim ColorLine As Integer = 10
            'Dim ScaleLine As Double = 1
            'Dim WidthLine As Double = 1
            'Dim ArrayPropertiesAlign As String(,) = {}
            'Dim boolReadStyleDim As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(putchTemplate, nameDeskAlign, ArrayPropertiesAlign, "Полилиния", "Трасса автомобильной дороги")
            'If IsArray(ArrayPropertiesAlign) = True Then
            '    For i As Integer = 0 To ArrayPropertiesAlign.GetUpperBound(1)
            '        If ArrayPropertiesAlign(0, i) Like "Name" Then
            '            NameTypeline = ArrayPropertiesAlign(1, i)
            '            If IsNothing(NameTypeline) = False Then
            '                Dim boolCopyBLK As Boolean = AcadFunction.FuncCopyLineStylesDestDraw(putchTemplate, NameTypeline, False)
            '            End If
            '        ElseIf ArrayPropertiesBLK(0, i) Like "Layer" Then
            '            Dim NameLayerTemp As String = ArrayPropertiesBLK(1, i)
            '            If IsNothing(NameLayerTemp) = False Then
            '                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(NameFileTemplate, NameLayerTemp)
            '                If acLyrTbl.Has(NameLayerTemp) = True Then
            '                    LayerName = NameLayerTemp
            '                End If
            '            End If
            '        ElseIf ArrayPropertiesBLK(0, i) Like "Color" Then
            '            ColorBlk = Val(ArrayPropertiesBLK(1, i))
            '            If ColorBlk < 0 Or ColorBlk > 255 Then
            '                ColorBlk = 10
            '            End If
            '        ElseIf ArrayPropertiesBLK(0, i) Like "Scale" Then
            '            Dim OriginScale As Double = Val(ArrayPropertiesBLK(1, i))
            '            If OriginScale > 0 Then
            '                ScaleBlk = OriginScale * (MAnno / 1000)
            '            End If
            '        End If
            '    Next i
            'End If


            Dim nameDeskPK As String = "Пикет"
            Dim nameDeskKM As String = "Километровый знак"


            'Dim ArrayPropertiesPK As String(,) = {}
            'Dim boolReadStyleDim1 As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(putchTemplate, nameDeskPK, ArrayPropertiesAlign, "Блок", "Трасса автомобильной дороги")
            'Dim ArrayPropertiesKM As String(,) = {}
            'Dim boolReadStyleDim2 As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(putchTemplate, nameDeskKM, ArrayPropertiesAlign, "Блок", "Трасса автомобильной дороги")
            Dim idTablePK As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("Align")

            Dim boolFindPK As Boolean = AcadFunction.FuncCopyBlockDestDraw(putchTemplate, "A-Пикет", False)
            Dim boolFindKM As Boolean = AcadFunction.FuncCopyBlockDestDraw(putchTemplate, "A-Километровый знак", False)
            Do
                Dim align As Alignment = Nothing
                Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу: ")
                If Not (acSSPrompt.Status = PromptStatus.OK) Then
                    Exit Do
                End If
                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)
                If TypeOf acEnt Is Alignment Then
                    align = acEnt
                    'преобразуем трассу в полилинию
                    Dim idPoly As ObjectId = align.GetPolyline
                    Dim polyAlign As Polyline = acTrans.GetObject(idPoly, OpenMode.ForWrite)

                    'acBlkTblRec.AppendEntity(polyAlign)
                    'acTrans.AddNewlyCreatedDBObject(polyAlign, True)
                    'находим пикеты
                    Dim myStations As Station() = align.GetStationSet(StationTypes.All, 1000, 100)
                    If IsArray(myStations) = True Then
                        For Each myStation As Station In myStations
                            Dim stType As StationTypes = myStation.StationType
                            If stType = StationTypes.Minor Then
                                Dim pK As Double = myStation.RawStation
                                Dim pos As Point2d = myStation.Location
                                Dim blkPK As BlockReference = AcadFunction.FuncInsertDrawingBlk(New Point3d(pos.X, pos.Y, 0), "A-Пикет")
                                Dim arrayAttrib As String(,) = {}
                                ReDim Preserve arrayAttrib(6, 0)
                                arrayAttrib(0, 0) = "ПК"
                                arrayAttrib(1, 0) = "ПК " & Math.Round(myStation.RawStation / 100, 0)
                                arrayAttrib(2, 0) = 2
                                arrayAttrib(3, 0) = 2
                                arrayAttrib(4, 0) = 0
                                arrayAttrib(5, 0) = ""
                                arrayAttrib(6, 0) = 0
                                Dim boolInsAttr = AcadFunction.FuncWriteAllBlockAttr(blkPK, arrayAttrib)
                                Dim ArrayData As String(,) = Nothing
                                ReDim ArrayData(2, 10)
                                ArrayData(0, 0) = "Align"
                                ArrayData(1, 0) = "Name"
                                ArrayData(2, 0) = "ПК " & Math.Round(myStation.RawStation / 100, 0)
                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(blkPK, idTablePK)
                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blkPK, ArrayData)
                            ElseIf stType = StationTypes.Major Then
                                Dim pK As Double = myStation.RawStation
                                Dim pos As Point2d = myStation.Location
                                Dim blkPK As BlockReference = AcadFunction.FuncInsertDrawingBlk(New Point3d(pos.X, pos.Y, 0), "A-Километровый знак")
                                Dim arrayAttrib As String(,) = {}
                                ReDim Preserve arrayAttrib(6, 0)
                                arrayAttrib(0, 0) = "AlignKM"
                                arrayAttrib(1, 0) = "КМ " & Math.Round(myStation.RawStation / 1000, 0)
                                arrayAttrib(2, 0) = 2
                                arrayAttrib(3, 0) = 2
                                arrayAttrib(4, 0) = 0
                                arrayAttrib(5, 0) = ""
                                arrayAttrib(6, 0) = 0
                                Dim boolInsAttr = AcadFunction.FuncWriteAllBlockAttr(blkPK, arrayAttrib)
                                Dim ArrayData As String(,) = Nothing
                                ReDim ArrayData(2, 10)
                                ArrayData(0, 0) = "Align"
                                ArrayData(1, 0) = "Name"
                                ArrayData(2, 0) = "ПК " & Math.Round(myStation.RawStation / 100, 0)
                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(blkPK, idTablePK)
                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blkPK, ArrayData)
                            End If
                        Next
                    End If
                End If
            Loop
            acTrans.Commit()
        End Using
    End Sub
End Class
