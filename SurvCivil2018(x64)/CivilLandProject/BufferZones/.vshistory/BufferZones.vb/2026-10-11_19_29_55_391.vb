Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Buffer
Public Class BufferZones
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'объединение буферных зон
    <CommandMethod("UnionBufferZones")>
    Public Sub UnionBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim boolAlign As Boolean = False
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            '=========================================================================================================
            'выбираем полилинии
            Dim polyIsh As Polyline = Nothing
            Dim startPolygon As Polygon = Nothing
            Dim listSelectIDCounters As List(Of ObjectId) = New List(Of ObjectId)
            Do
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim acSSPrompt As PromptEntityResult = acDoc.Editor.GetEntity(vbLf & "Выберите замкнутую полилинию: ")
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim idPline As ObjectId = acSSPrompt.ObjectId
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idPline, OpenMode.ForWrite)
                        If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim acPolyLine As Polyline = acEnt
                            If IsNothing(startPolygon) = True Then
                                startPolygon = netTopo.CreatePolygon(acPolyLine)
                                polyIsh = acPolyLine
                                listSelectIDCounters.Add(acPolyLine.Id)
                            Else
                                Dim secondPolygon As Polygon = netTopo.CreatePolygon(acPolyLine)
                                If netTopo.CanUnion(startPolygon, secondPolygon) = False Then
                                    MsgBox("Нельзя объеденить полигоны!!!. Выберите другой контур или прервите команду!!!")
                                Else
                                    listSelectIDCounters.Add(acPolyLine.Id)
                                    Dim geometryUnion As Geometry = netTopo.UnionTwoPolygons(startPolygon, secondPolygon)
                                    Dim countPolygon As Integer = geometryUnion.Boundary.NumGeometries
                                    If countPolygon > 0 Then
                                        For i As Integer = 0 To countPolygon - 1
                                            Dim tempGeo As Geometry = geometryUnion.Boundary.GetGeometryN(i)
                                            Dim coordPoly As List(Of Coordinate) = tempGeo.Coordinates.ToList
                                            Dim arrayCoord As Double(,) = {}
                                            Dim countArrayCoord As Integer = 0
                                            For j As Integer = 0 To coordPoly.Count - 1
                                                Dim coodPoint As Coordinate = coordPoly.Item(j)
                                                ReDim Preserve arrayCoord(1, countArrayCoord)
                                                arrayCoord(0, countArrayCoord) = coodPoint.X
                                                arrayCoord(1, countArrayCoord) = coodPoint.Y
                                                countArrayCoord += 1
                                            Next j
                                            Dim polyID As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                            Dim rezPoly As Polyline = tr.GetObject(polyID, OpenMode.ForWrite)
                                            rezPoly.ColorIndex = polyIsh.ColorIndex
                                            rezPoly.Linetype = polyIsh.Linetype
                                            rezPoly.LinetypeScale = polyIsh.LinetypeScale
                                            rezPoly.LineWeight = polyIsh.LineWeight
                                            Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(rezPoly)
                                            If setIds.Count > 0 Then
                                                For Each psId As ObjectId In setIds
                                                    Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.AddPropertySet(rezPoly, psId)
                                                Next
                                            End If
                                            If i = 0 Then
                                                startPolygon = netTopo.CreatePolygon(rezPoly)
                                                listSelectIDCounters.Add(rezPoly.Id)
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If listSelectIDCounters.Count > 1 Then
                            For i As Integer = 0 To listSelectIDCounters.Count - 2
                                Dim plineErase As Polyline = tr.GetObject(listSelectIDCounters.Item(i), OpenMode.ForWrite)
                                plineErase.Erase()
                            Next
                        End If
                        tr.Commit()
                        Exit Do
                    End If
                    tr.Commit()
                End Using
            Loop
        Catch ex As System.Exception
        End Try
    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////
    'обрезка зоны по полосе отвода
    <CommandMethod("DifferenceBufferZones")>
    Public Sub DifferenceBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim boolAlign As Boolean = False
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
        Dim startPolygon As Polygon = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию по которой будет происходить обрезка: ")
        If Not (rs.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        Dim poly As Polyline = New Polyline
                        Using startPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                            startPolygon = netTopo.CreatePolygon(startPoly)
                            Exit For
                        End Using
                    End If
                Next
            End If
            If startPolygon.Coordinates.Count > 2 Then
                Dim ptOpts As PromptSelectionOptions
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim pLineParcel As Polyline = acEnt
                                    Dim secondPolygon As Polygon = netTopo.CreatePolygon(pLineParcel)
                                    Dim boolIntersectPolygon As Boolean = netTopo.DoPolygonsIntersect(startPolygon, secondPolygon)
                                    If boolIntersectPolygon = True Then
                                        Dim geometryUnion As Geometry = netTopo.GetIntersection(startPolygon, secondPolygon)
                                        Dim allPolygons As List(Of Polygon) = netTopo.GetPolygonsOutsidePoly1(startPolygon, secondPolygon)
                                        If allPolygons.Count > 0 Then
                                            For i As Integer = 0 To allPolygons.Count - 1
                                                Dim workPolygon As Polygon = allPolygons.Item(i)
                                                Dim coordPoly As List(Of Coordinate) = workPolygon.Coordinates.ToList
                                                Dim arrayCoord As Double(,) = {}
                                                Dim countArrayCoord As Integer = 0
                                                For j As Integer = 0 To coordPoly.Count - 1
                                                    Dim coodPoint As Coordinate = coordPoly.Item(j)
                                                    ReDim Preserve arrayCoord(1, countArrayCoord)
                                                    arrayCoord(0, countArrayCoord) = coodPoint.X
                                                    arrayCoord(1, countArrayCoord) = coodPoint.Y
                                                    countArrayCoord += 1
                                                Next j
                                                Dim polyID As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                                Dim rezPoly As Polyline = tr.GetObject(polyID, OpenMode.ForWrite)
                                                rezPoly.ColorIndex = pLineParcel.ColorIndex
                                                rezPoly.Linetype = pLineParcel.Linetype
                                                rezPoly.LinetypeScale = pLineParcel.LinetypeScale
                                                rezPoly.LineWeight = pLineParcel.LineWeight
                                                Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(rezPoly)
                                                If setIds.Count > 0 Then
                                                    For Each psId As ObjectId In setIds
                                                        Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.AddPropertySet(rezPoly, psId)
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If
                                    pLineParcel.Erase()
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'удалить лишние вершины
    <CommandMethod("DeleteVertexZones")>
    Public Sub DeleteVertexZona()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ed As Editor = acDoc.Editor
        '==============================================================================
        'открываем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '=======================================================================================================
            'выбираем полилинию
            Dim acBFEnt As PromptEntityResult = acDoc.Editor.GetEntity(vbLf & "Выберите буферную зону: ")
            If acBFEnt.Status = PromptStatus.OK Then
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acBFEnt.ObjectId, OpenMode.ForWrite)
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim PlineBF As Autodesk.AutoCAD.DatabaseServices.Polyline = ent
                    '========================================================================================================
                    'выбираем перевую точку
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    '' Prompt for the start point
                    pPtOpts.Message = vbLf & "Выберите начальную точку: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim TStartPoint As Point3d = pPtRes.Value
                    TStartPoint = PlineBF.GetClosestPointTo(TStartPoint, False)
                    '==========================================================================================================
                    'выбираем вторую точку
                    pPtOpts.Message = vbLf & "Выберите конечную точку: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    Dim TEndPoint As Point3d = pPtRes.Value
                    TEndPoint = PlineBF.GetClosestPointTo(TEndPoint, False)
                    '===========================================================================================================
                    'вычисляем рассояния до выбранных точек
                    Dim DistStart As Double = PlineBF.GetDistAtPoint(TStartPoint)
                    Dim DistEnd As Double = PlineBF.GetDistAtPoint(TEndPoint)
                    If DistStart > DistEnd Then
                        Dim tempDist As Double = DistStart
                        DistStart = DistEnd
                        DistEnd = tempDist
                        Dim tempPointStart As Point3d = TEndPoint
                        TEndPoint = TStartPoint
                        TStartPoint = tempPointStart
                    End If
                    '============================================================================================================
                    'добавляем выбранные вершины в полилинию если конечно их нет
                    Dim vert1 As Point3d = Nothing
                    Dim vert2 As Point3d = Nothing
                    For i As Integer = 0 To PlineBF.NumberOfVertices - 1
                        If i = PlineBF.NumberOfVertices - 1 Then
                            vert1 = PlineBF.GetPoint3dAt(i)
                            vert2 = PlineBF.GetPoint3dAt(0)
                        Else
                            vert1 = PlineBF.GetPoint3dAt(i)
                            vert2 = PlineBF.GetPoint3dAt(i + 1)
                        End If

                        If vert1.GetVectorTo(vert2).Length > 0 Then
                            Dim tempdistT1 As Double = PlineBF.GetDistAtPoint(vert1)
                            Dim tempdistT2 As Double = PlineBF.GetDistAtPoint(vert2)
                            If DistStart > tempdistT1 And DistStart < tempdistT2 Then
                                PlineBF.AddVertexAt(i + 1, TStartPoint.Convert2d(New Plane), 0, 0, 0)
                            ElseIf DistEnd > tempdistT1 And DistEnd < tempdistT2 Then
                                PlineBF.AddVertexAt(i + 1, TEndPoint.Convert2d(New Plane), 0, 0, 0)
                            End If
                        End If
                    Next i

                    '===========================================================================================================
                    'пробегаем циклом по всем вершинам и удаляем те что попали внутрь промежутка
linenext:
                    Dim boolVertDel As Boolean = False
                    For i As Integer = 0 To PlineBF.NumberOfVertices - 1
                        Dim tempT As Point3d = PlineBF.GetPoint3dAt(i)
                        If tempT.GetVectorTo(TStartPoint).Length < 0.01 Then
                            boolVertDel = True
                            Continue For
                        ElseIf tempT.GetVectorTo(TEndPoint).Length < 0.01 Then
                            boolVertDel = False
                            Exit For
                        End If
                        If boolVertDel = True Then
                            PlineBF.RemoveVertexAt(i)
                            GoTo linenext
                        End If
                    Next i
                End If
            End If
            acTrans.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'разрезать полигон полилинией
    <CommandMethod("SplitBufferZones")>
    Public Sub SplitBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim boolAlign As Boolean = False
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
        Dim startPolygon As Polygon = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию которую собираетесь обрезать: ")
        If Not (rs.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        Dim poly As Polyline = New Polyline
                        Using startPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                            startPolygon = netTopo.CreatePolygon(startPoly)
                            Exit For
                        End Using
                    End If
                Next
            End If
            If startPolygon.Coordinates.Count > 2 Then
                Dim ptOpts As PromptSelectionOptions
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim pLineParcel As Polyline = acEnt
                                    Dim secondPolygon As Polygon = netTopo.CreatePolygon(pLineParcel)
                                    Dim boolIntersectPolygon As Boolean = netTopo.DoPolygonsIntersect(startPolygon, secondPolygon)
                                    If boolIntersectPolygon = True Then
                                        Dim geometryUnion As Geometry = netTopo.GetIntersection(startPolygon, secondPolygon)
                                        Dim allPolygons As List(Of Polygon) = netTopo.GetPolygonsOutsidePoly1(startPolygon, secondPolygon)
                                        If allPolygons.Count > 0 Then
                                            For i As Integer = 0 To allPolygons.Count - 1
                                                Dim workPolygon As Polygon = allPolygons.Item(i)
                                                Dim coordPoly As List(Of Coordinate) = workPolygon.Coordinates.ToList
                                                Dim arrayCoord As Double(,) = {}
                                                Dim countArrayCoord As Integer = 0
                                                For j As Integer = 0 To coordPoly.Count - 1
                                                    Dim coodPoint As Coordinate = coordPoly.Item(j)
                                                    ReDim Preserve arrayCoord(1, countArrayCoord)
                                                    arrayCoord(0, countArrayCoord) = coodPoint.X
                                                    arrayCoord(1, countArrayCoord) = coodPoint.Y
                                                    countArrayCoord += 1
                                                Next j
                                                Dim polyID As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                                Dim rezPoly As Polyline = tr.GetObject(polyID, OpenMode.ForWrite)
                                                rezPoly.ColorIndex = pLineParcel.ColorIndex
                                                rezPoly.Linetype = pLineParcel.Linetype
                                                rezPoly.LinetypeScale = pLineParcel.LinetypeScale
                                                rezPoly.LineWeight = pLineParcel.LineWeight
                                                Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(rezPoly)
                                                If setIds.Count > 0 Then
                                                    For Each psId As ObjectId In setIds
                                                        Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.AddPropertySet(rezPoly, psId)
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If
                                    pLineParcel.Erase()
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            tr.Commit()
        End Using
    End Sub


End Class
