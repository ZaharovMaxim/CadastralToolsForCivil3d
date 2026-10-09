Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports CadastralTools
Imports NetTopologySuite
Public Class SelectCadastralObject
    Public Shared Function selectAllObject(ByRef Db As Database) As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry)) = New Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncNetTopologySuite = New FuncNetTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim selectEntity As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                Dim dataStructuresObject As StructureElement = Nothing
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                If boolRecData = True Then
                    Dim nameObject As StructureElement.nameObject = dataStructuresObject.NameTypeObject
                    Dim typeGeometry As StructureElement.typeGeometryObject = dataStructuresObject.TypeGeometry
                    Dim objectGeometry As Geometries.Geometry = Nothing
                    If typeGeometry = StructureElement.typeGeometryObject.polygon Then
                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Area > 0 Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(polygon)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    ElseIf typeGeometry = StructureElement.typeGeometryObject.polyline Then
                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Length > 0 Then
                                Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                If IsNothing(line) = False Then
                                    line.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {line})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(line)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Dim insertPoint As Point3d = New Point3d()
                        If TypeOf selectEntity Is MLeader Then
                            Dim acLeader As MLeader = selectEntity
                            insertPoint = acLeader.GetFirstVertex(0)
                        ElseIf TypeOf selectEntity Is DBText Then
                            Dim acText As DBText = selectEntity
                            insertPoint = acText.Position
                        ElseIf TypeOf selectEntity Is Circle Then
                            Dim acCircle As Circle = selectEntity
                            insertPoint = acCircle.Center
                        ElseIf TypeOf selectEntity Is BlockReference Then
                            Dim acBlock As BlockReference = selectEntity
                            insertPoint = acBlock.Position
                        End If
                        Dim point As Geometries.Point = New Geometries.Point(insertPoint.X, insertPoint.Y, insertPoint.Z)
                        point.UserData = dataStructuresObject
                        If result.ContainsKey(nameObject) = False Then
                            result.Add(nameObject, New List(Of Geometries.Geometry) From {point})
                        Else
                            Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                            listObject.Add(point)
                            result.Item(nameObject) = listObject
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function
    Public Shared Function selectPolygonObject(ByRef Db As Database) As Dictionary(Of String, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of String, List(Of Geometries.Geometry)) = New Dictionary(Of String, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncNetTopologySuite = New FuncNetTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim selectEntity As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                Dim dataStructuresObject As StructureElement = Nothing
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                If boolRecData = True Then
                    Dim deskObject As String = dataStructuresObject.Description
                    Dim dictionaryPs As Dictionary(Of String, String) = FuncPropertySet.FuncReadDataIoDictionary(selectEntity, dataStructuresObject.TablePS)
                    If dictionaryPs.ContainsKey("LABEL") = True Then
                        deskObject = dictionaryPs.Item("LABEL")
                    End If

                    Dim typeGeometry As StructureElement.typeGeometryObject = dataStructuresObject.TypeGeometry
                    Dim objectGeometry As Geometries.Geometry = Nothing
                    If typeGeometry = StructureElement.typeGeometryObject.polygon Then
                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Area > 0 Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(deskObject) = False Then
                                        result.Add(deskObject, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(deskObject)
                                        listObject.Add(polygon)
                                        result.Item(deskObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function
    '==========================================================================================================================================================
    'фильтрация объектов на основе таблицы PS
    Public Shared Function selectCadastralObject(ByRef Db As Database, Optional tablePS As String = "") As Dictionary(Of String, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of String, List(Of Geometries.Geometry)) = New Dictionary(Of String, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncNetTopologySuite = New FuncNetTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim selectEntity As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    Dim dataStructuresObject As StructureElement = Nothing
                    Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                    If boolRecData = True Then
                        Dim tablePSData As String = dataStructuresObject.TablePS
                        Dim boolFindObject As Boolean = False
                        If tablePS.Trim.Length > 0 Then
                            If tablePS Like tablePSData Then
                                boolFindObject = True
                            End If
                        Else
                            boolFindObject = True
                        End If
                        If tablePSData Like "CADASTRAL_OSTROV" Then
                            boolFindObject = True
                        End If
                        If boolFindObject = True Then
                            If TypeOf selectEntity Is Polyline Then
                                Dim polyline As Polyline = selectEntity
                                If polyline.Area > 0 And dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                    Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                    If IsNothing(polygon) = False Then
                                        polygon.UserData = dataStructuresObject
                                        If result.ContainsKey(tablePSData) = False Then
                                            result.Add(tablePSData, New List(Of Geometries.Geometry) From {polygon})
                                        Else
                                            Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                            listObject.Add(polygon)
                                            result.Item(tablePSData) = listObject
                                        End If
                                    End If
                                ElseIf dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polyline Then
                                    If TypeOf selectEntity Is Polyline Then
                                        If polyline.Length > 0 Then
                                            Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                            If IsNothing(line) = False Then
                                                line.UserData = dataStructuresObject
                                                If result.ContainsKey(tablePSData) = False Then
                                                    result.Add(tablePSData, New List(Of Geometries.Geometry) From {line})
                                                Else
                                                    Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                                    listObject.Add(line)
                                                    result.Item(tablePSData) = listObject
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function
    '==========================================================================================================================================================
    'фильтрация объектов по дескриптору объектов (deskription, geometry) 2 вырианта или основной объект или внутренний островок
    Public Shared Function selectCadastralObjectByDesk(ByRef Db As Database, ByVal DescriptorObject As String) As Dictionary(Of String, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of String, List(Of Geometries.Geometry)) = New Dictionary(Of String, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncNetTopologySuite = New FuncNetTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                If id.IsNull = False And id.IsErased = False Then
                    Dim selectEntity As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    Dim dataStructuresObject As StructureElement = Nothing
                    Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                    If boolRecData = True Then
                        Dim deskFindObject As String = dataStructuresObject.Description
                        Dim tablePSData As String = dataStructuresObject.TablePS
                        Dim boolFindRingObject As Boolean = False
                        If tablePSData Like "CADASTRAL_OSTROV" Then
                            boolFindRingObject = True
                        End If

                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Area > 0 And dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(tablePSData) = False Then
                                        result.Add(tablePSData, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                        listObject.Add(polygon)
                                        result.Item(tablePSData) = listObject
                                    End If
                                End If
                            ElseIf dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polyline Then
                                If TypeOf selectEntity Is Polyline Then
                                    If polyline.Length > 0 Then
                                        Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                        If IsNothing(line) = False Then
                                            line.UserData = dataStructuresObject
                                            If result.ContainsKey(tablePSData) = False Then
                                                result.Add(tablePSData, New List(Of Geometries.Geometry) From {line})
                                            Else
                                                Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                                listObject.Add(line)
                                                result.Item(tablePSData) = listObject
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'площадные объекты
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'фильтрация площадных объектов (новый вариант)
    Public Shared Function getCadastralPolygon(ByRef Db As Database, Optional deskCadastralObject As String = "") As Dictionary(Of String, List(Of Geometries.Geometry))
        ' словарь с участками
        Dim result As New Dictionary(Of String, List(Of Geometries.Geometry))
        ' список островков
        Dim listOstrov As New List(Of Geometries.Polygon)

        If Db Is Nothing Then Return result

        Dim classTopology As New FuncNetTopologySuite
        Dim filter As String = If(deskCadastralObject, "").Trim()

        Try
            Dim tr As Transaction = Db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = TryCast(tr.GetObject(Db.CurrentSpaceId, OpenMode.ForRead, False), BlockTableRecord)
                If btr Is Nothing Then Return result

                For Each id As ObjectId In btr
                    If id.IsNull OrElse id.IsErased Then Continue For

                    Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If acEnt Is Nothing Then Continue For

                    ' читаем пользовательские свойства
                    Dim dataCadastralObject As StructureElement = Nothing
                    If Not FuncXRecords.getXRecords(acEnt, dataCadastralObject) Then Continue For
                    If dataCadastralObject Is Nothing Then Continue For

                    ' объект не площадной — пропускаем
                    If dataCadastralObject.TypeGeometry <> CadastralTools.StructureElement.typeGeometryObject.polygon Then Continue For

                    Dim deskObject As String = If(dataCadastralObject.Description, "").Trim()

                    ' островки обрабатываем отдельно
                    If deskObject Like "Внутренний островок*" Then
                        Dim polyOstrov As Polyline = TryCast(acEnt, Polyline)
                        If polyOstrov IsNot Nothing AndAlso polyOstrov.NumberOfVertices > 2 Then
                            Dim polygonOstrov As NetTopologySuite.Geometries.Polygon = classTopology.CreatePolygon(polyOstrov)
                            polygonOstrov.UserData = dataCadastralObject
                            listOstrov.Add(polygonOstrov)
                        End If
                        Continue For
                    End If

                    ' фильтр по типу
                    If filter.Length > 0 Then
                        If Not (deskObject Like filter) Then Continue For
                    End If

                    Dim poly As Polyline = TryCast(acEnt, Polyline)
                    If poly Is Nothing Then Continue For
                    If poly.Area <= 0 Then Continue For

                    Dim polygon As NetTopologySuite.Geometries.Polygon = classTopology.CreatePolygon(poly)
                    polygon.UserData = dataCadastralObject

                    Dim list As List(Of Geometries.Geometry) = Nothing
                    If Not result.TryGetValue(deskObject, list) Then
                        list = New List(Of Geometries.Geometry)
                        result.Add(deskObject, list)
                    End If
                    list.Add(polygon)
                Next

                If result.Count = 0 Then Return result

                Dim classNetTopo As New CadastralTools.FuncNetTopologySuite

                '--------------------------------------------------------------------------------
                ' 1) Вложенные полигоны одного типа → дырки
                '--------------------------------------------------------------------------------
                For Each kvp As KeyValuePair(Of String, List(Of Geometries.Geometry)) In result
                    Dim listSource As List(Of Geometries.Geometry) = kvp.Value
                    If listSource Is Nothing OrElse listSource.Count < 2 Then Continue For

                    For j As Integer = 0 To listSource.Count - 1
                        Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = TryCast(listSource.Item(j), NetTopologySuite.Geometries.Polygon)
                        If sourcePolygon Is Nothing Then Continue For

                        Dim dataSourcePolygon As StructureElement = TryCast(sourcePolygon.UserData, StructureElement)
                        If dataSourcePolygon Is Nothing Then Continue For

                        Dim areaSourcePolygon As Double = sourcePolygon.Area
                        If areaSourcePolygon <= 0 Then Continue For

                        Dim sourceChanged As Boolean = False

                        For k As Integer = 0 To listSource.Count - 1
                            If k = j Then Continue For   ' не сравниваем сам с собой

                            Dim destPolygon As NetTopologySuite.Geometries.Polygon = TryCast(listSource.Item(k), NetTopologySuite.Geometries.Polygon)
                            If destPolygon Is Nothing Then Continue For

                            Dim dataDestPolygon As StructureElement = TryCast(destPolygon.UserData, StructureElement)
                            If dataDestPolygon Is Nothing Then Continue For

                            ' тип объектов должен совпадать
                            If Not String.Equals(dataDestPolygon.NameTypeObject, dataSourcePolygon.NameTypeObject, StringComparison.Ordinal) Then Continue For

                            Dim areaDestPolygon As Double = destPolygon.Area
                            If areaDestPolygon <= 0 OrElse areaDestPolygon >= areaSourcePolygon Then Continue For

                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                Dim updated As NetTopologySuite.Geometries.Geometry = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon, dataDestPolygon)
                                If updated IsNot Nothing Then
                                    sourcePolygon = TryCast(updated, NetTopologySuite.Geometries.Polygon)
                                    If sourcePolygon IsNot Nothing Then
                                        sourceChanged = True
                                    End If
                                End If
                            End If
                        Next k

                        If sourceChanged Then
                            listSource.Item(j) = sourcePolygon
                        End If
                    Next j
                Next

                '--------------------------------------------------------------------------------
                ' 2) Островки → дырки в родительских полигонах
                '--------------------------------------------------------------------------------
                If listOstrov.Count = 0 Then Return result

                For Each kvp As KeyValuePair(Of String, List(Of Geometries.Geometry)) In result
                    Dim listSource As List(Of Geometries.Geometry) = kvp.Value
                    If listSource Is Nothing OrElse listSource.Count = 0 Then Continue For

                    For j As Integer = 0 To listSource.Count - 1
                        Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = TryCast(listSource.Item(j), NetTopologySuite.Geometries.Polygon)
                        If sourcePolygon Is Nothing Then Continue For

                        Dim userPolygon As StructureElement = TryCast(sourcePolygon.UserData, StructureElement)
                        If userPolygon Is Nothing Then Continue For

                        Dim areaSourcePolygon As Double = sourcePolygon.Area
                        If areaSourcePolygon <= 0 Then Continue For

                        Dim sourceChanged As Boolean = False

                        For k As Integer = 0 To listOstrov.Count - 1
                            Dim destPolygon As NetTopologySuite.Geometries.Polygon = listOstrov.Item(k)
                            If destPolygon Is Nothing Then Continue For

                            Dim userRing As StructureElement = TryCast(destPolygon.UserData, StructureElement)
                            If userRing Is Nothing Then Continue For

                            Dim parentType As String = CadastralTools.JsonValidator.GetValueFromJson(userRing.KeyParameter, "ParentNameType")
                            If String.IsNullOrEmpty(parentType) Then Continue For

                            ' сравнение без учёта регистра, точное
                            If Not String.Equals(parentType.Trim(), userPolygon.NameTypeObject, StringComparison.OrdinalIgnoreCase) Then Continue For

                            Dim areaDestPolygon As Double = destPolygon.Area
                            If areaDestPolygon <= 0 OrElse areaDestPolygon >= areaSourcePolygon Then Continue For

                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                Dim updated As NetTopologySuite.Geometries.Geometry = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon, userRing)
                                If updated IsNot Nothing Then
                                    sourcePolygon = TryCast(updated, NetTopologySuite.Geometries.Polygon)
                                    If sourcePolygon IsNot Nothing Then
                                        sourceChanged = True
                                    End If
                                End If
                            End If
                        Next k

                        If sourceChanged Then
                            listSource.Item(j) = sourcePolygon
                        End If
                    Next j
                Next
            End Using
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[getCadastralPolygon] " & ex.ToString())
        End Try
        Return result
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'линейные объекты
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'фильтрация линейных объектов (новый вариант)
    Public Shared Function getCadastralPolyline(ByRef Db As Database, Optional deskCadastralObject As String = "") As Dictionary(Of String, List(Of Geometries.Geometry))
        ' словарь с участками
        Dim result As New Dictionary(Of String, List(Of Geometries.Geometry))

        If Db Is Nothing Then Return result

        Dim classTopology As New FuncNetTopologySuite
        Dim filter As String = If(deskCadastralObject, "").Trim()

        Try
            Dim tr As Transaction = Db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = TryCast(tr.GetObject(Db.CurrentSpaceId, OpenMode.ForRead, False), BlockTableRecord)
                If btr Is Nothing Then Return result

                For Each id As ObjectId In btr
                    If id.IsNull OrElse id.IsErased Then Continue For

                    Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                    If acEnt Is Nothing Then Continue For

                    ' читаем пользовательские свойства
                    Dim dataCadastralObject As StructureElement = Nothing
                    If Not FuncXRecords.getXRecords(acEnt, dataCadastralObject) Then Continue For
                    If dataCadastralObject Is Nothing Then Continue For

                    ' объект не линейный — пропускаем
                    If dataCadastralObject.TypeGeometry <> CadastralTools.StructureElement.typeGeometryObject.polyline Then Continue For

                    Dim deskObject As String = If(dataCadastralObject.Description, "").Trim()

                    ' фильтр по типу
                    If filter.Length > 0 Then
                        If Not (deskObject Like filter) Then Continue For
                    End If

                    Dim poly As Polyline = TryCast(acEnt, Polyline)
                    If poly Is Nothing Then Continue For
                    If poly.Area <= 0 Then Continue For

                    Dim polygon As NetTopologySuite.Geometries.Polygon = classTopology.CreatePolygon(poly)
                    polygon.UserData = dataCadastralObject

                    Dim list As List(Of Geometries.Geometry) = Nothing
                    If Not result.TryGetValue(deskObject, list) Then
                        list = New List(Of Geometries.Geometry)
                        result.Add(deskObject, list)
                    End If
                    list.Add(polygon)
                Next
            End Using
        Catch ex As System.Exception
            System.Diagnostics.Trace.WriteLine("[getCadastralPolyline] " & ex.ToString())
        End Try
        Return result
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'точечные объекты
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает блоки из чертежа по имени и видимости слоя
    Public Shared Function GetBlockReferencesOnVisibleLayers(ByVal blockName As String) As List(Of BlockReference)
        Dim result As New List(Of BlockReference)()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return result

        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Фильтр по типу объекта и имени блока (имя слоя в фильтр не заложить — проверяем вручную)
        Dim filter As New SelectionFilter(New TypedValue() {
            New TypedValue(CInt(DxfCode.Start), "INSERT"),
            New TypedValue(CInt(DxfCode.BlockName), blockName)
        })

        Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        If selRes.Status <> PromptStatus.OK Then Return result

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Кэш состояний слоёв, чтобы не открывать один и тот же слой многократно
            Dim layerVisible As New Dictionary(Of ObjectId, Boolean)()

            For Each id As ObjectId In selRes.Value.GetObjectIds()
                Dim br As BlockReference = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockReference)
                If br Is Nothing Then Continue For

                Dim layerId As ObjectId = br.LayerId
                Dim isVisible As Boolean

                If Not layerVisible.TryGetValue(layerId, isVisible) Then
                    Dim ltr As LayerTableRecord =
                        TryCast(tr.GetObject(layerId, OpenMode.ForRead), LayerTableRecord)
                    isVisible = (ltr IsNot Nothing) AndAlso (Not ltr.IsOff) AndAlso (Not ltr.IsFrozen)
                    layerVisible(layerId) = isVisible
                End If

                If isVisible Then
                    result.Add(br)
                End If
            Next

            tr.Commit()
        End Using

        Return result
    End Function
    '===================================================================================================================
    'работа с точками
    Public Shared Function getPointByListBlockReference(ByRef db As Database, ByVal blockName As String) As Dictionary(Of String, NetTopologySuite.Geometries.Point)
        Dim result As New Dictionary(Of String, NetTopologySuite.Geometries.Point)
        Dim doc As Document = Application.DocumentManager.GetDocument(db)
        If doc Is Nothing Then Return result
        Dim ed As Editor = doc.Editor
        ' Фильтр по типу объекта и имени блока (имя слоя в фильтр не заложить — проверяем вручную)
        Dim filter As New SelectionFilter(New TypedValue() {
            New TypedValue(CInt(DxfCode.Start), "INSERT"),
            New TypedValue(CInt(DxfCode.BlockName), blockName)
        })

        Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        If selRes.Status <> PromptStatus.OK Then Return result

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Кэш состояний слоёв, чтобы не открывать один и тот же слой многократно
            Dim layerVisible As New Dictionary(Of ObjectId, Boolean)()

            For Each id As ObjectId In selRes.Value.GetObjectIds()
                Dim br As BlockReference = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockReference)
                If br Is Nothing Then Continue For

                Dim layerId As ObjectId = br.LayerId
                Dim isVisible As Boolean

                If Not layerVisible.TryGetValue(layerId, isVisible) Then
                    Dim ltr As LayerTableRecord =
                        TryCast(tr.GetObject(layerId, OpenMode.ForRead), LayerTableRecord)
                    isVisible = (ltr IsNot Nothing) AndAlso (Not ltr.IsOff) AndAlso (Not ltr.IsFrozen)
                    layerVisible(layerId) = isVisible
                End If

                If isVisible Then
                    Dim dataPoint As StructureElement = Nothing
                    Dim boolReadData As Boolean = FuncXRecords.getXRecords(br, dataPoint)
                    If IsNothing(dataPoint) = False Then
                        Dim userPoint As NetTopologySuite.Geometries.Point = New Geometries.Point(br.Position.X, br.Position.Y, br.Position.Z)
                        Dim dictAttrib As Dictionary(Of String, String) = AttributeCadastralBlock.readArrtibutes(db, br)
                        dictAttrib.Add("IDRNG", dataPoint.IdStructure)
                        userPoint.UserData = dictAttrib
                        result.Add(br.Handle.ToString, userPoint)
                    End If
                End If
            Next
            tr.Commit()
        End Using
        Return result
    End Function
End Class
