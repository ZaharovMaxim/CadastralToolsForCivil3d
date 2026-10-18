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
Public Class CadastralBufferZones
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'построение буферных зон вдоль выбранного контура (в т.ч. многоконтурного объекта)
    <CommandMethod("CreateBufferZones")>
    Public Sub CreateBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ШАГ 1. выбираем примитив (в том числе из внешней ссылки)
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите контур (полилинию): ")
        If Not (rs.Status = PromptStatus.OK) Then
            Exit Sub
        End If
        Try
            'ШАГИ 8-9. запрашиваем ширину зоны влево и вправо
            Dim leftWidth As Double = AskBufferWidth(ed, "влево")
            If leftWidth < 0 Then Exit Sub
            Dim rightWidth As Double = AskBufferWidth(ed, "вправо")
            If rightWidth < 0 Then Exit Sub
            If leftWidth = 0 And rightWidth = 0 Then
                ed.WriteMessage(vbLf & "Обе ширины равны нулю — зона не построена.")
                Exit Sub
            End If

            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                'трансформация контейнеров: для примитива из внешней ссылки/блока
                'переводит его координаты в координаты текущего чертежа
                Dim xform As Matrix3d = Matrix3d.Identity
                Dim listContainers As New List(Of ObjectId)(rs.GetContainers())
                listContainers.Reverse()
                For Each idContainer As ObjectId In listContainers
                    Dim brefContainer As BlockReference = TryCast(tr.GetObject(idContainer, OpenMode.ForRead), BlockReference)
                    If brefContainer IsNot Nothing Then
                        xform = xform * brefContainer.BlockTransform
                    End If
                Next
                Dim boolTransform As Boolean = (listContainers.Count > 0)

                Dim sourceEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(rs.ObjectId, OpenMode.ForRead)
                If Not (TypeOf sourceEnt Is Polyline) Then
                    ed.WriteMessage(vbLf & "Выбранный объект не является полилинией.")
                    tr.Commit()
                    Exit Sub
                End If
                Dim sourcePoly As Polyline = DirectCast(sourceEnt, Polyline)
                Dim boolClosed As Boolean = sourcePoly.Closed

                'список исходных контуров, вдоль которых строятся зоны
                Dim listSourceGeometry As New List(Of Geometry)
                Dim listIsClosed As New List(Of Boolean)

                'ШАГ 2. читаем XData (getXRecords -> StructureElement)
                Dim dataXRecord As CadastralTools.StructureElement = Nothing
                Dim boolData As Boolean = FuncXRecords.getXRecords(sourceEnt, dataXRecord)

                If boolData = True And dataXRecord IsNot Nothing Then
                    'ШАГ 4. проверяем выбранный элемент на многоконтурность
                    Dim conditionalNumber As String = CadastralTools.JsonValidator.GetValueFromJson(dataXRecord.KeyParameter, "ConditionalNumber")
                    Dim numberCounter As Integer = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalNumber)
                    If numberCounter > 0 Then
                        'ШАГ 6. читаем все элементы чертежа нужного типа геометрии
                        Dim dictCadastralObject As New Dictionary(Of String, List(Of Geometry))
                        If dataXRecord.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon Then
                            dictCadastralObject = SelectCadastralObject.getCadastralPolygon(db, dataXRecord.Description)
                        Else
                            dictCadastralObject = SelectCadastralObject.getCadastralPolyline(db, dataXRecord.Description)
                        End If
                        If dictCadastralObject.Count > 0 Then
                            Dim listAllCounters As New List(Of Geometry)
                            For Each kvp As KeyValuePair(Of String, List(Of Geometry)) In dictCadastralObject
                                listAllCounters.AddRange(kvp.Value)
                            Next
                            'ШАГ 7. находим все контура выбранного объекта
                            '(номер объекта — условный номер без номера контура)
                            Dim numberObject As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalNumber)
                            Dim dictCounters As Dictionary(Of Integer, Geometry) = Nothing
                            If dataXRecord.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon Then
                                dictCounters = CadastralTools.CadastralPolygon.getCountersByNumber(numberObject, listAllCounters)
                            Else
                                dictCounters = CadastralTools.CadastralLine.getCountersByNumber(numberObject, listAllCounters)
                            End If
                            If Not IsNothing(dictCounters) Then
                                For Each kvpCounter As KeyValuePair(Of Integer, Geometry) In dictCounters
                                    If kvpCounter.Value Is Nothing OrElse kvpCounter.Value.IsEmpty Then Continue For
                                    listSourceGeometry.Add(kvpCounter.Value)
                                    listIsClosed.Add(TypeOf kvpCounter.Value Is Polygon)
                                Next
                            End If
                        End If
                    End If
                End If

                'ШАГИ 3, 5. данных нет или контур единственный —
                'зона строится только по выбранному примитиву
                If listSourceGeometry.Count = 0 Then
                    Dim sourceGeometry As Geometry = Nothing
                    If boolClosed Then
                        sourceGeometry = CreatePolygonFromPolyline(sourcePoly, xform, boolTransform)
                    Else
                        sourceGeometry = CreateLineStringFromPolyline(sourcePoly, xform, boolTransform)
                    End If
                    If sourceGeometry IsNot Nothing AndAlso Not sourceGeometry.IsEmpty Then
                        listSourceGeometry.Add(sourceGeometry)
                        listIsClosed.Add(boolClosed)
                    End If
                End If

                If listSourceGeometry.Count = 0 Then
                    ed.WriteMessage(vbLf & "Не удалось получить геометрию контура.")
                    tr.Commit()
                    Exit Sub
                End If

                'построение буферных зон (FuncNetTopologySuite / BufferZoneCreator)
                Dim listAllZonePolygons As New List(Of Polygon)
                For i As Integer = 0 To listSourceGeometry.Count - 1
                    Dim zoneGeometry As Geometry = BuildZoneGeometry(listSourceGeometry.Item(i), listIsClosed.Item(i), leftWidth, rightWidth)
                    If zoneGeometry Is Nothing OrElse zoneGeometry.IsEmpty Then Continue For
                    ExtractPolygonsFromGeometry(zoneGeometry, listAllZonePolygons)
                Next

                If listAllZonePolygons.Count = 0 Then
                    ed.WriteMessage(vbLf & "Буферные зоны не построены.")
                    tr.Commit()
                    Exit Sub
                End If

                'перекрывающиеся зоны объединяем
                Dim listUnionPolygons As List(Of Polygon) = netTopo.UnionSeparate(listAllZonePolygons)
                If listUnionPolygons Is Nothing OrElse listUnionPolygons.Count = 0 Then
                    ed.WriteMessage(vbLf & "Не удалось объединить буферные зоны.")
                    tr.Commit()
                    Exit Sub
                End If

                'ШАГ 9. результат выводим в текущий чертеж
                '(createPolylineByPolygon доработана — строит островки)
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                For i As Integer = 0 To listUnionPolygons.Count - 1
                    Dim zonePoly As Polygon = listUnionPolygons.Item(i)
                    Dim rezPoly As Polyline = classDrawObject.createPolylineByPolygon(zonePoly)
                    If Not IsNothing(rezPoly) Then
                        rezPoly.ColorIndex = sourcePoly.ColorIndex
                        rezPoly.Linetype = sourcePoly.Linetype
                        rezPoly.LinetypeScale = sourcePoly.LinetypeScale
                        rezPoly.LineWeight = sourcePoly.LineWeight
                    End If
                Next

                tr.Commit()
                ed.WriteMessage(vbLf & "Построено полигонов зон: " & listUnionPolygons.Count.ToString)
            End Using
        Catch ex As System.Exception
            ed.WriteMessage(vbLf & "Ошибка: " & ex.Message)
        End Try
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'запрос ширины буферной зоны (возвращает -1 при отмене)
    Private Function AskBufferWidth(ByVal ed As Editor, ByVal direction As String) As Double
        Dim pdo As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ширина зоны " & direction & " <0.0>: ")
        pdo.AllowNegative = False
        pdo.AllowZero = True
        Dim pdr As PromptDoubleResult = ed.GetDouble(pdo)
        If pdr.Status <> PromptStatus.OK Then
            Return -1.0
        End If
        Return pdr.Value
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'создание NTS-полигона из полилинии с учетом трансформации контейнеров (xref/блоки)
    Private Function CreatePolygonFromPolyline(ByVal pl As Polyline, ByVal xform As Matrix3d, ByVal boolTransform As Boolean) As Polygon
        If pl.NumberOfVertices < 3 Then
            Return Nothing
        End If
        Dim listCoord As New List(Of Coordinate)
        For i As Integer = 0 To pl.NumberOfVertices - 1
            Dim pt As Point3d = pl.GetPoint3dAt(i)
            If boolTransform = True Then
                pt = pt.TransformBy(xform)
            End If
            listCoord.Add(New Coordinate(pt.X, pt.Y))
        Next
        If listCoord.Count > 2 Then
            If Not listCoord(0).Equals2D(listCoord(listCoord.Count - 1)) Then
                listCoord.Add(New Coordinate(listCoord(0)))
            End If
        End If
        Dim gf As New GeometryFactory()
        Return gf.CreatePolygon(gf.CreateLinearRing(listCoord.ToArray()))
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'создание NTS-линии из полилинии с учетом трансформации контейнеров (xref/блоки)
    Private Function CreateLineStringFromPolyline(ByVal pl As Polyline, ByVal xform As Matrix3d, ByVal boolTransform As Boolean) As LineString
        If pl.NumberOfVertices < 2 Then
            Return Nothing
        End If
        Dim listCoord As New List(Of Coordinate)
        For i As Integer = 0 To pl.NumberOfVertices - 1
            Dim pt As Point3d = pl.GetPoint3dAt(i)
            If boolTransform = True Then
                pt = pt.TransformBy(xform)
            End If
            listCoord.Add(New Coordinate(pt.X, pt.Y))
        Next
        Dim gf As New GeometryFactory()
        Return gf.CreateLineString(listCoord.ToArray())
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'построение буферной зоны вдоль контура: влево — leftWidth, вправо — rightWidth
    'замкнутый контур: кольцевая зона Buffer(+влево) \ Buffer(-вправо)
    'разомкнутый контур: односторонние буферы (+влево, -вправо)
    Private Function BuildZoneGeometry(ByVal sourceGeometry As Geometry, ByVal boolClosed As Boolean,
                                       ByVal leftWidth As Double, ByVal rightWidth As Double) As Geometry
        Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
        If boolClosed = True Then
            Dim zoneOuter As Geometry = Nothing
            If leftWidth > 0 Then
                zoneOuter = netBuffer.CreateCustomBuffer(sourceGeometry, leftWidth, 8, EndCapStyle.Round)
            End If
            If rightWidth > 0 Then
                Dim zoneInner As Geometry = netBuffer.CreateCustomBuffer(sourceGeometry, -rightWidth, 8, EndCapStyle.Round)
                If zoneOuter Is Nothing Then
                    zoneOuter = zoneInner
                ElseIf zoneInner IsNot Nothing AndAlso Not zoneInner.IsEmpty Then
                    zoneOuter = zoneOuter.Difference(zoneInner)
                End If
            End If
            Return zoneOuter
        Else
            Dim zoneLeft As Geometry = Nothing
            Dim zoneRight As Geometry = Nothing
            If leftWidth > 0 Then
                zoneLeft = SingleSideBuffer(sourceGeometry, leftWidth)
            End If
            If rightWidth > 0 Then
                zoneRight = SingleSideBuffer(sourceGeometry, -rightWidth)
            End If
            If zoneLeft IsNot Nothing And zoneRight IsNot Nothing Then
                Return zoneLeft.Union(zoneRight)
            End If
            Return If(zoneLeft IsNot Nothing, zoneLeft, zoneRight)
        End If
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'односторонний буфер: + расстояние — влево от линии, - — вправо
    Private Function SingleSideBuffer(ByVal sourceGeometry As Geometry, ByVal distance As Double) As Geometry
        Dim bp As BufferParameters = New BufferParameters With {
            .QuadrantSegments = 8,
            .EndCapStyle = EndCapStyle.Flat,
            .IsSingleSided = True
        }
        Dim op As BufferOp = New BufferOp(sourceGeometry, bp)
        Return op.GetResultGeometry(distance)
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'извлечение всех полигонов из геометрии (Polygon / MultiPolygon / GeometryCollection)
    Private Sub ExtractPolygonsFromGeometry(ByVal geom As Geometry, ByRef listPolygons As List(Of Polygon))
        If geom Is Nothing OrElse geom.IsEmpty Then
            Return
        End If
        If TypeOf geom Is Polygon Then
            listPolygons.Add(DirectCast(geom, Polygon))
        ElseIf TypeOf geom Is MultiPolygon Then
            Dim mp As MultiPolygon = DirectCast(geom, MultiPolygon)
            For i As Integer = 0 To mp.NumGeometries - 1
                ExtractPolygonsFromGeometry(mp.GetGeometryN(i), listPolygons)
            Next
        ElseIf TypeOf geom Is GeometryCollection Then
            Dim gc As GeometryCollection = DirectCast(geom, GeometryCollection)
            For i As Integer = 0 To gc.NumGeometries - 1
                ExtractPolygonsFromGeometry(gc.GetGeometryN(i), listPolygons)
            Next
        End If
    End Sub
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'выбор вложенной полилинии (в т.ч. из внешней ссылки) с трансформацией контейнеров.
    'Возвращает Nothing при отмене/если объект не полилиния.
    'xform/boolTransform — трансформация координат в WCS текущего чертежа
    Private Function SelectNestedPolyline(ByVal ed As Editor, ByVal tr As Transaction, ByVal message As String,
                                          ByRef xform As Matrix3d, ByRef boolTransform As Boolean) As Polyline
        xform = Matrix3d.Identity
        boolTransform = False
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(message)
        If rs.Status <> PromptStatus.OK Then
            Return Nothing
        End If
        Dim listContainers As New List(Of ObjectId)(rs.GetContainers())
        listContainers.Reverse()
        For Each idContainer As ObjectId In listContainers
            Dim brefContainer As BlockReference = TryCast(tr.GetObject(idContainer, OpenMode.ForRead), BlockReference)
            If brefContainer IsNot Nothing Then
                xform = xform * brefContainer.BlockTransform
            End If
        Next
        boolTransform = (listContainers.Count > 0)
        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(rs.ObjectId, OpenMode.ForRead)
        If TypeOf ent Is Polyline Then
            Return DirectCast(ent, Polyline)
        End If
        Return Nothing
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'проверка: полилиния редактируемая (не из внешней ссылки)
    Private Function IsPolylineEditable(ByVal tr As Transaction, ByVal pl As Polyline) As Boolean
        Try
            If pl.OwnerId.IsNull Then
                Return False
            End If
            Dim owner As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(pl.OwnerId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = TryCast(owner, BlockTableRecord)
            If btr IsNot Nothing AndAlso btr.IsFromExternalReference Then
                Return False
            End If
            Return True
        Catch
            Return False
        End Try
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'проверка: объект принадлежит внешней ссылке
    Private Function IsIdFromXRef(ByVal tr As Transaction, ByVal id As ObjectId) As Boolean
        Try
            Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(id, OpenMode.ForRead)
            If obj.OwnerId.IsNull Then
                Return False
            End If
            Dim owner As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(obj.OwnerId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = TryCast(owner, BlockTableRecord)
            Return (btr IsNot Nothing AndAlso btr.IsFromExternalReference)
        Catch
            Return False
        End Try
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'клонирование StructureElement
    Private Function CloneStructureElement(ByVal source As CadastralTools.StructureElement) As CadastralTools.StructureElement
        Dim clone As New CadastralTools.StructureElement()
        clone.Label = source.Label
        clone.NameClassObject = source.NameClassObject
        clone.TypeGeometry = source.TypeGeometry
        clone.NameTypeObject = source.NameTypeObject
        clone.TableObject = source.TableObject
        clone.Description = source.Description
        clone.KeyParameter = source.KeyParameter
        clone.TablePS = source.TablePS
        clone.IdStructure = source.IdStructure
        clone.IdElement = source.IdElement
        clone.Note = source.Note
        clone.IdObject = source.IdObject
        Return clone
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'копирует XData (StructureElement) с образца на полилинию;
    'при заданном newConditionalNumber заменяет ConditionalNumber в KeyParameter
    Private Sub CopyXRecordsToPolyline(ByVal templateElement As CadastralTools.StructureElement,
                                       ByVal targetPoly As Polyline,
                                       Optional ByVal newConditionalNumber As String = "")
        If templateElement Is Nothing Then
            Return
        End If
        Dim newElement As CadastralTools.StructureElement = CloneStructureElement(templateElement)
        If newConditionalNumber.Length > 0 Then
            newElement.KeyParameter = CadastralTools.JsonValidator.SetJsonFieldValue(newElement.KeyParameter, "ConditionalNumber", newConditionalNumber)
        End If
        Dim entTarget As Autodesk.AutoCAD.DatabaseServices.Entity = targetPoly
        FuncXRecords.setXRecords(entTarget, newElement)
    End Sub


    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    'объединение буферных зон (StructureElements берется у первой выбранной полилинии)
    <CommandMethod("UnionBufferZones")>
    Public Sub UnionBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Try
                'первая полилиния — донор StructureElements и основа для обновления
                Dim xformFirst As Matrix3d = Matrix3d.Identity
                Dim boolTransformFirst As Boolean = False
                Dim firstPoly As Polyline = SelectNestedPolyline(ed, tr, vbLf & "Выберите первую замкнутую полилинию: ", xformFirst, boolTransformFirst)
                If firstPoly Is Nothing Then
                    tr.Commit() : Exit Sub
                End If
                'читаем StructureElements первой полилинии
                Dim dataFirst As CadastralTools.StructureElement = Nothing
                FuncXRecords.getXRecords(firstPoly, dataFirst)

                Dim firstPolygon As Polygon = CreatePolygonFromPolyline(firstPoly, xformFirst, boolTransformFirst)
                If firstPolygon Is Nothing Then
                    ed.WriteMessage(vbLf & "Первая полилиния не образует полигон.")
                    tr.Commit() : Exit Sub
                End If

                'выбираем остальные полилинии (Enter — завершить выбор)
                Dim listPolygons As New List(Of Polygon) From {firstPolygon}
                Dim listSelectedIds As New List(Of ObjectId)()
                Do
                    Dim xformNext As Matrix3d = Matrix3d.Identity
                    Dim boolTransformNext As Boolean = False
                    Dim nextPoly As Polyline = SelectNestedPolyline(ed, tr, vbLf & "Выберите следующую замкнутую полилинию (Enter — завершить): ", xformNext, boolTransformNext)
                    If nextPoly Is Nothing Then
                        Exit Do
                    End If
                    Dim nextPolygon As Polygon = CreatePolygonFromPolyline(nextPoly, xformNext, boolTransformNext)
                    If nextPolygon Is Nothing Then
                        Continue Do
                    End If
                    If netTopo.CanUnion(firstPolygon, nextPolygon) = False Then
                        MsgBox("Нельзя объединить полигоны!!! Выберите другой контур или завершите выбор (Enter).")
                        Continue Do
                    End If
                    listPolygons.Add(nextPolygon)
                    listSelectedIds.Add(nextPoly.ObjectId)
                Loop

                If listPolygons.Count < 2 Then
                    ed.WriteMessage(vbLf & "Выбран только один полигон — объединять нечего.")
                    tr.Commit() : Exit Sub
                End If

                'объединяем (FuncNetTopologySuite)
                Dim unionGeometry As Geometry = netTopo.UnionMultiplePolygons(listPolygons)
                If unionGeometry Is Nothing OrElse unionGeometry.IsEmpty Then
                    ed.WriteMessage(vbLf & "Не удалось объединить полигоны.")
                    tr.Commit() : Exit Sub
                End If
                Dim listResultPolygons As New List(Of Polygon)
                ExtractPolygonsFromGeometry(unionGeometry, listResultPolygons)
                If listResultPolygons.Count = 0 Then
                    tr.Commit() : Exit Sub
                End If

                Dim boolEditFirst As Boolean = IsPolylineEditable(tr, firstPoly)
                If boolEditFirst = True Then
                    'полигон, обновляющий первую полилинию — содержащий ее центроид
                    Dim pointFirst As NetTopologySuite.Geometries.Point = firstPolygon.Centroid
                    Dim indexBase As Integer = 0
                    For i As Integer = 0 To listResultPolygons.Count - 1
                        If listResultPolygons.Item(i).Contains(pointFirst) Then
                            indexBase = i
                            Exit For
                        End If
                    Next
                    'обновляем первую полилинию — StructureElements сохраняются
                    '(островки достраиваются отдельными полилиниями)
                    classDrawObject.createPolylineByPolygon(listResultPolygons.Item(indexBase), firstPoly)
                    listResultPolygons.RemoveAt(indexBase)
                End If

                'остальные полигоны — новые полилинии с XData первой
                For i As Integer = 0 To listResultPolygons.Count - 1
                    Dim rezPoly As Polyline = classDrawObject.createPolylineByPolygon(listResultPolygons.Item(i))
                    If Not IsNothing(rezPoly) Then
                        rezPoly.ColorIndex = firstPoly.ColorIndex
                        rezPoly.Linetype = firstPoly.Linetype
                        rezPoly.LinetypeScale = firstPoly.LinetypeScale
                        rezPoly.LineWeight = firstPoly.LineWeight
                        CopyXRecordsToPolyline(dataFirst, rezPoly)
                    End If
                Next

                'удаляем объединенные полилинии (кроме первой)
                If boolEditFirst = True Then
                    For Each idSel As ObjectId In listSelectedIds
                        If idSel = firstPoly.ObjectId Then
                            Continue For
                        End If
                        Try
                            If IsIdFromXRef(tr, idSel) = True Then
                                Continue For
                            End If
                            Dim entSel As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idSel, OpenMode.ForWrite)
                            entSel.Erase()
                        Catch
                        End Try
                    Next
                End If

                tr.Commit()
                ed.WriteMessage(vbLf & "Объединение завершено. Полигонов в результате: " & (listPolygons.Count - listSelectedIds.Count + listResultPolygons.Count + If(boolEditFirst, 1, 0)).ToString)
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Ошибка: " & ex.Message)
            End Try
        End Using
    End Sub
    '////////////////////////////////////////////////////////////////////////////////////////////////////////
    'обрезка зоны по полосе отвода (разность полигонов)
    <CommandMethod("DifferenceBufferZones")>
    Public Sub DifferenceBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Try
                'режущий полигон (в т.ч. из внешней ссылки)
                Dim xformCutter As Matrix3d = Matrix3d.Identity
                Dim boolTransformCutter As Boolean = False
                Dim cutterPoly As Polyline = SelectNestedPolyline(ed, tr, vbLf & "Выберите полилинию по которой будет происходить обрезка: ", xformCutter, boolTransformCutter)
                If cutterPoly Is Nothing Then
                    tr.Commit() : Exit Sub
                End If
                Dim cutterGeometry As Polygon = CreatePolygonFromPolyline(cutterPoly, xformCutter, boolTransformCutter)
                If cutterGeometry Is Nothing Then
                    ed.WriteMessage(vbLf & "Режущая полилиния не образует полигон.")
                    tr.Commit() : Exit Sub
                End If

                'выбираем буферные зоны
                Dim pso As PromptSelectionOptions = New PromptSelectionOptions With {
                    .MessageForAdding = vbLf & "Выберите буферные зоны: "
                }
                Dim acSSPrompt As PromptSelectionResult = ed.GetSelection(pso)
                If acSSPrompt.Status <> PromptStatus.OK Then
                    tr.Commit() : Exit Sub
                End If
                For Each acSSObj As SelectedObject In acSSPrompt.Value
                    If acSSObj Is Nothing Then
                        Continue For
                    End If
                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                    If Not TypeOf acEnt Is Polyline Then
                        Continue For
                    End If
                    Dim zonePoly As Polyline = DirectCast(acEnt, Polyline)
                    If zonePoly.Closed = False Then
                        Continue For
                    End If
                    Dim zoneGeometry As Polygon = CreatePolygonFromPolyline(zonePoly, Matrix3d.Identity, False)
                    If zoneGeometry Is Nothing OrElse zoneGeometry.IsEmpty Then
                        Continue For
                    End If
                    If netTopo.DoPolygonsIntersect(cutterGeometry, zoneGeometry) = False Then
                        Continue For
                    End If
                    'части зоны вне режущего полигона (FuncNetTopologySuite)
                    Dim listParts As List(Of Polygon) = netTopo.GetPolygonsOutsidePoly1(cutterGeometry, zoneGeometry)
                    'XData зоны — копируем на новые полилинии
                    Dim dataZone As CadastralTools.StructureElement = Nothing
                    FuncXRecords.getXRecords(zonePoly, dataZone)
                    Dim boolEditable As Boolean = IsPolylineEditable(tr, zonePoly)
                    If listParts.Count = 0 Then
                        'зона полностью внутри полосы отвода — удаляем
                        If boolEditable = True Then
                            Try
                                Dim entErase As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(zonePoly.ObjectId, OpenMode.ForWrite)
                                entErase.Erase()
                            Catch
                            End Try
                        End If
                        Continue For
                    End If
                    If boolEditable = True Then
                        'обновляем исходную полилинию первой частью — XData сохраняются
                        classDrawObject.createPolylineByPolygon(listParts.Item(0), zonePoly)
                        listParts.RemoveAt(0)
                    End If
                    'остальные части — новые полилинии с XData зоны
                    For i As Integer = 0 To listParts.Count - 1
                        Dim rezPoly As Polyline = classDrawObject.createPolylineByPolygon(listParts.Item(i))
                        If Not IsNothing(rezPoly) Then
                            rezPoly.ColorIndex = zonePoly.ColorIndex
                            rezPoly.Linetype = zonePoly.Linetype
                            rezPoly.LinetypeScale = zonePoly.LinetypeScale
                            rezPoly.LineWeight = zonePoly.LineWeight
                            CopyXRecordsToPolyline(dataZone, rezPoly)
                        End If
                    Next
                Next
                tr.Commit()
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Ошибка: " & ex.Message)
            End Try
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'удалить лишние вершины (полилиния редактируется на месте, XData сохраняются)
    <CommandMethod("DeleteVertexZones")>
    Public Sub DeleteVertexZona()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        '==============================================================================
        'открываем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '=======================================================================================================
            'выбираем полилинию
            Dim acBFEnt As PromptEntityResult = ed.GetEntity(vbLf & "Выберите буферную зону: ")
            If acBFEnt.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acBFEnt.ObjectId, OpenMode.ForWrite)
            If Not TypeOf ent Is Polyline Then
                Exit Sub
            End If
            Dim PlineBF As Polyline = DirectCast(ent, Polyline)
            '========================================================================================================
            'выбираем начальную точку
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions(vbLf & "Выберите начальную точку: ")
            Dim pPtRes As PromptPointResult = ed.GetPoint(pPtOpts)
            If pPtRes.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim TStartPoint As Point3d = PlineBF.GetClosestPointTo(pPtRes.Value, False)
            '==========================================================================================================
            'выбираем конечную точку
            pPtOpts.Message = vbLf & "Выберите конечную точку: "
            pPtRes = ed.GetPoint(pPtOpts)
            If pPtRes.Status <> PromptStatus.OK Then
                Exit Sub
            End If
            Dim TEndPoint As Point3d = PlineBF.GetClosestPointTo(pPtRes.Value, False)
            '===========================================================================================================
            'вычисляем расстояния до выбранных точек
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
            If DistEnd - DistStart < 0.01 Then
                ed.WriteMessage(vbLf & "Начальная и конечная точки совпадают.")
                Exit Sub
            End If
            '============================================================================================================
            'добавляем выбранные точки в полилинию если конечно их нет
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

            '============================================================================================================
            'считаем вершины внутри промежутка и проверяем минимальное количество
            Dim countDelete As Integer = 0
            For i As Integer = 0 To PlineBF.NumberOfVertices - 1
                Dim distVertex As Double = PlineBF.GetDistAtPoint(PlineBF.GetPoint3dAt(i))
                If distVertex > DistStart + 0.01 And distVertex < DistEnd - 0.01 Then
                    countDelete += 1
                End If
            Next i
            If countDelete = 0 Then
                ed.WriteMessage(vbLf & "Внутри промежутка нет вершин.")
                Exit Sub
            End If
            If PlineBF.NumberOfVertices - countDelete < 3 Then
                MsgBox("Нельзя удалить вершины: в полилинии останется слишком мало вершин.")
                Exit Sub
            End If

            '============================================================================================================
            'удаляем вершины попавшие внутрь промежутка (полилиния обновляется на месте)
            Dim boolVertDel As Boolean = False
            Dim indexVert As Integer = 0
            While indexVert < PlineBF.NumberOfVertices
                Dim tempT As Point3d = PlineBF.GetPoint3dAt(indexVert)
                If tempT.GetVectorTo(TStartPoint).Length < 0.01 Then
                    boolVertDel = True
                    indexVert += 1
                    Continue While
                ElseIf tempT.GetVectorTo(TEndPoint).Length < 0.01 Then
                    Exit While
                End If
                If boolVertDel = True Then
                    PlineBF.RemoveVertexAt(indexVert)
                Else
                    indexVert += 1
                End If
            End While
            acTrans.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'разрезать полигон полилинией — результат: зона из 2 контуров ConditionalNumber=N(1) и N(2)
    <CommandMethod("SplitBufferZones")>
    Public Sub SplitBufferZones()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Try
                'полигон зоны
                Dim xformZone As Matrix3d = Matrix3d.Identity
                Dim boolTransformZone As Boolean = False
                Dim zonePoly As Polyline = SelectNestedPolyline(ed, tr, vbLf & "Выберите полилинию которую собираетесь разрезать: ", xformZone, boolTransformZone)
                If zonePoly Is Nothing Then
                    tr.Commit() : Exit Sub
                End If
                Dim zoneGeometry As Polygon = CreatePolygonFromPolyline(zonePoly, xformZone, boolTransformZone)
                If zoneGeometry Is Nothing Then
                    ed.WriteMessage(vbLf & "Выбранная полилиния не образует полигон.")
                    tr.Commit() : Exit Sub
                End If
                If IsPolylineEditable(tr, zonePoly) = False Then
                    ed.WriteMessage(vbLf & "Выбранный объект во внешней ссылке — редактирование невозможно.")
                    tr.Commit() : Exit Sub
                End If
                'XData зоны
                Dim dataZone As CadastralTools.StructureElement = Nothing
                FuncXRecords.getXRecords(zonePoly, dataZone)
                Dim conditionalNumber As String = ""
                If dataZone IsNot Nothing Then
                    conditionalNumber = CadastralTools.JsonValidator.GetValueFromJson(dataZone.KeyParameter, "ConditionalNumber")
                End If
                Dim numberObject As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalNumber)
                If numberObject = 0 Then
                    ed.WriteMessage(vbLf & "У объекта нет условного номера (ConditionalNumber) — разрезание невозможно.")
                    tr.Commit() : Exit Sub
                End If

                'режущая полилиния
                Dim xformCut As Matrix3d = Matrix3d.Identity
                Dim boolTransformCut As Boolean = False
                Dim cutPoly As Polyline = SelectNestedPolyline(ed, tr, vbLf & "Выберите полилинию которой будет происходить рез: ", xformCut, boolTransformCut)
                If cutPoly Is Nothing Then
                    tr.Commit() : Exit Sub
                End If
                Dim cutGeometry As LineString = CreateLineStringFromPolyline(cutPoly, xformCut, boolTransformCut)
                If cutGeometry Is Nothing Then
                    ed.WriteMessage(vbLf & "Режущая полилиния недопустима.")
                    tr.Commit() : Exit Sub
                End If

                'разрезаем (FuncNetTopologySuite)
                Dim processor As PolygonLineIntersectionProcessor = New PolygonLineIntersectionProcessor
                Dim listParts As List(Of Polygon) = processor.SplitPolygonWithLine(zoneGeometry, cutGeometry)
                If listParts Is Nothing OrElse listParts.Count < 2 Then
                    ed.WriteMessage(vbLf & "Полигон не разрезан.")
                    tr.Commit() : Exit Sub
                End If
                'если частей больше двух — берем две наибольшие
                If listParts.Count > 2 Then
                    listParts = listParts.OrderByDescending(Function(p) p.Area).Take(2).ToList()
                    ed.WriteMessage(vbLf & "Получено более двух частей — построены две наибольшие.")
                Else
                    listParts = listParts.OrderByDescending(Function(p) p.Area).ToList()
                End If

                'часть с центроидом исходной зоны обновляет исходную полилинию
                Dim pointZone As NetTopologySuite.Geometries.Point = zoneGeometry.Centroid
                Dim indexMain As Integer = 0
                If listParts.Item(1).Contains(pointZone) Then
                    indexMain = 1
                End If
                Dim partMain As Polygon = listParts.Item(indexMain)
                Dim partSecond As Polygon = listParts.Item(1 - indexMain)

                'обновляем исходную полилинию — XData сохраняются, прописываем ConditionalNumber = N(1)
                classDrawObject.createPolylineByPolygon(partMain, zonePoly)
                If dataZone IsNot Nothing Then
                    CopyXRecordsToPolyline(dataZone, zonePoly, numberObject.ToString & "(1)")
                End If

                'второй контур — новая полилиния с ConditionalNumber = N(2)
                Dim secondPoly As Polyline = classDrawObject.createPolylineByPolygon(partSecond)
                If Not IsNothing(secondPoly) Then
                    secondPoly.ColorIndex = zonePoly.ColorIndex
                    secondPoly.Linetype = zonePoly.Linetype
                    secondPoly.LinetypeScale = zonePoly.LinetypeScale
                    secondPoly.LineWeight = zonePoly.LineWeight
                    If dataZone IsNot Nothing Then
                        CopyXRecordsToPolyline(dataZone, secondPoly, numberObject.ToString & "(2)")
                    End If
                End If

                tr.Commit()
                ed.WriteMessage(vbLf & "Полигон разрезан: контура " & numberObject.ToString & "(1) и " & numberObject.ToString & "(2).")
            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Ошибка: " & ex.Message)
            End Try
        End Using
    End Sub


End Class
