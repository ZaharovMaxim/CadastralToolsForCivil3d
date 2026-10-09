Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Operation.Distance
Imports NetTopologySuite.Operation.Union
Imports NetTopologySuite.Simplify

Public Class FuncNetTopologySuite
    Private ReadOnly _geometryFactory As GeometryFactory
    Public Sub New()
        ' Создаем фабрику геометрий с точностью 0.001
        _geometryFactory = New GeometryFactory(New PrecisionModel(1000))
    End Sub
    ' Создает полигон из полилинии
    Public Function CreatePolygon(ByVal polyline As Polyline) As Polygon
        If polyline.NumberOfVertices < 3 Then
            Return Nothing
        End If
        Dim points As List(Of (x As Double, y As Double)) = New List(Of (x As Double, y As Double))
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim vert As Point2d = polyline.GetPoint2dAt(i)
            points.Add((vert.X, vert.Y))
        Next i
        If polyline.NumberOfVertices > 2 Then
            If Not (points.Item(0).x = points.Item(points.Count - 1).x And points.Item(0).y = points.Item(points.Count - 1).y) Then
                points.Add((polyline.StartPoint.X, polyline.StartPoint.Y))
            End If
        End If
        Dim coordinates = points.Select(Function(p) New Coordinate(p.x, p.y)).ToArray()
        Dim linearRing = _geometryFactory.CreateLinearRing(coordinates)
        Return _geometryFactory.CreatePolygon(linearRing)
    End Function
    ' Создает полигон из массива точек
    Public Function CreatePolygon(coordinates(,) As Double) As Polygon
        If coordinates.GetLength(0) < 3 Then
            Return Nothing
        End If
        Dim linearRing = CreateLinearRing(coordinates)
        Return _geometryFactory.CreatePolygon(linearRing)
    End Function
    Private Function CreateLinearRing(coordinates(,) As Double) As LinearRing
        Dim pointCount = coordinates.GetLength(0)
        Dim coordArray = New Coordinate(pointCount + 1) {} ' +1 для замыкания

        For i = 0 To pointCount - 1
            coordArray(i) = New Coordinate(coordinates(i, 0), coordinates(i, 1))
        Next

        ' Замыкаем полигон
        coordArray(pointCount) = New Coordinate(coordinates(0, 0), coordinates(0, 1))

        Return _geometryFactory.CreateLinearRing(coordArray)
    End Function
    'создание полилинии
    Public Function CreateLineString(ByVal polyline As Polyline) As LineString
        If polyline.NumberOfVertices < 2 Then
            Return Nothing
        End If
        Dim points As List(Of (x As Double, y As Double)) = New List(Of (x As Double, y As Double))
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim vert As Point2d = polyline.GetPoint2dAt(i)
            points.Add((vert.X, vert.Y))
        Next i
        Dim coordinates = points.Select(Function(p) New Coordinate(p.x, p.y)).ToArray()
        Return _geometryFactory.CreateLineString(coordinates)
    End Function
    'создание линии
    Public Function CreateLineString(ByVal line As Line) As LineString
        Dim points As List(Of (x As Double, y As Double)) = New List(Of (x As Double, y As Double))
        points.Add((line.StartPoint.X, line.StartPoint.Y))
        points.Add((line.EndPoint.X, line.EndPoint.Y))
        Dim coordinates = points.Select(Function(p) New Coordinate(p.x, p.y)).ToArray()
        Return _geometryFactory.CreateLineString(coordinates)
    End Function

    '=============================================================================================================================
    'работа с полигононами
    '=============================================================================================================================
    ''' Проверяет, пересекаются ли два полигона
    Public Function DoPolygonsIntersect(polygon1 As Polygon, polygon2 As Polygon) As Boolean
        Try
            ' Проверяем валидность полигонов
            If Not polygon1.IsValid OrElse Not polygon2.IsValid Then
                Console.WriteLine("Один из полигонов невалиден")
                Return False
            End If

            Return polygon1.Intersects(polygon2)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Получает область пересечения полигонов
    Public Function GetIntersection(polygon1 As Polygon, polygon2 As Polygon) As Geometry
        Try
            If Not polygon1.Intersects(polygon2) Then
                Return Nothing
            End If
            Return polygon1.Intersection(polygon2)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' Получает детальную информацию о пересечении
    Public Function GetIntersectionInfo(polygon1 As Polygon, polygon2 As Polygon) As IntersectionInfo
        Dim info = New IntersectionInfo()
        Try
            info.Intersects = polygon1.Intersects(polygon2)
            info.Touches = polygon1.Touches(polygon2)
            info.Contains = polygon1.Contains(polygon2)
            info.Within = polygon1.Within(polygon2)
            info.Overlaps = polygon1.Overlaps(polygon2)
            info.Covers = polygon1.Covers(polygon2)
            info.CoveredBy = polygon1.CoveredBy(polygon2)
            info.Crosses = polygon1.Crosses(polygon2)

            If info.Intersects Then
                Dim intersection = polygon1.Intersection(polygon2)
                info.IntersectionArea = intersection.Area
                info.IntersectionGeometry = intersection
            End If

            info.UnionArea = polygon1.Union(polygon2).Area
            info.SymmetricDifferenceArea = polygon1.SymmetricDifference(polygon2).Area

            ' Вычисляем процент пересечения относительно каждого полигона
            If info.IntersectionArea > 0 Then
                info.IntersectionPercentagePolygon1 = (info.IntersectionArea / polygon1.Area) * 100
                info.IntersectionPercentagePolygon2 = (info.IntersectionArea / polygon2.Area) * 100
            End If
        Catch ex As Exception
        End Try

        Return info
    End Function
    ''' Проверяет, является ли полигон выпуклым
    Public Function IsConvex(polygon As Polygon) As Boolean
        Try
            Dim convexHull = polygon.ConvexHull()
            Return polygon.EqualsTopologically(convexHull)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Вычисляет расстояние между полигонами
    Public Function GetDistance(polygon1 As Polygon, polygon2 As Polygon) As Double
        Try
            Return polygon1.Distance(polygon2)
        Catch ex As Exception
            Return -1
        End Try
    End Function
    '============================================================================================================================
    'работа с полилинией
    '============================================================================================================================
    ''' Проверяет, пересекает ли полилиния полигон
    Public Function DoesLineIntersectPolygon(line As LineString, polygon As Polygon) As Boolean
        Try
            Return line.Intersects(polygon)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Проверяет, находится ли полилиния полностью внутри полигона
    Public Function IsLineInsidePolygon(line As LineString, polygon As Polygon) As Boolean
        Try
            Return polygon.Contains(line)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Проверяет, касается ли полилиния границы полигона
    Public Function DoesLineTouchPolygon(line As LineString, polygon As Polygon) As Boolean
        Try
            Return line.Touches(polygon)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Проверяет, пересекает ли полилиния границу полигона
    Public Function DoesLineCrossPolygon(line As LineString, polygon As Polygon) As Boolean
        Try
            Return line.Crosses(polygon)
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' Получает точки пересечения полилинии с полигоном
    Public Function GetIntersectionPoints(line As LineString, polygon As Polygon) As List(Of Point)
        Dim intersectionPoints As New List(Of Point)()

        Try
            Dim intersection = line.Intersection(polygon)

            If intersection IsNot Nothing Then
                If TypeOf intersection Is Point Then
                    intersectionPoints.Add(DirectCast(intersection, Point))
                ElseIf TypeOf intersection Is MultiPoint Then
                    Dim multiPoint = DirectCast(intersection, MultiPoint)
                    For i = 0 To multiPoint.NumGeometries - 1
                        intersectionPoints.Add(DirectCast(multiPoint.GetGeometryN(i), Point))
                    Next
                ElseIf TypeOf intersection Is GeometryCollection Then
                    Dim collection = DirectCast(intersection, GeometryCollection)
                    For i = 0 To collection.NumGeometries - 1
                        Dim geom = collection.GetGeometryN(i)
                        If TypeOf geom Is Point Then
                            intersectionPoints.Add(DirectCast(geom, Point))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
        End Try
        Return intersectionPoints
    End Function
    ''' Получает сегменты полилинии, которые находятся внутри полигона
    Public Function GetLineSegmentsInsidePolygon(line As LineString, polygon As Polygon) As List(Of LineString)
        Dim segmentsInside As New List(Of LineString)()

        Try
            Dim intersection = line.Intersection(polygon)

            If intersection IsNot Nothing Then
                If TypeOf intersection Is LineString Then
                    segmentsInside.Add(DirectCast(intersection, LineString))
                ElseIf TypeOf intersection Is MultiLineString Then
                    Dim multiLine = DirectCast(intersection, MultiLineString)
                    For i = 0 To multiLine.NumGeometries - 1
                        segmentsInside.Add(DirectCast(multiLine.GetGeometryN(i), LineString))
                    Next
                End If
            End If
        Catch ex As Exception
        End Try

        Return segmentsInside
    End Function
    ''' Получает длину полилинии внутри полигона
    Public Function GetLengthInsidePolygon(line As LineString, polygon As Polygon) As Double
        Try
            Dim intersection = line.Intersection(polygon)
            If intersection IsNot Nothing Then
                Return intersection.Length
            End If
            Return 0
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ''' Получает процент полилинии внутри полигона
    Public Function GetPercentageInsidePolygon(line As LineString, polygon As Polygon) As Double
        Try
            Dim totalLength = line.Length
            If totalLength = 0 Then Return 0

            Dim insideLength = GetLengthInsidePolygon(line, polygon)
            Return (insideLength / totalLength) * 100
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ''' Проверяет, пересекает ли полилиния границу полигона в указанной точке
    Public Function DoesLineCrossAtPoint(line As LineString, polygon As Polygon, point As Point) As Boolean
        Try
            Dim boundary = polygon.Boundary
            Return line.Intersects(boundary) AndAlso boundary.Intersects(point)
        Catch ex As Exception
            Return False
        End Try
    End Function
    '============================================================================================================================
    'нахождение внутренних точек полигона
    '============================================================================================================================
    ' Центроид (центр масс) - может быть вне полигона для вогнутых фигур
    Public Function GetCentroid(polygon As Polygon) As Point
        Return polygon.Centroid
    End Function
    ' Внутренняя точка - гарантированно внутри полигона
    Public Function GetInteriorPoint(polygon As Polygon) As Point
        Return polygon.InteriorPoint
    End Function
    ' Геометрический центр (центр bounding box) - может быть вне полигона
    Public Function GetGeometricCenter(polygon As Polygon) As Point
        Dim envelope = polygon.EnvelopeInternal
        Dim centerX = (envelope.MinX + envelope.MaxX) / 2
        Dim centerY = (envelope.MinY + envelope.MaxY) / 2
        Return _geometryFactory.CreatePoint(New Coordinate(centerX, centerY))
    End Function
    ' Центр максимальной вписанной окружности (пока не реализован в NTS)
    ' Альтернатива: точка внутри полигона, наиболее удаленная от границ
    Public Function GetPointFarthestFromBoundary(polygon As Polygon) As Point
        Try
            ' Используем внутреннюю точку как начальное приближение
            Dim interiorPoint = polygon.InteriorPoint

            ' Находим точку на наибольшем расстоянии от границы
            Dim farthestPoint = interiorPoint
            Dim maxDistance = polygon.Boundary.Distance(interiorPoint)

            ' Простой алгоритм поиска более удаленной точки
            Dim pointsToTest = GenerateTestPointsInsidePolygon(polygon, 20)

            For Each testPoint In pointsToTest
                If polygon.Contains(testPoint) Then
                    Dim distance = polygon.Boundary.Distance(testPoint)
                    If distance > maxDistance Then
                        maxDistance = distance
                        farthestPoint = testPoint
                    End If
                End If
            Next

            Return farthestPoint
        Catch ex As Exception
            Return polygon.InteriorPoint
        End Try
    End Function
    ''' Генерирует тестовые точки внутри полигона
    Private Function GenerateTestPointsInsidePolygon(polygon As Polygon, count As Integer) As List(Of Point)
        Dim points As New List(Of Point)()
        Dim envelope = polygon.EnvelopeInternal
        Dim random As New Random()

        For i = 1 To count
            Dim x = envelope.MinX + random.NextDouble() * envelope.Width
            Dim y = envelope.MinY + random.NextDouble() * envelope.Height
            Dim point = _geometryFactory.CreatePoint(New Coordinate(x, y))

            If polygon.Contains(point) Then
                points.Add(point)
            End If
        Next

        Return points
    End Function
    ' Проверяет, находится ли точка внутри полигона
    Public Function IsPointInsidePolygon(point As Point, polygon As Polygon) As Boolean
        Return polygon.Contains(point)
    End Function
    ' Находит ближайшую точку внутри полигона к заданной точке
    Public Function GetNearestPointInsidePolygon(referencePoint As Point, polygon As Polygon) As Point
        Try
            If polygon.Contains(referencePoint) Then
                Return referencePoint
            End If

            ' Используем DistanceOp для нахождения ближайшей точки
            Dim distanceOp = New DistanceOp(polygon, referencePoint)
            Dim nearestPoints = distanceOp.NearestPoints()

            If nearestPoints IsNot Nothing AndAlso nearestPoints.Length >= 2 Then
                Return _geometryFactory.CreatePoint(nearestPoints(0))
            End If

            Return polygon.InteriorPoint
        Catch ex As Exception
            Return polygon.InteriorPoint
        End Try
    End Function
    ' Сравнивает разные методы нахождения центра
    Public Sub CompareCenterMethods(polygon As Polygon, description As String)
        Console.WriteLine($"=== {description} ===")

        Dim centroid = GetCentroid(polygon)
        Dim interiorPoint = GetInteriorPoint(polygon)
        Dim geometricCenter = GetGeometricCenter(polygon)
        Dim farthestPoint = GetPointFarthestFromBoundary(polygon)

        Console.WriteLine($"Центроид:          ({centroid.X:F2}, {centroid.Y:F2}) - Внутри: {IsPointInsidePolygon(centroid, polygon)}")
        Console.WriteLine($"Внутренняя точка:  ({interiorPoint.X:F2}, {interiorPoint.Y:F2}) - Внутри: {IsPointInsidePolygon(interiorPoint, polygon)}")
        Console.WriteLine($"Геометрический:    ({geometricCenter.X:F2}, {geometricCenter.Y:F2}) - Внутри: {IsPointInsidePolygon(geometricCenter, polygon)}")
        Console.WriteLine($"Удаленная от границы: ({farthestPoint.X:F2}, {farthestPoint.Y:F2}) - Внутри: {IsPointInsidePolygon(farthestPoint, polygon)}")

        Console.WriteLine()
    End Sub
    '========================================================================================================
    'объединение полигонов
    '========================================================================================================
    ' Простое объединение двух полигонов
    Public Function UnionTwoPolygons(polygon1 As Polygon, polygon2 As Polygon) As Geometry
        Try
            Return polygon1.Union(polygon2)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ' Каскадное объединение нескольких полигонов (более эффективное)
    Public Function UnionMultiplePolygons(polygons As IEnumerable(Of Polygon)) As Geometry
        Try
            If polygons Is Nothing OrElse Not polygons.Any() Then
                Return Nothing
            End If

            If polygons.Count() = 1 Then
                Return polygons.First()
            End If

            ' Используем каскадный union для большей эффективности
            Dim unionOp = New CascadedPolygonUnion(polygons)
            Return unionOp.Union()
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ' Объединение через буферизацию (для сложных случаев)
    Public Function UnionViaBuffer(polygons As IEnumerable(Of Polygon), bufferDistance As Double) As Geometry
        Try
            Dim geometryCollection = _geometryFactory.CreateGeometryCollection(polygons.Cast(Of Geometry)().ToArray())
            Return geometryCollection.Buffer(bufferDistance)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ' Объединение с предварительным упрощением
    Public Function UnionWithSimplification(polygons As IEnumerable(Of Polygon), tolerance As Double) As Geometry
        Try
            Dim unionResult = UnionMultiplePolygons(polygons)
            If unionResult IsNot Nothing Then
                Return SimplifyGeometry(unionResult, tolerance)
            End If
            Return Nothing
        Catch ex As Exception
            Console.WriteLine($"Ошибка при объединении с упрощением: {ex.Message}")
            Return Nothing
        End Try
    End Function
    Public Function SimplifyGeometry(geometry As Geometry, tolerance As Double) As Geometry
        Try
            If geometry Is Nothing Then
                Return Nothing
            End If

            ' Используем TopologyPreservingSimplifier для сохранения топологии
            Dim simplifier = New TopologyPreservingSimplifier(geometry)
            simplifier.DistanceTolerance = tolerance
            Return simplifier.GetResultGeometry()
        Catch ex As Exception
            Return geometry ' Возвращаем исходную геометрию в случае ошибки
        End Try
    End Function
    ' Раздельное объединение (возвращает коллекцию отдельных полигонов)
    Public Function UnionSeparate(polygons As IEnumerable(Of Polygon)) As List(Of Polygon)
        Dim result As New List(Of Polygon)()

        Try
            Dim unionResult = UnionMultiplePolygons(polygons)
            If unionResult IsNot Nothing Then
                If TypeOf unionResult Is Polygon Then
                    result.Add(DirectCast(unionResult, Polygon))
                ElseIf TypeOf unionResult Is MultiPolygon Then
                    Dim multiPoly = DirectCast(unionResult, MultiPolygon)
                    For i = 0 To multiPoly.NumGeometries - 1
                        result.Add(DirectCast(multiPoly.GetGeometryN(i), Polygon))
                    Next
                End If
            End If
        Catch ex As Exception
        End Try
        Return result
    End Function
    ' Проверяет, можно ли объединить полигоны
    Public Function CanUnion(polygon1 As Polygon, polygon2 As Polygon) As Boolean
        Try
            Return polygon1.Intersects(polygon2) OrElse polygon1.Touches(polygon2)
        Catch ex As Exception
            Console.WriteLine($"Ошибка при проверке возможности объединения: {ex.Message}")
            Return False
        End Try
    End Function
    ' Получает площадь объединения
    Public Function GetUnionArea(polygons As IEnumerable(Of Polygon)) As Double
        Try
            Dim unionResult = UnionMultiplePolygons(polygons)
            If unionResult IsNot Nothing Then
                Return unionResult.Area
            End If
            Return 0
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ' Вычисляет выгоду от объединения (сокращение площади)
    Public Function GetUnionBenefit(polygons As IEnumerable(Of Polygon)) As Double
        Try
            Dim totalArea = polygons.Sum(Function(p) p.Area)
            Dim unionArea = GetUnionArea(polygons)
            Return totalArea - unionArea ' Положительное значение = выгода
        Catch ex As Exception
            Return 0
        End Try
    End Function
    ' Визуализирует процесс объединения
    Public Sub VisualizeUnionProcess(polygons As List(Of Polygon), description As String)
        Console.WriteLine($"=== {description} ===")
        Console.WriteLine($"Количество полигонов: {polygons.Count}")
        Console.WriteLine($"Общая площадь отдельных полигонов: {polygons.Sum(Function(p) p.Area):F2}")

        For i = 0 To polygons.Count - 1
            Console.WriteLine($"Полигон {i + 1}: площадь = {polygons(i).Area:F2}, вершин = {polygons(i).NumPoints}")
        Next

        Dim unionResult = UnionMultiplePolygons(polygons)
        If unionResult IsNot Nothing Then
            Console.WriteLine($"Площадь объединения: {unionResult.Area:F2}")
            Console.WriteLine($"Выгода от объединения: {GetUnionBenefit(polygons):F2}")
            Console.WriteLine($"Тип результата: {unionResult.GeometryType}")

            If TypeOf unionResult Is MultiPolygon Then
                Dim multiPoly = DirectCast(unionResult, MultiPolygon)
                Console.WriteLine($"Количество полигонов после объединения: {multiPoly.NumGeometries}")
            End If
        End If
        Console.WriteLine()
    End Sub
    Public Function CreateArcBetweenLines(line1 As LineString, line2 As LineString,
                                     radius As Double, Optional segments As Integer = 10) As Geometry

        ' Находим точку соединения
        Dim end1 = line1.EndPoint
        Dim start2 = line2.StartPoint

        If Not end1.Equals(start2) Then
            Throw New ArgumentException("Отрезки не соединены")
        End If

        Dim jointPoint = end1.Coordinate

        ' Находим направления отрезков
        Dim dir1 = CalculateDirection(line1)
        Dim dir2 = CalculateDirection(line2)

        ' Вычисляем угол между отрезками
        Dim angle1 = Math.Atan2(dir1.Y, dir1.X)
        Dim angle2 = Math.Atan2(dir2.Y, dir2.X)
        Dim angle = angle2 - angle1

        ' Создаем дугу
        Dim arcPoints As New List(Of Coordinate)

        For i As Integer = 0 To segments
            Dim t As Double = CDbl(i) / segments
            Dim currentAngle = angle1 + angle * t

            Dim x = jointPoint.X + radius * Math.Cos(currentAngle)
            Dim y = jointPoint.Y + radius * Math.Sin(currentAngle)

            arcPoints.Add(New Coordinate(x, y))
        Next

        Return New LineString(arcPoints.ToArray())
    End Function
    Private Function CalculateDirection(line As LineString) As Coordinate
        Dim start = line.StartPoint.Coordinate
        Dim [end] = line.EndPoint.Coordinate

        Return New Coordinate([end].X - start.X, [end].Y - start.Y)
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '3 функции для получения полигонов полученных при пересечении полигонов
    Public Function GetPolygonsOutsideReference(referencePolygon As Polygon, targetPolygon As Polygon) As List(Of Polygon)
        Dim resultPolygons As New List(Of Polygon)()

        ' Метод 1: Простая разность - что targetPolygon находится вне referencePolygon
        Dim outsidePolygons = targetPolygon.Difference(referencePolygon)
        ExtractAllPolygons(outsidePolygons, resultPolygons)

        Return resultPolygons
    End Function
    ' Получает полигоны из targetPolygon, которые полностью находятся вне referencePolygon
    Public Function GetCompletelyOutsidePolygons(referencePolygon As Polygon, targetPolygon As Polygon) As List(Of Polygon)
        Dim completelyOutside As New List(Of Polygon)()
        ' Разбиваем targetPolygon на части относительно referencePolygon
        Dim difference = targetPolygon.Difference(referencePolygon)
        Dim allParts As New List(Of Polygon)()
        ExtractAllPolygons(difference, allParts)
        ' Фильтруем только те полигоны, которые не пересекаются с referencePolygon
        For Each part In allParts
            If Not part.Intersects(referencePolygon) Then
                completelyOutside.Add(part)
            End If
        Next
        Return completelyOutside
    End Function
    ' Получает все полигоны из коллекции, которые находятся вне referencePolygon
    Public Function FilterPolygonsOutsideReference(referencePolygon As Polygon, polygonsToFilter As List(Of Polygon)) As List(Of Polygon)
        Dim filteredPolygons As New List(Of Polygon)()
        For Each poly In polygonsToFilter
            ' Если полигон полностью вне referencePolygon
            If Not poly.Intersects(referencePolygon) Then
                filteredPolygons.Add(poly)
            Else
                ' Если есть пересечение, берем только часть вне referencePolygon
                Dim outsidePart = poly.Difference(referencePolygon)
                ExtractAllPolygons(outsidePart, filteredPolygons)
            End If
        Next
        Return filteredPolygons
    End Function

    Private Sub ExtractAllPolygons(geometry As Geometry, ByRef polygonList As List(Of Polygon))
        If geometry Is Nothing OrElse geometry.IsEmpty Then
            Return
        End If
        If TypeOf geometry Is Polygon Then
            polygonList.Add(DirectCast(geometry, Polygon))
        ElseIf TypeOf geometry Is MultiPolygon Then
            Dim multiPoly = DirectCast(geometry, MultiPolygon)
            For i As Integer = 0 To multiPoly.NumGeometries - 1
                polygonList.Add(DirectCast(multiPoly.GetGeometryN(i), Polygon))
            Next
        ElseIf TypeOf geometry Is GeometryCollection Then
            Dim collection = DirectCast(geometry, GeometryCollection)
            For i As Integer = 0 To collection.NumGeometries - 1
                Dim geom = collection.GetGeometryN(i)
                If TypeOf geom Is Polygon Then
                    polygonList.Add(DirectCast(geom, Polygon))
                End If
            Next
        End If
    End Sub

    Public Function GetPolygonsOutsidePoly1(poly1 As Polygon, poly2 As Polygon) As List(Of Polygon)
        Dim polygonsOutside As New List(Of Polygon)()
        ' Получаем часть poly2, которая находится вне poly1
        Dim poly2Outside = poly2.Difference(poly1)
        ' Извлекаем все полигоны из результата
        ExtractPolygons(poly2Outside, polygonsOutside)
        Return polygonsOutside
    End Function

    Private Sub ExtractPolygons(geometry As Geometry, ByRef polygonList As List(Of Polygon))
        If geometry Is Nothing OrElse geometry.IsEmpty Then
            Return
        End If

        Select Case geometry.GeometryType
            Case "Polygon"
                polygonList.Add(DirectCast(geometry, Polygon))

            Case "MultiPolygon"
                Dim multiPoly = DirectCast(geometry, MultiPolygon)
                For i As Integer = 0 To multiPoly.NumGeometries - 1
                    polygonList.Add(DirectCast(multiPoly.GetGeometryN(i), Polygon))
                Next

            Case "GeometryCollection"
                Dim collection = DirectCast(geometry, GeometryCollection)
                For i As Integer = 0 To collection.NumGeometries - 1
                    ExtractPolygons(collection.GetGeometryN(i), polygonList)
                Next
        End Select
    End Sub



End Class
Public Class IntersectionInfo
    Public Property Intersects As Boolean
    Public Property Touches As Boolean
    Public Property Contains As Boolean
    Public Property Within As Boolean
    Public Property Overlaps As Boolean
    Public Property Covers As Boolean
    Public Property CoveredBy As Boolean
    Public Property Crosses As Boolean
    Public Property IntersectionArea As Double
    Public Property UnionArea As Double
    Public Property SymmetricDifferenceArea As Double
    Public Property IntersectionPercentagePolygon1 As Double
    Public Property IntersectionPercentagePolygon2 As Double
    Public Property IntersectionGeometry As Geometry

    Public Sub PrintInfo()
        Console.WriteLine("=== ИНФОРМАЦИЯ О ПЕРЕСЕЧЕНИИ ===")
        Console.WriteLine($"Пересекаются: {Intersects}")
        Console.WriteLine($"Соприкасаются: {Touches}")
        Console.WriteLine($"Polygon1 содержит Polygon2: {Contains}")
        Console.WriteLine($"Polygon1 внутри Polygon2: {Within}")
        Console.WriteLine($"Перекрываются: {Overlaps}")
        Console.WriteLine($"Polygon1 покрывает Polygon2: {Covers}")
        Console.WriteLine($"Polygon1 покрывается Polygon2: {CoveredBy}")
        Console.WriteLine($"Пересекаются (crosses): {Crosses}")

        If Intersects Then
            Console.WriteLine($"Площадь пересечения: {IntersectionArea:F4}")
            Console.WriteLine($"Процент пересечения от Polygon1: {IntersectionPercentagePolygon1:F2}%")
            Console.WriteLine($"Процент пересечения от Polygon2: {IntersectionPercentagePolygon2:F2}%")
        End If

        Console.WriteLine($"Площадь объединения: {UnionArea:F4}")
        Console.WriteLine($"Площадь симметричной разности: {SymmetricDifferenceArea:F4}")
        Console.WriteLine("================================")
        Console.WriteLine()
    End Sub
End Class

Public Class PolygonLineIntersectionInfo
    Public Property Intersects As Boolean
    Public Property IsInside As Boolean
    Public Property TouchesBoundary As Boolean
    Public Property CrossesBoundary As Boolean
    Public Property IntersectionPoints As List(Of Point)
    Public Property SegmentsInside As List(Of LineString)
    Public Property LengthInside As Double
    Public Property PercentageInside As Double
    Public Property TotalIntersections As Integer

    Public Sub New()
        IntersectionPoints = New List(Of Point)()
        SegmentsInside = New List(Of LineString)()
    End Sub




End Class
Public Class ArcCreator

    Public Function CreateSmoothConnection(line1 As LineString, line2 As LineString,
                                         radius As Double) As Geometry

        ' Проверяем соединение отрезков
        If Not line1.EndPoint.Equals(line2.StartPoint) Then
            Return Nothing
        End If

        Dim joint = line1.EndPoint.Coordinate

        ' Находим контрольные точки для дуги
        Dim controlPoint1 = FindOffsetPoint(line1, joint, radius)
        Dim controlPoint2 = FindOffsetPoint(line2, joint, radius)

        ' Создаем дугу как часть окружности
        Dim arc = CreateCircularArc(controlPoint1, joint, controlPoint2, 10)

        Return arc
    End Function

    Private Function FindOffsetPoint(line As LineString, fromPoint As Coordinate,
                                   distance As Double) As Coordinate

        Dim direction = CalculateDirection(line, fromPoint)
        Return New Coordinate(
            fromPoint.X - direction.X * distance,
            fromPoint.Y - direction.Y * distance
        )
    End Function

    Private Function CalculateDirection(line As LineString, fromPoint As Coordinate) As Coordinate
        ' Упрощенный расчет направления
        If fromPoint.Equals(line.StartPoint.Coordinate) Then
            Dim nextCoord = line.Coordinates(1)
            Return New Coordinate(nextCoord.X - fromPoint.X, nextCoord.Y - fromPoint.Y)
        Else
            Dim prev = line.Coordinates(line.Coordinates.Length - 2)
            Return New Coordinate(fromPoint.X - prev.X, fromPoint.Y - prev.Y)
        End If
    End Function

    Private Function CreateCircularArc(p1 As Coordinate, center As Coordinate,
                                     p2 As Coordinate, segments As Integer) As Geometry

        Dim points As New List(Of Coordinate) From {p1}

        ' Вычисляем углы
        Dim angle1 = Math.Atan2(p1.Y - center.Y, p1.X - center.X)
        Dim angle2 = Math.Atan2(p2.Y - center.Y, p2.X - center.X)

        ' Создаем промежуточные точки
        For i As Integer = 1 To segments - 1
            Dim t As Double = CDbl(i) / segments
            Dim angle = angle1 + (angle2 - angle1) * t
            Dim radius = CalculateDistance(p1, center)

            Dim x = center.X + radius * Math.Cos(angle)
            Dim y = center.Y + radius * Math.Sin(angle)

            points.Add(New Coordinate(x, y))
        Next

        points.Add(p2)

        Return New LineString(points.ToArray())
    End Function

    Private Function CalculateDistance(p1 As Coordinate, p2 As Coordinate) As Double
        Dim dx = p2.X - p1.X
        Dim dy = p2.Y - p1.Y
        Return Math.Sqrt(dx * dx + dy * dy)
    End Function

End Class
Public Class PolygonLineIntersectionProcessor

    Private _geometryFactory As GeometryFactory

    Public Sub New()
        _geometryFactory = New GeometryFactory()
    End Sub

    ''' <summary>
    ''' Основной метод: разбивает полигон полилинией и возвращает все образовавшиеся полигоны
    ''' </summary>
    Public Function SplitPolygonWithLine(polygon As Polygon, lineString As LineString) As List(Of Polygon)
        Dim resultPolygons As New List(Of Polygon)()

        ' Проверяем, пересекаются ли полигон и полилиния
        If Not polygon.Intersects(lineString) Then
            ' Если не пересекаются, возвращаем исходный полигон
            resultPolygons.Add(polygon)
            Return resultPolygons
        End If

        ' Создаем буфер вокруг полилинии (очень маленький)
        Dim bufferDistance As Double = 0.001
        Dim lineBuffer As Geometry = lineString.Buffer(bufferDistance)

        ' Используем симметрическую разность для разбиения полигона
        Dim splitResult As Geometry = polygon.SymmetricDifference(lineBuffer)

        ' Извлекаем все полигоны из результата
        ExtractAllPolygons(splitResult, resultPolygons)

        ' Фильтруем слишком маленькие полигоны (опционально)
        resultPolygons = FilterSmallPolygons(resultPolygons, 0.01)

        Return resultPolygons
    End Function

    ''' <summary>
    ''' Альтернативный метод: использует разность и объединение для более точного разбиения
    ''' </summary>
    Public Function SplitPolygonWithLineAdvanced(polygon As Polygon, lineString As LineString) As List(Of Polygon)
        Dim resultPolygons As New List(Of Polygon)()

        If Not polygon.Intersects(lineString) Then
            resultPolygons.Add(polygon)
            Return resultPolygons
        End If

        ' Создаем буфер вокруг линии
        Dim bufferDistance As Double = GetOptimalBufferDistance(polygon)
        Dim lineBuffer As Geometry = lineString.Buffer(bufferDistance)

        ' Получаем пересечение буфера линии с полигоном
        Dim intersection As Geometry = polygon.Intersection(lineBuffer)

        If intersection.IsEmpty Then
            resultPolygons.Add(polygon)
            Return resultPolygons
        End If

        ' Вычитаем пересечение из исходного полигона
        Dim difference As Geometry = polygon.Difference(intersection)

        ' Добавляем все части
        ExtractAllPolygons(difference, resultPolygons)

        Return resultPolygons
    End Function

    ''' <summary>
    ''' Метод для получения полигонов по разные стороны от полилинии
    ''' </summary>
    Public Function GetPolygonsOnBothSides(polygon As Polygon, lineString As LineString) As Dictionary(Of String, List(Of Polygon))
        Dim result As New Dictionary(Of String, List(Of Polygon))()

        Dim allSplitPolygons = SplitPolygonWithLine(polygon, lineString)

        ' Разделяем полигоны на те, что по разные стороны от линии
        Dim leftSide As New List(Of Polygon)()
        Dim rightSide As New List(Of Polygon)()

        For Each poly In allSplitPolygons
            Dim position = DeterminePolygonPosition(poly, lineString)
            If position = "left" Then
                leftSide.Add(poly)
            ElseIf position = "right" Then
                rightSide.Add(poly)
            Else
                ' Если не удалось определить, добавляем в оба списка
                leftSide.Add(poly)
                rightSide.Add(poly)
            End If
        Next

        result("left") = leftSide
        result("right") = rightSide
        result("all") = allSplitPolygons

        Return result
    End Function

    Private Function GetOptimalBufferDistance(polygon As Polygon) As Double
        ' Вычисляем оптимальное расстояние для буфера на основе размера полигона
        Dim envelope = polygon.EnvelopeInternal
        Dim minSize = Math.Min(envelope.Width, envelope.Height)
        Return minSize * 0.001 ' 0.1% от минимального размера
    End Function

    Private Sub ExtractAllPolygons(geometry As Geometry, ByRef polygonList As List(Of Polygon))
        If geometry Is Nothing OrElse geometry.IsEmpty Then
            Return
        End If

        Select Case geometry.GeometryType
            Case "Polygon"
                polygonList.Add(DirectCast(geometry, Polygon))

            Case "MultiPolygon"
                Dim multiPoly = DirectCast(geometry, MultiPolygon)
                For i As Integer = 0 To multiPoly.NumGeometries - 1
                    polygonList.Add(DirectCast(multiPoly.GetGeometryN(i), Polygon))
                Next

            Case "GeometryCollection"
                Dim collection = DirectCast(geometry, GeometryCollection)
                For i As Integer = 0 To collection.NumGeometries - 1
                    ExtractAllPolygons(collection.GetGeometryN(i), polygonList)
                Next
        End Select
    End Sub

    Private Function FilterSmallPolygons(polygons As List(Of Polygon), minArea As Double) As List(Of Polygon)
        Dim filtered As New List(Of Polygon)()

        For Each poly In polygons
            If poly.Area >= minArea Then
                filtered.Add(poly)
            End If
        Next

        Return filtered
    End Function

    Private Function DeterminePolygonPosition(polygon As Polygon, lineString As LineString) As String
        ' Упрощенный метод определения положения полигона относительно линии
        Try
            Dim centroid = polygon.Centroid
            Dim firstPoint = lineString.GetCoordinateN(0)
            Dim lastPoint = lineString.GetCoordinateN(lineString.NumPoints - 1)

            ' Вычисляем вектор направления линии
            Dim lineVector As New Coordinate(lastPoint.X - firstPoint.X, lastPoint.Y - firstPoint.Y)

            ' Вычисляем вектор от первой точки линии к центроиду полигона
            Dim toCentroidVector As New Coordinate(centroid.X - firstPoint.X, centroid.Y - firstPoint.Y)

            ' Векторное произведение для определения стороны
            Dim crossProduct = lineVector.X * toCentroidVector.Y - lineVector.Y * toCentroidVector.X

            If crossProduct > 0 Then
                Return "left"
            ElseIf crossProduct < 0 Then
                Return "right"
            Else
                Return "unknown"
            End If
        Catch ex As Exception
            Return "unknown"
        End Try
    End Function
End Class

Public Class PolygonIntersectionInnerProcessor

    Private _geometryFactory As GeometryFactory

    Public Sub New()
        _geometryFactory = New GeometryFactory()
    End Sub

    ''' <summary>
    ''' Основной метод: возвращает только полигоны из зоны пересечения
    ''' </summary>
    Public Function GetIntersectionPolygons(polygon1 As Polygon, polygon2 As Polygon) As List(Of Polygon)
        Dim intersectionPolygons As New List(Of Polygon)()

        ' Получаем пересечение двух полигонов
        Dim intersection = polygon1.Intersection(polygon2)

        ' Извлекаем все полигоны из результата пересечения
        ExtractAllPolygons(intersection, intersectionPolygons)

        Return intersectionPolygons
    End Function

    ''' <summary>
    ''' Метод с проверкой на существование пересечения
    ''' </summary>
    Public Function GetIntersectionPolygonsWithCheck(polygon1 As Polygon, polygon2 As Polygon) As List(Of Polygon)
        Dim intersectionPolygons As New List(Of Polygon)()

        ' Проверяем, пересекаются ли полигоны
        If Not polygon1.Intersects(polygon2) Then
            Console.WriteLine("Полигоны не пересекаются")
            Return intersectionPolygons ' возвращаем пустой список
        End If

        ' Получаем пересечение
        Dim intersection = polygon1.Intersection(polygon2)

        If intersection.IsEmpty Then
            Console.WriteLine("Пересечение пустое")
            Return intersectionPolygons
        End If

        ' Извлекаем полигоны
        ExtractAllPolygons(intersection, intersectionPolygons)

        Console.WriteLine($"Найдено полигонов в зоне пересечения: {intersectionPolygons.Count}")
        Console.WriteLine($"Общая площадь пересечения: {intersection.Area:F4}")

        Return intersectionPolygons
    End Function

    ''' <summary>
    ''' Метод с фильтрацией по минимальной площади
    ''' </summary>
    Public Function GetIntersectionPolygonsWithFilter(polygon1 As Polygon, polygon2 As Polygon, Optional minArea As Double = 0.001) As List(Of Polygon)
        Dim allIntersectionPolygons = GetIntersectionPolygons(polygon1, polygon2)
        Dim filteredPolygons As New List(Of Polygon)()

        ' Фильтруем полигоны по минимальной площади
        For Each poly In allIntersectionPolygons
            If poly.Area >= minArea Then
                filteredPolygons.Add(poly)
            End If
        Next

        If allIntersectionPolygons.Count <> filteredPolygons.Count Then
            Console.WriteLine($"Отфильтровано {allIntersectionPolygons.Count - filteredPolygons.Count} маленьких полигонов")
        End If

        Return filteredPolygons
    End Function

    ''' <summary>
    ''' Получает пересечение для нескольких пар полигонов
    ''' </summary>
    Public Function GetMultipleIntersections(polygonPairs As List(Of (Polygon, Polygon))) As Dictionary(Of Integer, List(Of Polygon))
        Dim results As New Dictionary(Of Integer, List(Of Polygon))()

        For i As Integer = 0 To polygonPairs.Count - 1
            Dim pair = polygonPairs(i)
            Dim intersection = GetIntersectionPolygons(pair.Item1, pair.Item2)
            results(i) = intersection
        Next

        Return results
    End Function

    Private Sub ExtractAllPolygons(geometry As Geometry, ByRef polygonList As List(Of Polygon))
        If geometry Is Nothing OrElse geometry.IsEmpty Then
            Return
        End If

        Select Case geometry.GeometryType
            Case "Polygon"
                polygonList.Add(DirectCast(geometry, Polygon))

            Case "MultiPolygon"
                Dim multiPoly = DirectCast(geometry, MultiPolygon)
                For i As Integer = 0 To multiPoly.NumGeometries - 1
                    polygonList.Add(DirectCast(multiPoly.GetGeometryN(i), Polygon))
                Next

            Case "GeometryCollection"
                Dim collection = DirectCast(geometry, GeometryCollection)
                For i As Integer = 0 To collection.NumGeometries - 1
                    Dim geom = collection.GetGeometryN(i)
                    If TypeOf geom Is Polygon Then
                        polygonList.Add(DirectCast(geom, Polygon))
                    End If
                Next
        End Select
    End Sub
End Class