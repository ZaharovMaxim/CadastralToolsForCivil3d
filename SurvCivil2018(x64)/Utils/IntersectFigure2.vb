Imports System.Globalization
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Public Class IntersectFigure2
    Public Shared Function FindPolygonIntersections(polygon1Coords As Double(), polygon2Coords As Double()) As Polygon()
        ' Создаем геометрическую фабрику
        Dim geometryFactory As GeometryFactory = New GeometryFactory()

        ' Создаем первый многоугольник
        Dim polygon1 = CreatePolygon(geometryFactory, polygon1Coords)
        If polygon1 Is Nothing OrElse Not polygon1.IsValid Then
            Return Nothing
        End If

        ' Создаем второй многоугольник
        Dim polygon2 = CreatePolygon(geometryFactory, polygon2Coords)
        If polygon2 Is Nothing OrElse Not polygon2.IsValid Then
            Return Nothing
        End If

        ' Находим пересечение
        Dim intersection = polygon1.Intersection(polygon2)

        ' Если пересечение пустое
        If intersection Is Nothing OrElse intersection.IsEmpty Then
            Return Nothing
        End If

        ' Обрабатываем разные типы результата пересечения
        If TypeOf intersection Is Polygon Then
            ' Одно простое пересечение
            Return New Polygon() {DirectCast(intersection, Polygon)}
        ElseIf TypeOf intersection Is MultiPolygon Then
            ' Несколько областей пересечения
            Dim multiPoly = DirectCast(intersection, MultiPolygon)
            Dim result(multiPoly.NumGeometries - 1) As Polygon
            For i As Integer = 0 To multiPoly.NumGeometries - 1
                result(i) = DirectCast(multiPoly.GetGeometryN(i), Polygon)
                Dim newPoly As Polygon = result(i)
                Dim newPolyline As Polyline = New Polyline()
                newPolyline.SetDatabaseDefaults()
                For j As Integer = 0 To newPoly.NumPoints - 1
                    Dim coord As Coordinate = newPoly.Coordinates.ElementAt(j)
                    Dim pt2d As Point2d = New Point2d(coord.X, coord.Y)
                    newPolyline.AddVertexAt(j, pt2d, 0, 0, 0)
                Next j
                'acBlkTblRec.AppendEntity(newPolyline)
                'acTrans.AddNewlyCreatedDBObject(newPolyline, True)

                'acTrans.Commit()
            Next
            Return result
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Создает многоугольник из массива координат
    ''' </summary>
    Private Shared Function CreatePolygon(factory As GeometryFactory, coords As Double()) As Polygon
        If coords Is Nothing OrElse coords.Length < 6 Then
            Return Nothing ' Нужно минимум 3 точки (6 координат) для многоугольника
        End If

        ' Координаты должны образовывать замкнутую область (первая и последняя точки совпадают)
        Dim coordinates As New List(Of Coordinate)

        For i As Integer = 0 To coords.Length - 1 Step 2
            If i + 1 >= coords.Length Then Exit For
            coordinates.Add(New Coordinate(coords(i), coords(i + 1)))
        Next

        ' Замыкаем многоугольник, если нужно
        If coordinates(0).Equals(coordinates(coordinates.Count - 1)) = False Then
            coordinates.Add(New Coordinate(coords(0), coords(1)))
        End If

        Try
            Return factory.CreatePolygon(coordinates.ToArray())
        Catch ex As Exception
            ' В случае недопустимой геометрии
            Return Nothing
        End Try
    End Function

    <CommandMethod("userPolygon3")>
    Public Sub userPolygon3()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim acTrans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using acTrans
                Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acSSPr As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию 1: ")
                If Not (acSSPr.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim poly1 As Polyline = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
                Dim polyList1 As Double() = Nothing
                Dim count1 As Integer = 1

                For i As Integer = 0 To poly1.NumberOfVertices - 1
                    Dim pt1 As Point2d = poly1.GetPoint2dAt(i)
                    ReDim Preserve polyList1(count1)
                    polyList1(count1 - 1) = pt1.X
                    polyList1(count1) = pt1.Y
                    count1 += 2
                Next i

                Dim acSSPr2 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию 2: ")
                If Not (acSSPr2.Status = PromptStatus.OK) Then
                    Exit Sub
                End If

                Dim poly2 As Polyline = acTrans.GetObject(acSSPr2.ObjectId, OpenMode.ForWrite)
                Dim polyList2 As Double() = Nothing
                Dim count2 As Integer = 1
                For i As Integer = 0 To poly2.NumberOfVertices - 1
                    Dim pt1 As Point2d = poly2.GetPoint2dAt(i)
                    ReDim Preserve polyList2(count2)
                    polyList2(count2 - 1) = pt1.X
                    polyList2(count2) = pt1.Y
                    count2 += 2
                Next i

                Dim intersections = FindPolygonIntersections(polyList1, polyList2)

                If intersections IsNot Nothing Then
                    MsgBox("Найдено " & intersections.Length & " областей пересечения")
                    For Each polygon In intersections

                        Console.WriteLine($"Площадь области пересечения: {polygon.Area}")
                    Next
                Else
                    Console.WriteLine("Пересечений не найдено")
                End If

                'Dim newPoly As Polyline = New Polyline()
                'newPoly.SetDatabaseDefaults()
                'Dim count As Integer = 0
                'For Each pt As Point In arrayResult
                '    Dim pt2d As Point2d = New Point2d(pt.X, pt.Y)
                '    newPoly.AddVertexAt(count, pt2d, 0, 0, 0)
                '    count += 1
                'Next



                'acBlkTblRec.AppendEntity(newPoly)
                'acTrans.AddNewlyCreatedDBObject(newPoly, True)

                acTrans.Commit()
            End Using

        Catch ex As System.Exception
        End Try

    End Sub

End Class
