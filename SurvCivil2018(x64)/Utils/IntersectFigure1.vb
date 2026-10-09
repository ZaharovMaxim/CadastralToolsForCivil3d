Imports System.Globalization
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Public Class IntersectFigure1
    Structure Point
        Public X As Double
        Public Y As Double
        Public Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub
    End Structure

    Function Inside(p As Point, e1 As Point, e2 As Point) As Boolean
        ' Возвращает True, если точка p находится слева от направленного отрезка e1->e2
        Return (e2.X - e1.X) * (p.Y - e1.Y) - (e2.Y - e1.Y) * (p.X - e1.X) >= 0
    End Function

    Function ComputeIntersection(p1 As Point, p2 As Point, e1 As Point, e2 As Point) As Point
        ' Находит точку пересечения линий p1->p2 и e1->e2
        Dim A1 As Double = p2.Y - p1.Y
        Dim B1 As Double = p1.X - p2.X
        Dim C1 As Double = A1 * p1.X + B1 * p1.Y

        Dim A2 As Double = e2.Y - e1.Y
        Dim B2 As Double = e1.X - e2.X
        Dim C2 As Double = A2 * e1.X + B2 * e1.Y

        Dim det As Double = A1 * B2 - A2 * B1
        If Math.Abs(det) < 0.0000000001 Then
            Return p1 ' Линии почти параллельны
        End If

        Dim x As Double = (B2 * C1 - B1 * C2) / det
        Dim y As Double = (A1 * C2 - A2 * C1) / det
        Return New Point(x, y)
    End Function

    Function Clip(subject As List(Of Point), clipPoly As List(Of Point)) As List(Of Point)
        Dim output As New List(Of Point)(subject)
        For i As Integer = 0 To clipPoly.Count - 1
            Dim input As List(Of Point) = New List(Of Point)(output)
            output.Clear()
            If input.Count = 0 Then Exit For

            Dim e1 As Point = clipPoly(i)
            Dim e2 As Point = clipPoly((i + 1) Mod clipPoly.Count)
            Dim S As Point = input(input.Count - 1)

            For Each E As Point In input
                If Inside(E, e1, e2) Then
                    If Not Inside(S, e1, e2) Then
                        output.Add(ComputeIntersection(S, E, e1, e2))
                    End If
                    output.Add(E)
                ElseIf Inside(S, e1, e2) Then
                    output.Add(ComputeIntersection(S, E, e1, e2))
                End If
                S = E
            Next
        Next
        Return output
    End Function

    Function IntersectPolygons(poly1 As List(Of Point), poly2 As List(Of Point)) As List(Of Point)
        Return Clip(poly1, poly2)
    End Function
    <CommandMethod("userPolygon1")>
    Public Sub userPolygon1()
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
                Dim polyList1 As New List(Of Point)()
                For i As Integer = 0 To poly1.NumberOfVertices - 1
                    Dim pt1 As Point2d = poly1.GetPoint2dAt(i)
                    polyList1.Add(New Point(pt1.X, pt1.Y))
                Next i

                Dim acSSPr2 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию 2: ")
                If Not (acSSPr2.Status = PromptStatus.OK) Then
                    Exit Sub
                End If

                Dim poly2 As Polyline = acTrans.GetObject(acSSPr2.ObjectId, OpenMode.ForWrite)
                Dim polyList2 As New List(Of Point)()
                For i As Integer = 0 To poly2.NumberOfVertices - 1
                    Dim pt1 As Point2d = poly2.GetPoint2dAt(i)
                    polyList2.Add(New Point(pt1.X, pt1.Y))
                Next i

                Dim arrayResult As List(Of Point) = IntersectPolygons(polyList1, polyList2)

                Dim newPoly As Polyline = New Polyline()
                newPoly.SetDatabaseDefaults()
                Dim count As Integer = 0
                For Each pt As Point In arrayResult
                    Dim pt2d As Point2d = New Point2d(pt.X, pt.Y)
                    newPoly.AddVertexAt(count, pt2d, 0, 0, 0)
                    count += 1
                Next



                acBlkTblRec.AppendEntity(newPoly)
                    acTrans.AddNewlyCreatedDBObject(newPoly, True)

                acTrans.Commit()
            End Using

        Catch ex As System.Exception
        End Try

    End Sub
    Sub Main()
        Dim ci As New CultureInfo("en-US")
        ' Чтение первого многоугольника
        Dim n1 As Integer = Integer.Parse(Console.ReadLine())
        Dim poly1 As New List(Of Point)()
        For i As Integer = 1 To n1
            Dim parts() As String = Console.ReadLine().Split()
            poly1.Add(New Point(Double.Parse(parts(0), ci), Double.Parse(parts(1), ci)))
        Next

        ' Чтение второго многоугольника
        Dim n2 As Integer = Integer.Parse(Console.ReadLine())
        Dim poly2 As New List(Of Point)()
        For i As Integer = 1 To n2
            Dim parts() As String = Console.ReadLine().Split()
            poly2.Add(New Point(Double.Parse(parts(0), ci), Double.Parse(parts(1), ci)))
        Next

        ' Вычисление пересечения и вывод результатов
        Dim result As List(Of Point) = IntersectPolygons(poly1, poly2)
        Console.WriteLine(result.Count)
        For Each pt As Point In result
            Console.WriteLine($"{pt.X:F6} {pt.Y:F6}")
        Next
    End Sub

End Class
