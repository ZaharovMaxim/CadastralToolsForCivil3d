Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class IntersectFigure
    Public Function GetIntersectedPolygon(subjectPolygon As Point2d(), clipPolygon As Point2d()) As Point2d()
        If subjectPolygon.Length < 3 OrElse clipPolygon.Length < 3 Then
            Return New Point2d() {}
        End If

        Dim outputList As New List(Of Point2d)(subjectPolygon)

        For i As Integer = 0 To clipPolygon.Length - 1
            Dim clipEdgeStart As Point2d = clipPolygon(i)
            Dim clipEdgeEnd As Point2d = clipPolygon((i + 1) Mod clipPolygon.Length)

            Dim inputList As New List(Of Point2d)(outputList)
            outputList.Clear()

            If inputList.Count = 0 Then
                Exit For
            End If

            Dim S As Point2d = inputList(inputList.Count - 1)

            For Each E In inputList
                If IsInside(clipEdgeStart, clipEdgeEnd, E) Then
                    If Not IsInside(clipEdgeStart, clipEdgeEnd, S) Then
                        outputList.Add(ComputeIntersection(S, E, clipEdgeStart, clipEdgeEnd))
                    End If
                    outputList.Add(E)
                ElseIf IsInside(clipEdgeStart, clipEdgeEnd, S) Then
                    outputList.Add(ComputeIntersection(S, E, clipEdgeStart, clipEdgeEnd))
                End If
                S = E
            Next
        Next

        Return outputList.ToArray()
    End Function

    Private Function IsInside(edgeStart As Point2d, edgeEnd As Point2d, testPoint As Point2d) As Boolean
        Dim crossProduct As Single = (edgeEnd.X - edgeStart.X) * (testPoint.Y - edgeStart.Y) -
                                    (edgeEnd.Y - edgeStart.Y) * (testPoint.X - edgeStart.X)
        Return crossProduct >= 0
    End Function

    Private Function ComputeIntersection(line1Start As Point2d, line1End As Point2d,
                                        line2Start As Point2d, line2End As Point2d) As Point2d
        Dim A1 As Single = line1End.Y - line1Start.Y
        Dim B1 As Single = line1Start.X - line1End.X
        Dim C1 As Single = A1 * line1Start.X + B1 * line1Start.Y

        Dim A2 As Single = line2End.Y - line2Start.Y
        Dim B2 As Single = line2Start.X - line2End.X
        Dim C2 As Single = A2 * line2Start.X + B2 * line2Start.Y

        Dim det As Single = A1 * B2 - A2 * B1

        If det = 0 Then
            Return Nothing ' Параллельные линии
        Else
            Dim x As Single = (B2 * C1 - B1 * C2) / det
            Dim y As Single = (A1 * C2 - A2 * C1) / det
            Return New Point2d(x, y)
        End If
    End Function

    <CommandMethod("userPolygon")>
    Public Sub userPolygon()
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
                Dim arrayPoint1 As Point2d() = {}
                For i As Integer = 0 To poly1.NumberOfVertices - 1
                    ReDim Preserve arrayPoint1(i)
                    arrayPoint1(i) = poly1.GetPoint2dAt(i)
                Next i

                Dim acSSPr2 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию 2: ")
                If Not (acSSPr2.Status = PromptStatus.OK) Then
                    Exit Sub
                End If

                Dim poly2 As Polyline = acTrans.GetObject(acSSPr2.ObjectId, OpenMode.ForWrite)
                Dim arrayPoint2 As Point2d() = {}
                For i As Integer = 0 To poly2.NumberOfVertices - 1
                    ReDim Preserve arrayPoint2(i)
                    arrayPoint2(i) = poly2.GetPoint2dAt(i)
                Next i

                Dim arrayResult As Point2d() = GetIntersectedPolygon(arrayPoint1, arrayPoint2)

                If IsArray(arrayResult) = True Then
                    Dim newPoly As Polyline = New Polyline()
                    newPoly.SetDatabaseDefaults()
                    For i As Integer = 1 To arrayResult.Length - 1
                        Dim pt As Point2d = arrayResult(i)
                        newPoly.AddVertexAt(i, pt, 0, 0, 0)
                    Next i
                    acBlkTblRec.AppendEntity(newPoly)
                    acTrans.AddNewlyCreatedDBObject(newPoly, True)
                End If
                acTrans.Commit()
            End Using

        Catch ex As System.Exception
        End Try

    End Sub

End Class
