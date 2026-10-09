Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class EditlineVertexToCoGoPoint
    '=============================================================================================
    'Уточнение вершин линии
    <CommandMethod("EditlineVertexToCoGoPoint")>
    Public Sub EditlineVertexToCoGoPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim strDWGName As String = acDoc.Name
        Try
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try

        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim PointsCogoCollection As CogoPointCollection = Nothing
        Dim ArrayPoint As Integer() = Nothing
        Dim countPline As Integer = 0
        Dim countErrPline As Integer = 0
        Dim Point1 As Point3d = Nothing
        Dim Point2 As Point3d = Nothing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim resPoint As PromptPointResult = Nothing
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'укажите 1 точку
            Dim optPoint As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние от объекта до COGO точки (первая точка): ")
            resPoint = ed.GetPoint(optPoint)
            If resPoint.Status = PromptStatus.OK Then
                Point1 = resPoint.Value
            Else
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'укажите 2 точку
            Dim optPoint1 As PromptPointOptions = New PromptPointOptions(vbLf & "Укажите максимальное расстояние от объекта до COGO точки (вторая точка): ")
            resPoint = ed.GetPoint(optPoint1)
            If resPoint.Status = PromptStatus.OK Then
                Point2 = resPoint.Value
            Else
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'вычисляем радиус поиска точки
            Dim Lenght As Double = Math.Sqrt((Point1.X - Point2.X) ^ 2 + (Point1.Y - Point2.Y) ^ 2)
            Lenght = Math.Round(Lenght, 2)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объекты
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Try
                            If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = obj
                                If Pline.IsWriteEnabled = False Then
                                    Pline.UpgradeOpen()
                                End If
                                'находим cogo точки в габаритном контейнере
                                Dim CountVertex As Integer = FuncSurvey.FuncFindNumberCoGoPointToBoundExtents(Pline, ArrayPoint)
                                'начинаем перебор вершин полилинии
                                Dim vn As Integer = Pline.NumberOfVertices
                                For i As Integer = 0 To vn - 1
                                    Dim HPoint As Double = 0
                                    Dim PointCogoVert As CogoPoint = Nothing
                                    Dim Point As Point3d = Pline.GetPoint3dAt(i)
                                    If IsArray(ArrayPoint) = True Then
                                        PointCogoVert = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                    End If
                                    Try
                                        Dim pt As Point2d = New Point2d(PointCogoVert.Easting, PointCogoVert.Northing)
                                        If PointCogoVert.Easting <> Point.X Or PointCogoVert.Northing <> Point.Y Then
                                            Pline.SetPointAt(i, pt)
                                            Point = Pline.GetPoint3dAt(i)
                                        End If
                                    Catch
                                    End Try
                                Next i
                                '2d полилиния
                            ElseIf TypeOf obj Is Polyline2d Then
                                Dim Pline As Polyline2d = obj
                                If Pline.IsWriteEnabled = False Then
                                    Pline.UpgradeOpen()
                                End If
                                'создаем габаритный контейнер полилинии для фильтрации точек
                                Dim CountVertex As Integer = FuncSurvey.FuncFindNumberCoGoPointToBoundExtents(Pline, ArrayPoint)
                                'начинаем перебор вершин полилинии
                                For Each VertID As ObjectId In Pline
                                    Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForWrite)
                                    Dim HPoint As Double = 0
                                    Dim PointCogoVert As CogoPoint = Nothing
                                    Dim Point As Point3d = New Point3d(v2d.Position.X, v2d.Position.Y, 0)
                                    If IsArray(ArrayPoint) = True Then
                                        PointCogoVert = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                    End If
                                    Try
                                        Dim pt As Point3d = New Point3d(PointCogoVert.Easting, PointCogoVert.Northing, 0)
                                        If PointCogoVert.Easting <> Point.X Or PointCogoVert.Northing <> Point.Y Then
                                            v2d.Position = pt
                                        End If
                                    Catch
                                    End Try
                                Next
                                '3d полилиния
                            ElseIf TypeOf obj Is Polyline3d Then
                                Dim Pline As Polyline3d = obj
                                If Pline.IsWriteEnabled = False Then
                                    Pline.UpgradeOpen()
                                End If
                                'создаем габаритный контейнер полилинии для фильтрации точек
                                Dim CountVertex As Integer = FuncSurvey.FuncFindNumberCoGoPointToBoundExtents(Pline, ArrayPoint)
                                'начинаем перебор вершин полилинии
                                For Each VertID As ObjectId In Pline
                                    Dim v3d As PolylineVertex3d = acTrans.GetObject(VertID, OpenMode.ForWrite)
                                    Dim HPoint As Double = 0
                                    Dim PointCogoVert As CogoPoint = Nothing
                                    Dim Point As Point3d = New Point3d(v3d.Position.X, v3d.Position.Y, 0)
                                    If IsArray(ArrayPoint) = True Then
                                        PointCogoVert = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                    End If
                                    Try
                                        Dim pt As Point3d = New Point3d(PointCogoVert.Easting, PointCogoVert.Northing, PointCogoVert.Elevation)
                                        If PointCogoVert.Easting <> Point.X Or PointCogoVert.Northing <> Point.Y Or PointCogoVert.Elevation <> v3d.Position.Z Then
                                            v3d.Position = pt
                                        End If
                                    Catch
                                    End Try
                                Next
                            ElseIf TypeOf obj Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                'ссылка на объект
                                Dim FLine As FeatureLine = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                Dim NameFeat As String = FLine.Name
                                Dim desk As String = FLine.Description
                                Dim siteID As ObjectId = FLine.SiteId
                                Dim NameStyleFline As String = FLine.StyleName
                                FLine.Erase()
                                Dim ArrayData As String(,) = Nothing
                                'читаем наименования таблиц PS
                                Dim ArrayTBL As String() = Nothing
                                Dim boolFindTdl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(FLine, ArrayTBL)
                                If IsArray(ArrayTBL) = True Then
                                    Dim boolreadPs As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(FLine, ArrayData)
                                End If
                                Dim colPt As Point3dCollection = FLine.GetPoints(FeatureLinePointType.AllPoints)
                                'делаем новую полилинию
                                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                                acPoly.SetDatabaseDefaults()
                                If colPt.Count > 1 Then
                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                                    For i As Integer = 0 To colPt.Count - 1
                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
                                    Next
                                End If
                                acBlkTblRec.AppendEntity(acPoly)
                                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                                'находим cogo точки в габаритном контейнере
                                Dim CountVertex As Integer = FuncSurvey.FuncFindNumberCoGoPointToBoundExtents(acPoly, ArrayPoint)
                                'начинаем перебор вершин полилинии
                                Dim vn As Integer = acPoly.NumberOfVertices
                                For i As Integer = 0 To vn - 1
                                    Dim HPoint As Double = 0
                                    Dim PointCogoVert As CogoPoint = Nothing
                                    Dim Point As Point3d = acPoly.GetPoint3dAt(i)
                                    If IsArray(ArrayPoint) = True Then
                                        PointCogoVert = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                    End If
                                    Try
                                        Dim pt1 As Point2d = New Point2d(PointCogoVert.Easting, PointCogoVert.Northing)
                                        If PointCogoVert.Easting <> Point.X Or PointCogoVert.Northing <> Point.Y Then
                                            acPoly.SetPointAt(i, pt1)
                                            Point = acPoly.GetPoint3dAt(i)
                                        End If
                                    Catch
                                    End Try
                                Next i
                                'ищем поверхность
                                Dim idSurfColl As ObjectIdCollection = CivilDoc.GetSurfaceIds
                                Dim idActiveSurf As ObjectId = Nothing
                                If idSurfColl.Count > 0 Then
                                    idActiveSurf = idSurfColl.Item(0)
                                End If
                                'создаем характерную линию заново
                                '===========================================================================
                                Dim idFeat As ObjectId = Autodesk.Civil.DatabaseServices.FeatureLine.Create(NameFeat, acPoly.Id, siteID)
                                Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(idFeat, OpenMode.ForWrite)
                                Dim pt As Point3dCollection = FeatLine.GetPoints(FeatureLinePointType.ElevationPoint)
                                If pt.Count > 0 Then
                                    For Each pttemp As Point3d In pt
                                        FeatLine.DeleteElevationPoint(pttemp)
                                    Next
                                End If
                                FeatLine.Description = desk
                                FeatLine.StyleName = NameStyleFline
                                'назначаем отметки
                                If IsNothing(idActiveSurf) = False Then
                                    Try
                                        FeatLine.AssignElevationsFromSurface(idActiveSurf, False)
                                    Catch ex As System.ArgumentException
                                    End Try
                                End If
                                If IsArray(ArrayTBL) = True Then
                                    For i As Integer = 0 To ArrayTBL.Length - 1
                                        Dim idPSDict As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTBL(i))
                                        If idPSDict.IsNull = False Or IsNothing(idPSDict) = False Then
                                            Dim boolAddPSDict As Boolean = FuncPropertySet.FuncAddDataSetProperties(FeatLine, idPSDict)
                                        End If
                                    Next i
                                    If IsArray(ArrayData) = True Then
                                        Dim boolrecData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(FeatLine, ArrayData)
                                    End If
                                End If
                            End If
                            countPline += 1
                        Catch ex As System.Exception
                            countErrPline += 1
                        End Try
                    End If
                Next
            End If
            acTrans.Commit()
            locker.Dispose()
        End Using
        MsgBox("Макро выполнено." & vbLf & "Успешно уточнены вершины  " & countPline & " полилинии(й)," & vbLf & "Не удалось преобразовать " & countErrPline & " полилиний.")
    End Sub
End Class

