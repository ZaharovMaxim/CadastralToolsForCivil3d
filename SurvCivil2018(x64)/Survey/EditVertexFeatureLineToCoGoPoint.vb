Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class EditVertexFeatureLineToCoGoPoint
    'функция для поиска точки с определенными координатами
    Public Shared Function getCogoPoint(ByVal PointCoord As Point2d, ByVal NumPoint As Integer(), Optional ByVal dopusk As Double = 0) As CogoPoint
        Dim result As CogoPoint = Nothing
        Dim Lenght As Double = 10000000
        Dim PointId As ObjectId = Nothing
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    For i As Integer = 0 To NumPoint.GetUpperBound(0)
                        PointId = CivilDoc.CogoPoints.GetPointByPointNumber(NumPoint(i))
                        Dim PointCOGO1 As CogoPoint = PointId.GetObject(OpenMode.ForWrite)
                        Dim X1 As Double = PointCOGO1.Easting
                        Dim Y1 As Double = PointCOGO1.Northing
                        Dim t1 As Point2d = New Point2d(PointCoord.X, PointCoord.Y)
                        Dim t2 As Point2d = New Point2d(X1, Y1)
                        Dim Dist As Double = t1.GetVectorTo(t2).Length
                        Dist = Math.Round(Dist, 3)
                        If Dist = 0 Then
                            result = PointCOGO1
                            Exit For
                        ElseIf Dist < Lenght Then
                            Lenght = Dist
                            result = PointCOGO1
                        End If
                    Next
                    If Lenght > dopusk Then
                        result = Nothing
                    End If
                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    result = Nothing
                End Try
            End Using
        End Using
        Return result
    End Function
    'интерполирует высоты вершин между точками COGO
    <CommandMethod("InterpolFeatureLineVertex")>
    Public Sub InterpolFeatureLineVertex()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '==============================================================================
        'открываем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Try
                            If TypeOf acEnt Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                'забираем все точки которые находятся в габаритном контейнере
                                Dim ArrayCOGO As Integer() = Nothing
                                Dim boolFindArrayCogo As Boolean = FuncSurvey.FuncWriteArrayNumberCoGoPointToBoundExtentsObject(acEnt, ArrayCOGO)
                                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                                Dim PointColl As Point3dCollection = FeatLine.GetPoints(ftype)

                                If PointColl.Count > 1 Then
                                    Dim ArrayPoint As Double(,) = Nothing
                                    Dim CountArrayPoint As Integer = 0
                                    For Each pt As Point3d In PointColl
                                        'ищем cogo точку в узле
                                        Dim pointCogo As Autodesk.Civil.DatabaseServices.CogoPoint = FuncSurvey.FuncFindCogoPointByCoordinates(pt, ArrayCOGO)
                                        If IsNothing(pointCogo) = False Then
                                            ReDim Preserve ArrayPoint(3, CountArrayPoint)
                                            ArrayPoint(0, CountArrayPoint) = pt.X
                                            ArrayPoint(1, CountArrayPoint) = pt.Y
                                            ArrayPoint(2, CountArrayPoint) = pointCogo.Elevation
                                            ArrayPoint(3, CountArrayPoint) = 0
                                            CountArrayPoint += 1
                                        Else
                                            ReDim Preserve ArrayPoint(3, CountArrayPoint)
                                            ArrayPoint(0, CountArrayPoint) = pt.X
                                            ArrayPoint(1, CountArrayPoint) = pt.Y
                                            ArrayPoint(2, CountArrayPoint) = 0
                                            If ArrayPoint.GetUpperBound(1) > 0 Then
                                                Dim ptTemp As Point2d = New Point2d(ArrayPoint(0, CountArrayPoint - 1), ArrayPoint(1, CountArrayPoint - 1))
                                                ArrayPoint(3, CountArrayPoint) = pt.Convert2d(New Plane).GetVectorTo(ptTemp).Length
                                            Else
                                                ArrayPoint(3, CountArrayPoint) = 0
                                            End If
                                            CountArrayPoint += 1
                                        End If
                                    Next
                                    '=======================================================================================
                                    'разбираем полученный массив
                                    If IsArray(ArrayPoint) = True Then
                                        If ArrayPoint.GetUpperBound(1) > 0 Then
                                            'берем высоты первой и последней точек от ближайшей точки cogo
                                            '===============================================================
                                            If ArrayPoint(2, 0) = 0 Then
                                                ArrayPoint(2, 0) = PointColl.Item(0).Z
                                            End If
                                            If ArrayPoint(2, ArrayPoint.GetUpperBound(1)) = 0 Then
                                                ArrayPoint(2, ArrayPoint.GetUpperBound(1)) = PointColl.Item(PointColl.Count - 1).Z
                                            End If
                                            '=====================================================================
                                            'заполняем высоты для остальных точек
                                            Dim startElevPoint As Double = 0
                                            Dim endElevPoint As Double = 0
                                            Dim dist As Double = 0
                                            For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                                                'высота первой точки
                                                Dim fPoint As Double = ArrayPoint(2, i)
                                                If fPoint = 0 Then
                                                    startElevPoint = ArrayPoint(2, i - 1)
                                                    dist = ArrayPoint(3, i)

                                                    For j As Integer = i + 1 To ArrayPoint.GetUpperBound(1)
                                                        Dim ftempPoint As Double = ArrayPoint(2, j)
                                                        If ftempPoint = 0 Then
                                                            dist = dist + ArrayPoint(3, j)
                                                        Else
                                                            endElevPoint = ArrayPoint(2, j)
                                                            Dim temppt1 As Point2d = New Point2d(ArrayPoint(0, j - 1), ArrayPoint(1, j - 1))
                                                            Dim temppt2 As Point2d = New Point2d(ArrayPoint(0, j), ArrayPoint(1, j))
                                                            dist = dist + temppt1.GetVectorTo(temppt2).Length
                                                            Exit For
                                                        End If
                                                    Next j
                                                    '===========================================
                                                    'делаем расчет
                                                    Dim deltaH As Double = endElevPoint - startElevPoint
                                                    Dim ddist As Double = ArrayPoint(3, i) / dist
                                                    Dim rez1 As Double = ddist * deltaH
                                                    ArrayPoint(2, i) = startElevPoint + rez1
                                                End If
                                            Next i
                                            'уточняем вершины
                                            For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                                                Dim elevPoint As Double = ArrayPoint(2, i)
                                                If i < FeatLine.PointsCount Then
                                                    FeatLine.SetPointElevation(i, elevPoint)
                                                End If
                                            Next i
                                        End If
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            acTrans.Commit()
            locker.Dispose()
        End Using
    End Sub
    'уточняет отметки характерной линии только в тех вершинах в которых есть точка COGO
    <CommandMethod("EditFeatureLine")>
    Public Sub EditFeatureLine()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '==============================================================================
        'открываем транзакцию
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Try
                            If TypeOf acEnt Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                'забираем все точки которые находятся в габаритном контейнере
                                Dim ArrayCOGO As Integer() = Nothing
                                Dim boolFindArrayCogo As Boolean = FuncSurvey.FuncWriteArrayNumberCoGoPointToBoundExtentsObject(acEnt, ArrayCOGO)
                                If IsArray(ArrayCOGO) = False Then
                                    Exit Sub
                                End If
                                Dim PointColl As Point3dCollection = New Point3dCollection
                                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                                PointColl = FeatLine.GetPoints(ftype)
                                If PointColl.Count > 1 Then
                                    Dim ArrayPoint As Double(,) = Nothing
                                    Dim CountArrayPoint As Integer = 0
                                    For Each pt As Point3d In PointColl
                                        'ищем cogo точку в узле
                                        Dim pt2d As Point2d = pt.Convert2d(New Plane)
                                        Dim pointCogo As Autodesk.Civil.DatabaseServices.CogoPoint = FuncSurvey.FuncFindCogoPointByCoordinates(pt, ArrayCOGO)
                                        If IsNothing(pointCogo) = False Then
                                            Dim Elev As Double = pointCogo.Elevation
                                            Try
                                                FeatLine.SetPointElevation(CountArrayPoint, Elev)
                                            Catch ex As ArgumentOutOfRangeException
                                            End Try
                                        End If
                                        CountArrayPoint += 1
                                    Next
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            acTrans.Commit()
            locker.Dispose()
        End Using
    End Sub
    'назначает отметки вершинам характерной линии которые име.т высоту 0
    <CommandMethod("InterpolVertexFeatureLineByLength")>
    Public Sub InterpolVertexFeatureLineByLength()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '==============================================================================
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Try
                            If TypeOf acEnt Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                Dim FeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                'читаем ВСЕ точки, чтобы индексы совпадали с SetPointElevation(i)
                                Dim pts As Point3dCollection = FeatLine.GetPoints(Autodesk.Civil.FeatureLinePointType.AllPoints)
                                Dim n As Integer = pts.Count
                                If n > 2 Then
                                    'накопленные плановые длины от вершины 0 до i
                                    Dim cum As Double() = New Double(n - 1) {}
                                    Dim z As Double() = New Double(n - 1) {}
                                    cum(0) = 0.0
                                    z(0) = pts(0).Z
                                    For i As Integer = 1 To n - 1
                                        Dim dx As Double = pts(i).X - pts(i - 1).X
                                        Dim dy As Double = pts(i).Y - pts(i - 1).Y
                                        cum(i) = cum(i - 1) + Math.Sqrt(dx * dx + dy * dy)
                                        z(i) = pts(i).Z
                                    Next
                                    'ищем группы подряд идущих вершин с нулевой отметкой
                                    Dim iStart As Integer = 0
                                    While iStart < n
                                        If z(iStart) <> 0.0 Then
                                            iStart += 1
                                            Continue While
                                        End If
                                        Dim iEnd As Integer = iStart
                                        While iEnd + 1 < n AndAlso z(iEnd + 1) = 0.0
                                            iEnd += 1
                                        End While
                                        'левый якорь — последняя вершина слева с отметкой <> 0
                                        Dim a As Integer = -1
                                        For k As Integer = iStart - 1 To 0 Step -1
                                            If z(k) <> 0.0 Then
                                                a = k
                                                Exit For
                                            End If
                                        Next
                                        'правый якорь — первая вершина справа с отметкой <> 0
                                        Dim b As Integer = -1
                                        For k As Integer = iEnd + 1 To n - 1
                                            If z(k) <> 0.0 Then
                                                b = k
                                                Exit For
                                            End If
                                        Next
                                        'интерполируем только если есть оба якоря
                                        If a >= 0 AndAlso b >= 0 Then
                                            Dim dz As Double = z(b) - z(a)
                                            Dim dl As Double = cum(b) - cum(a)
                                            If dl > 0.0 Then
                                                For k As Integer = iStart To iEnd
                                                    Dim zNew As Double = z(a) + dz * (cum(k) - cum(a)) / dl
                                                    FeatLine.SetPointElevation(k, Math.Round(zNew, 3))
                                                Next
                                            End If
                                        End If
                                        iStart = iEnd + 1
                                    End While
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End If
            acTrans.Commit()
            locker.Dispose()
        End Using
    End Sub
End Class

