Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class EditVertexFeatureLineToCoGoPoint
    'функция для поиска точки с определенными координатами
    Public Shared Function FindCogoPoint(ByVal PointCoord As Point2d, ByVal NumPoint As Integer(), Optional ByVal dopusk As Double = 0) As CogoPoint
        FindCogoPoint = Nothing
        Dim Lenght As Double = 10000000
        Dim PointId As ObjectId = Nothing
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
                        FindCogoPoint = PointCOGO1
                        Exit Function
                    ElseIf Dist < Lenght Then
                        Lenght = Dist
                        FindCogoPoint = PointCOGO1
                    End If
                Next
                If Lenght > dopusk Then
                    FindCogoPoint = Nothing
                End If
            Catch
                FindCogoPoint = Nothing
            End Try
            tr.Commit()
            locker.Dispose()
        End Using
    End Function

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
                                Dim PointColl As Point3dCollection = New Point3dCollection
                                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                                PointColl = FeatLine.GetPoints(ftype)
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
                                                FeatLine.SetPointElevation(i, elevPoint)
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

    <CommandMethod("InterpolFeatureLineVertex2")>
    Public Sub InterpolFeatureLineVertex2()
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
                                Dim PointColl As Point3dCollection = New Point3dCollection
                                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                                PointColl = FeatLine.GetPoints(ftype)
                                If PointColl.Count > 1 Then
                                    Dim ArrayPoint As Double(,) = Nothing
                                    Dim CountArrayPoint As Integer = 0
                                    For Each pt As Point3d In PointColl
                                        'ищем cogo точку в узле
                                        If pt.Z <> 0 Then
                                            ReDim Preserve ArrayPoint(3, CountArrayPoint)
                                            ArrayPoint(0, CountArrayPoint) = pt.X
                                            ArrayPoint(1, CountArrayPoint) = pt.Y
                                            ArrayPoint(2, CountArrayPoint) = pt.Z
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
                                                FeatLine.SetPointElevation(i, elevPoint)
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
End Class

