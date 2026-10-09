Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class FuncBredge
    'коррекция положения балки на допуск по зазору
    Public Shared Function FuncCorrBalka(ByRef lineShortBearm As Line, ByRef pointEndSectionColl As Point3dCollection, ByVal heightBearm As Double, ByVal widthDownBearm As Double, ByVal widthUpBearm As Double, ByVal startDistToPointPr As Double, ByVal endDistToPointPr As Double, ByVal zazor As Double, Optional ByRef pointStartSectionColl As Point3dCollection = Nothing) As Boolean
        FuncCorrBalka = False
        '1. Удлинняем балку
        Dim lineBearm As Line = lineShortBearm.Clone()
        Dim boolExtBearm As Boolean = FuncExtendBearm(lineBearm, startDistToPointPr, endDistToPointPr)
        '1.делаем смещение балки вверх
        'находим угол наклона балки
        Dim i As Double = (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z) / lineBearm.Length
        Dim wAngle As Double = Math.Asin(i)
        Dim insCoord As Point2d = New Point2d(0, 0)
        Dim deltaXZ As Point2d = MathFunction.FuncReturnPolarCoordinates2(insCoord, wAngle + Math.PI / 2, heightBearm)
        'получаем новые координаты верха балки
        Dim startCoordUpBearm As Point2d = MathFunction.FuncReturnPolarCoordinates2(lineBearm.StartPoint.Convert2d(New Plane), lineBearm.Angle, deltaXZ.X)
        Dim endCoordUpBearm As Point2d = MathFunction.FuncReturnPolarCoordinates2(lineBearm.EndPoint.Convert2d(New Plane), lineBearm.Angle, deltaXZ.X)
        ''создаем верх балки
        Dim startPointUpBalkaPr As Point3d = New Point3d(startCoordUpBearm.X, startCoordUpBearm.Y, lineBearm.StartPoint.Z + deltaXZ.Y)
        Dim endPointUpBalkaPr As Point3d = New Point3d(endCoordUpBearm.X, endCoordUpBearm.Y, lineBearm.EndPoint.Z + deltaXZ.Y)
        Dim lineUpBalka As Line = New Line(startPointUpBalkaPr, endPointUpBalkaPr)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'распаралеливаем линии
        Dim deltaZazor As Double = 0
        Dim pointStart1 As Point3d = Nothing
        Dim pointStart2 As Point3d = Nothing
        Dim pointStart3 As Point3d = Nothing
        Dim pointStart4 As Point3d = Nothing
        Dim pointEnd1 As Point3d = Nothing
        Dim pointEnd2 As Point3d = Nothing
        Dim pointEnd3 As Point3d = Nothing
        Dim pointEnd4 As Point3d = Nothing
        Dim H As Double = lineBearm.EndPoint.Z - lineBearm.StartPoint.Z
        'считаем зазор по низу балки
        '====================================================================================================================
        'смещение низ лево
        Dim dbCollection1 As DBObjectCollection = lineBearm.GetOffsetCurves(widthDownBearm / 2)
        If dbCollection1.Count = 1 Then
            Dim tempLine As Line = dbCollection1.Item(0)
            Dim pointCollection1 As Point3dCollection = New Point3dCollection()
            Dim lineSection As Line = New Line(pointEndSectionColl.Item(0), pointEndSectionColl.Item(1))
            lineSection.IntersectWith(tempLine, Intersect.ExtendBoth, New Plane(), pointCollection1, 1, 1)
            If pointCollection1.Count = 1 Then
                'ищем угол
                Dim angle As Double = pointCollection1.Item(0).Convert2d(New Plane).GetVectorTo(tempLine.StartPoint.Convert2d(New Plane)).Angle
                Dim dist As Double = tempLine.StartPoint.Convert2d(New Plane).GetDistanceTo(pointCollection1.Item(0).Convert2d(New Plane))
                Dim dAngle As Double = Math.Abs(angle - tempLine.Angle)
                If dAngle > 2 Then
                    Dim tempd As Double = dist + zazor
                    If tempd > deltaZazor Then
                        deltaZazor = tempd
                    End If
                Else
                    If dist < zazor Then
                        Dim tempd As Double = zazor - dist
                        If tempd > deltaZazor Then
                            deltaZazor = tempd
                        End If
                    End If
                End If
            End If
            pointStart1 = tempLine.StartPoint
            pointEnd1 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'низ право
        Dim dbCollection2 As DBObjectCollection = lineBearm.GetOffsetCurves(-1 * widthDownBearm / 2)
        If dbCollection2.Count = 1 Then
            Dim tempLine As Line = dbCollection2.Item(0)
            Dim pointCollection2 As Point3dCollection = New Point3dCollection()
            Dim lineSection As Line = New Line(pointEndSectionColl.Item(0), pointEndSectionColl.Item(1))
            lineSection.IntersectWith(tempLine, Intersect.ExtendBoth, New Plane(), pointCollection2, 1, 1)
            If pointCollection2.Count = 1 Then
                Dim angle As Double = pointCollection2.Item(0).Convert2d(New Plane).GetVectorTo(tempLine.StartPoint.Convert2d(New Plane)).Angle
                Dim dist As Double = tempLine.StartPoint.Convert2d(New Plane).GetDistanceTo(pointCollection2.Item(0).Convert2d(New Plane))
                Dim dAngle As Double = Math.Abs(angle - tempLine.Angle)
                If dAngle > 2 Then
                    Dim tempd As Double = dist + zazor
                    If tempd > deltaZazor Then
                        deltaZazor = tempd
                    End If
                Else
                    If dist < zazor Then
                        Dim tempd As Double = zazor - dist
                        If tempd > deltaZazor Then
                            deltaZazor = tempd
                        End If
                    End If
                End If
            End If
            pointStart2 = tempLine.StartPoint
            pointEnd2 = tempLine.EndPoint
        End If
        'считаем зазор по верху балки
        '====================================================================================================================
        Dim dbCollection3 As DBObjectCollection = lineUpBalka.GetOffsetCurves(widthUpBearm / 2)
        If dbCollection3.Count = 1 Then
            Dim tempLine As Line = dbCollection3.Item(0)
            Dim pointCollection3 As Point3dCollection = New Point3dCollection()
            Dim lineSection As Line = New Line(pointEndSectionColl.Item(2), pointEndSectionColl.Item(3))
            lineSection.IntersectWith(tempLine, Intersect.ExtendBoth, New Plane(), pointCollection3, 1, 1)
            If pointCollection3.Count = 1 Then
                Dim angle As Double = pointCollection3.Item(0).Convert2d(New Plane).GetVectorTo(tempLine.StartPoint.Convert2d(New Plane)).Angle
                Dim dist As Double = tempLine.StartPoint.Convert2d(New Plane).GetDistanceTo(pointCollection3.Item(0).Convert2d(New Plane))
                Dim dAngle As Double = Math.Abs(angle - tempLine.Angle)
                If dAngle > 2 Then
                    Dim tempd As Double = dist + zazor
                    If tempd > deltaZazor Then
                        deltaZazor = tempd
                    End If
                Else
                    If dist < zazor Then
                        Dim tempd As Double = zazor - dist
                        If tempd > deltaZazor Then
                            deltaZazor = tempd
                        End If
                    End If
                End If
            End If
            pointStart3 = tempLine.StartPoint
            pointEnd3 = tempLine.EndPoint
        End If
        '====================================================================================================================
        Dim dbCollection4 As DBObjectCollection = lineUpBalka.GetOffsetCurves(-1 * widthUpBearm / 2)
        If dbCollection4.Count = 1 Then
            Dim tempLine As Line = dbCollection4.Item(0)
            Dim pointCollection4 As Point3dCollection = New Point3dCollection()
            Dim lineSection As Line = New Line(pointEndSectionColl.Item(2), pointEndSectionColl.Item(3))
            lineSection.IntersectWith(tempLine, Intersect.ExtendBoth, New Plane(), pointCollection4, 1, 1)
            If pointCollection4.Count = 1 Then
                Dim angle As Double = pointCollection4.Item(0).Convert2d(New Plane).GetVectorTo(tempLine.StartPoint.Convert2d(New Plane)).Angle
                Dim dist As Double = tempLine.StartPoint.Convert2d(New Plane).GetDistanceTo(pointCollection4.Item(0).Convert2d(New Plane))
                Dim dAngle As Double = Math.Abs(angle - tempLine.Angle)
                If dAngle > 2 Then
                    Dim tempd As Double = dist + zazor
                    If tempd > deltaZazor Then
                        deltaZazor = tempd
                    End If
                Else
                    If dist < zazor Then
                        Dim tempd As Double = zazor - dist
                        If tempd > deltaZazor Then
                            deltaZazor = tempd
                        End If
                    End If
                End If
            End If
            pointStart4 = tempLine.StartPoint
            pointEnd4 = tempLine.EndPoint
        End If
        'поправка в линию
        If deltaZazor > 0.001 Then
            Dim boolMoveBearm As Boolean = FuncBredge.FuncMoveBearm(lineShortBearm, deltaZazor)
            Return False
        Else
            pointEndSectionColl = New Point3dCollection
            pointEndSectionColl.Add(pointEnd1)
            pointEndSectionColl.Add(pointEnd2)
            pointEndSectionColl.Add(pointEnd3)
            pointEndSectionColl.Add(pointEnd4)
            pointStartSectionColl = New Point3dCollection()
            pointStartSectionColl.Add(pointStart1)
            pointStartSectionColl.Add(pointStart2)
            pointStartSectionColl.Add(pointStart3)
            pointStartSectionColl.Add(pointStart4)
            Return True
        End If
    End Function
    'коррекция горизонтального расстояния за угол наклона линии
    Public Shared Function FuncFindPointBearm(ByVal axisPline As Polyline, ByVal startPoint As Point3d, ByVal radius As Double, ByVal surf As TinSurface, ByVal heightBearm As Double, ByVal dEarth As Double) As Point3d
        Dim tempStartPoint As Point3d = axisPline.GetClosestPointTo(startPoint, True)
        tempStartPoint = New Point3d(tempStartPoint.X, tempStartPoint.Y, startPoint.Z)
        'находим расстояние на полилинии
        Dim startdist As Double = axisPline.GetDistAtPoint(tempStartPoint)
        Dim findDist As Double = radius
        For i As Integer = 0 To 1000
            Dim delta As Double = startdist + findDist
            Dim tempPoint As Point3d = axisPline.GetPointAtDist(delta)
            Dim elev As Double = startPoint.Z
            Try
                elev = surf.FindElevationAtXY(tempPoint.X, tempPoint.Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            elev = elev - heightBearm - dEarth
            tempPoint = New Point3d(tempPoint.X, tempPoint.Y, elev)
            Dim tempDist As Double = startPoint.GetVectorTo(tempPoint).Length
            Dim dLenght As Double = radius - tempDist
            If Math.Abs(dLenght) < 0.001 Then
                Return tempPoint
            Else
                findDist = findDist + dLenght
            End If
        Next
    End Function
    'функция удлинняет балку
    Public Shared Function FuncExtendBearm(ByRef lineBearm As Line, ByVal startLenght As Double, ByVal endLenght As Double, Optional round As Integer = 3) As Boolean
        FuncExtendBearm = False
        Try
            Dim L As Double = lineBearm.Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = lineBearm.StartPoint.X + k1 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y1 As Double = lineBearm.StartPoint.Y + k1 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z1 As Double = lineBearm.StartPoint.Z + k1 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim startPoint As Point3d = New Point3d(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = lineBearm.EndPoint.X + k2 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y2 As Double = lineBearm.EndPoint.Y + k2 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z2 As Double = lineBearm.EndPoint.Z + k2 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim endPoint As Point3d = New Point3d(X2, Y2, z2)
            lineBearm.StartPoint = startPoint
            lineBearm.EndPoint = endPoint
            Return True
        Catch ex As Exception
        End Try
    End Function
    'функция делает перенос оси балки
    Public Shared Function FuncMoveBearm(ByRef lineBearm As Line, ByVal deltaLenght As Double, Optional round As Integer = 3) As Boolean
        FuncMoveBearm = False
        Try
            Dim L As Double = lineBearm.Length
            Dim k As Double = deltaLenght / L
            Dim X1 As Double = lineBearm.StartPoint.X + k * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y1 As Double = lineBearm.StartPoint.Y + k * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z1 As Double = lineBearm.StartPoint.Z + k * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim startPoint As Point3d = New Point3d(X1, Y1, z1)

            Dim X2 As Double = lineBearm.EndPoint.X + k * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y2 As Double = lineBearm.EndPoint.Y + k * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z2 As Double = lineBearm.EndPoint.Z + k * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim endPoint As Point3d = New Point3d(X2, Y2, z2)

            lineBearm.StartPoint = startPoint
            lineBearm.EndPoint = endPoint
            Return True
        Catch ex As Exception
        End Try
    End Function
End Class
