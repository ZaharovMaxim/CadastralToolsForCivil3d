Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Public Class MathFunction
    '==================================================================================================================
    'работа с полилинией
    Structure PerVertexData
        Public pt As Autodesk.AutoCAD.Geometry.Point2d
        Public bulge As Double
        Public startWidth As Double
        Public endWidth As Double
    End Structure
    Public Structure Rectangle
        Public Property X As Integer
        Public Property Y As Integer
        Public Property Width As Integer
        Public Property Height As Integer
        Public Sub New(x As Integer, y As Integer, width As Integer, height As Integer)
            Me.X = x
            Me.Y = y
            Me.Width = width
            Me.Height = height
        End Sub
    End Structure
    ' Типы пересечений
    Public Enum IntersectionType
        NoIntersection
        Intersection
        OneInsideAnother
    End Enum
    ' Структура прямоугольника
    Public Structure Rectangle2
        Public X1 As Integer ' Левый нижний угол
        Public Y1 As Integer
        Public X2 As Integer ' Правый верхний угол
        Public Y2 As Integer

        Public Sub New(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
            Me.X1 = x1
            Me.Y1 = y1
            Me.X2 = x2
            Me.Y2 = y2
        End Sub
    End Structure

    '===================================================================================================
    'функция округляет чило вниз 3.7777777 - 7.77
    Public Shared Function FuncTrimDigitFloo(ByVal a As Double, ByVal n As Integer) As Double
        Return Math.Floor(a * 10 ^ n) / 10 ^ n
    End Function
    '===================================================================================================
    'функция округляет чило вверх 3.7777777 - 7.78
    Public Shared Function FuncTrimDigitCall(ByVal a As Double, ByVal n As Integer) As Double
        Return Math.Ceiling(a * 10 ^ n) / 10 ^ n
    End Function
    'массивы
    '=====================================================================================
    'Функция сортирует числовой двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortDblArray(ByRef GeoArray(,) As Double, ByVal ind As Integer) As Boolean
        Dim Temp1 As Double() = Nothing
        Dim Temp2 As Double() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        If IsArray(GeoArray) = False Then
            FuncSortDblArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortDblArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(1) Then
            FuncSortDblArray = False
            Exit Function
        End If
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st1 = GeoArray(ind, i)
                    st2 = Temp2(ind)
                    If st2 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k
                    End If
                Next j
            Next i
        Catch
            FuncSortDblArray = False
        End Try
        FuncSortDblArray = True
    End Function
    '=====================================================================================
    'Функция сортирует строковый двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortStrArray(ByRef GeoArray(,) As String, ByVal ind As Integer) As Boolean
        FuncSortStrArray = False
        If IsArray(GeoArray) = False Then
            FuncSortStrArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortStrArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(0) Then
            FuncSortStrArray = False
            Exit Function
        End If
        Dim Temp1 As String() = Nothing
        Dim Temp2 As String() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        Dim st3 As String = Nothing
        Dim st4 As String = Nothing
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    If IsNumeric(GeoArray(ind, i)) And IsNumeric(Temp2(ind)) Then
                        st1 = Val(GeoArray(ind, i))
                        st2 = Val(Temp2(ind))
                        If st2 < st1 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    Else
                        st3 = GeoArray(ind, i)
                        st4 = Temp2(ind)
                        If st4 < st3 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    End If
                Next j
            Next i
        Catch
            FuncSortStrArray = False
        End Try
        FuncSortStrArray = True
    End Function
    '======================================================================================
    'функция ищет координаты точки в двумерном массиве
    Public Shared Function FuncFindCoordinatesPointToArray(ByVal ArrayPoint As Double(,), ByVal pt1 As Point2d) As Boolean
        FuncFindCoordinatesPointToArray = False
        Try
            For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                Dim x As Double = ArrayPoint(0, i)
                Dim y As Double = ArrayPoint(1, i)
                If x = pt1.X And y = pt1.Y Then FuncFindCoordinatesPointToArray = True
            Next i
        Catch ex As Exception
        End Try
    End Function
    '======================================================================================
    'функция ищет значение в двумерном массиве
    Public Shared Function FuncFindValueToArray(ByVal ArrayData As String(), ByVal value As String) As Boolean
        FuncFindValueToArray = False
        If IsNothing(value) Then Return False
        If IsArray(ArrayData) = False Then Return False
        Try
            If IsArray(ArrayData) = True Then
                For i As Integer = 0 To ArrayData.Length - 1
                    Dim userValue As String = ArrayData(i)
                    If IsNothing(userValue) = False Then
                        If userValue.Trim Like value.Trim Then
                            Return True
                        End If
                    End If
                Next i
            End If
        Catch ex As Exception
        End Try
    End Function

    '======================================================================================
    'функция ищет значение в двумерном массиве
    Public Shared Function FuncFindValueToArray2d(ByVal ArrayData As String(,), ByVal nameField As String, Optional indexField As Integer = 0, Optional indexVal As Integer = 1) As String
        FuncFindValueToArray2d = ""
        If IsNothing(nameField) Then Return ""
        If IsArray(ArrayData) = False Then Return ""
        If ArrayData.GetUpperBound(0) < indexField Then Return ""
        If ArrayData.GetUpperBound(0) < indexVal Then Return ""
        Try
            If IsArray(ArrayData) = True Then
                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                    Dim userField As String = ArrayData(indexField, i)
                    If IsNothing(userField) = False Then
                        If userField.Trim Like nameField.Trim Then
                            Dim valStr As String = ArrayData(indexVal, i)
                            If IsNothing(valStr) = True Then valStr = ""
                            Return valStr
                        End If
                    End If
                Next i
            End If
        Catch ex As Exception
        End Try
    End Function
    '======================================================================================
    'функция ищет значение в двумерном массиве
    Public Shared Function FuncFindIndexValueToArray2d(ByVal ArrayData As String(,), ByVal nameField As String, Optional indexField As Integer = 0) As Integer
        FuncFindIndexValueToArray2d = -1
        If IsNothing(nameField) Then Return -1
        If IsArray(ArrayData) = False Then Return -1
        If ArrayData.GetUpperBound(0) < indexField Then Return -1
        Try
            If IsArray(ArrayData) = True Then
                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                    Dim userField As String = ArrayData(indexField, i)
                    If IsNothing(userField) = False Then
                        If userField.Trim Like nameField.Trim Then
                            Return i
                        End If
                    End If
                Next i
            End If
        Catch ex As Exception
        End Try
    End Function



    '===============================================================================================
    'функция вычисляет точку пересечения двух отрезков
    Public Shared Function FuncIntersectPolyline(ByVal Pt1 As Point2d, ByVal Pt2 As Point2d, ByVal Pt3 As Point2d, ByVal pt4 As Point2d, ByRef InPoint As Point2d) As Boolean
        FuncIntersectPolyline = False
        Dim x1 As Double = Pt1.X
        Dim y1 As Double = Pt1.Y

        Dim x2 As Double = Pt2.X
        Dim y2 As Double = Pt2.Y

        Dim x3 As Double = Pt3.X
        Dim y3 As Double = Pt3.Y

        Dim x4 As Double = pt4.X
        Dim y4 As Double = pt4.Y
        Try
            Dim d As Double = (Pt1.X - Pt2.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (pt4.X - Pt3.X)
            Dim da As Double = (Pt1.X - Pt3.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt3.Y) * (pt4.X - Pt3.X)
            Dim db As Double = (Pt1.X - Pt2.X) * (Pt1.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (Pt1.X - Pt3.X)

            Dim ta As Double = da / d
            Dim tb As Double = db / d

            If (ta >= 0 And ta <= 1 And tb >= 0 And tb <= 1) Then
                Dim dx As Double = Pt1.X + ta * (Pt2.X - Pt1.X)
                Dim dy As Double = Pt1.Y + ta * (Pt2.Y - Pt1.Y)
                InPoint = New Point2d(dx, dy)
                FuncIntersectPolyline = True
            End If
        Catch ex As Exception
        End Try
    End Function
    '====================================================================================
    'функция проверяет лежит ли точка внутри полигона (Новый вариант)
    Public Shared Function FuncCrossPointPolyline(ByVal acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline, ByVal p As Point2d) As Boolean
        Dim p1, p2 As Point2d
        Dim inside As Boolean = False
        Dim poly As Double(,) = Nothing
        Dim countPoly As Integer = 0
        If acPoly.NumberOfVertices < 3 Then
            Return inside
        Else
            For i As Integer = 0 To acPoly.NumberOfVertices - 1
                ReDim Preserve poly(1, i)
                poly(0, i) = acPoly.GetPoint2dAt(i).X
                poly(1, i) = acPoly.GetPoint2dAt(i).Y
                countPoly += 1
            Next i
            If poly(0, 0) <> poly(0, poly.GetUpperBound(1)) And poly(1, 0) <> poly(1, poly.GetUpperBound(1)) Then
                ReDim Preserve poly(1, countPoly)
                poly(0, countPoly) = acPoly.GetPoint2dAt(0).X
                poly(1, countPoly) = acPoly.GetPoint2dAt(0).Y
            End If
        End If
        '=========================================================================================
        Dim XOld As Double = poly(0, poly.GetUpperBound(1))
        Dim YOld As Double = poly(1, poly.GetUpperBound(1))
        Dim oldPoint As Point2d = New Point2d(XOld, YOld)
        For i As Integer = 0 To poly.GetUpperBound(1)
            Dim newPoint As Point2d = New Point2d(poly(0, i), poly(1, i))
            If newPoint.X > oldPoint.X Then
                p1 = oldPoint
                p2 = newPoint
            Else
                p1 = newPoint
                p2 = oldPoint
            End If
            If (newPoint.X < p.X) = (p.X <= oldPoint.X) AndAlso (p.Y - CLng(p1.Y)) * (p2.X - p1.X) < (p2.Y - CLng(p1.Y)) * (p.X - p1.X) Then
                inside = Not inside
            End If
            oldPoint = newPoint
        Next
        Return inside
    End Function
    '====================================================================================
    'функция проверяет лежит ли точка внутри полигона (нужны библиотеки Map3D)
    Public Shared Function FuncCrossPointPolygon(ByVal acEntObj As Autodesk.AutoCAD.DatabaseServices.Entity, ByVal Pt As Point3d) As Boolean
        FuncCrossPointPolygon = False
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Try
            Dim mp As MPolygon = New MPolygon()
            If Not IsDBNull(acEntObj) Then
                If TypeOf acEntObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim Poly As Autodesk.AutoCAD.DatabaseServices.Polyline = acEntObj
                    mp.AppendLoopFromBoundary(Poly, True, Tolerance.Global.EqualPoint)
                ElseIf TypeOf acEntObj Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim Poly As Polyline2d = acEntObj
                    mp.AppendLoopFromBoundary(Poly, True, Tolerance.Global.EqualPoint)
                ElseIf TypeOf acEntObj Is Autodesk.AutoCAD.DatabaseServices.Circle Then
                    Dim Poly As Circle = acEntObj
                    mp.AppendLoopFromBoundary(Poly, True, Tolerance.Global.EqualPoint)
                End If
            End If
            If (mp.IsPointInsideMPolygon(Pt, Tolerance.Global.EqualPoint).Count > 0) Then
                FuncCrossPointPolygon = True
            End If
        Catch
            Exit Function
        End Try
    End Function
    '==================================================================================================================================================================================================
    'функция ищет координаты середины дуги
    Public Shared Function FuncFindMiddleArc(ByVal ptStart As Point2d, ByVal ptEnd As Point2d, ByVal ptCentre As Point2d) As Point2d
        FuncFindMiddleArc = Nothing
        Try
            'вычисляем дирекционное направление на 1 точку
            Dim VectStart As Vector2d = ptCentre.GetVectorTo(ptStart)
            'вычисляем дирекционное направление на 2 точку
            Dim VectEnd As Vector2d = ptCentre.GetVectorTo(ptEnd)
            'вычисляем среднее значение
            Dim MiddleDirection As Double = (VectEnd.Angle + VectStart.Angle)
            If MiddleDirection > 2 * Math.PI Then
                MiddleDirection = MiddleDirection - Math.PI
            End If
            MiddleDirection = MiddleDirection / 2
            'определяем длину вектора (радиус дуги)
            Dim LArc As Double = VectStart.Length
            'определяем координаты
            Dim X As Double = ptCentre.X + LArc * Math.Cos(MiddleDirection)
            Dim Y As Double = ptCentre.X + LArc * Math.Sin(MiddleDirection)
            FuncFindMiddleArc = New Point2d(X, Y)
        Catch ex As Exception
        End Try
    End Function
    '==================================================================================================================================================================================================
    'функция ищет координаты середины дуги
    Public Shared Function FuncFindCenterObjectByGeometricExtents(ByVal acEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Point3d
        FuncFindCenterObjectByGeometricExtents = New Point3d(-1, -1, -1)
        Try
            Dim extPoint As Extents3d = acEnt.GeometricExtents
            FuncFindCenterObjectByGeometricExtents = New Point3d((extPoint.MinPoint.X + extPoint.MaxPoint.X) / 2, (extPoint.MinPoint.Y + extPoint.MaxPoint.Y) / 2, 0)
        Catch ex As Exception
        End Try
    End Function
    '==================================================================================================================================================================================================
    'функция ищет координаты середины отрезка
    Public Shared Function FuncFindCenterLine(ByVal startPoint As Point2d, ByVal endPoint As Point2d) As Point2d
        FuncFindCenterLine = New Point2d(-1, -1)
        Try
            FuncFindCenterLine = New Point2d((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2)
        Catch ex As Exception
        End Try
    End Function
    '======================================================================================
    'функция ищет центр окружности по координатам 3 точек
    Public Shared Function FuncFindCenterCircleByThreePoint(ByVal pt1 As Point3d, ByVal pt2 As Point3d, ByVal pt3 As Point3d) As Point2d
        FuncFindCenterCircleByThreePoint = Nothing
        Try
            Dim X1 As Double = pt1.X
            Dim Y1 As Double = pt1.Y

            Dim X2 As Double = pt2.X
            Dim Y2 As Double = pt2.Y

            Dim X3 As Double = pt3.X
            Dim Y3 As Double = pt3.Y

            Dim A As Double = X2 - X1
            Dim B As Double = Y2 - Y1
            Dim C As Double = X3 - X1
            Dim D As Double = Y3 - Y1

            Dim E As Double = A * (X1 + X2) + B * (Y1 + Y2)
            Dim F As Double = C * (X1 + X3) + D * (Y1 + Y3)
            Dim G As Double = 2 * (A * (Y3 - Y2) - B * (X3 - X2))
            If G = 0 Then
                Exit Function
            End If

            Dim Cx As Double = Math.Round((D * E - B * F) / G, 3)
            Dim Cy As Double = Math.Round((A * F - C * E) / G, 3)
            FuncFindCenterCircleByThreePoint = New Point2d(Cx, Cy)
        Catch ex As Exception
        End Try
    End Function

    '=================================================================================================
    'определяет справа или слева определяемая точка относительно полилинии (слева - false)
    Public Shared Function FuncIsRightDirection(ByVal pCurv As Curve, ByVal p As Point3d, ByVal vDir As Vector3d) As Boolean
        Dim vNormal As Vector3d = Vector3d.ZAxis
        If pCurv.IsPlanar Then
            Dim plane As Plane = pCurv.GetPlane()
            vNormal = plane.Normal
            p = p.Project(plane, vDir)
        End If
        Dim pNear As Point3d = pCurv.GetClosestPointTo(p, True)
        Dim vSide As Vector3d = p - pNear
        Dim vDeriv As Vector3d = pCurv.GetFirstDerivative(pNear)
        If vNormal.CrossProduct(vDeriv).DotProduct(vSide) < 0.0 Then
            Return True
        Else
            Return False
        End If
    End Function
    '=========================================================================================
    'функция удаляет совпадающие вершины (всех 2д и 3д)
    Public Shared Function FuncKadastrPlineDelDubleVertex(ObjEnt As Autodesk.AutoCAD.DatabaseServices.Entity, Optional ByVal DopuskXY As Double = 0, Optional ByVal DopuskH As Double = 0) As Boolean
        FuncKadastrPlineDelDubleVertex = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If ObjEnt.Id.IsNull = False Or ObjEnt.Id.IsErased = False Then
                'обычная полилиния
                If TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pLine As Autodesk.AutoCAD.DatabaseServices.Polyline = ObjEnt
                    Try
line1:
                        'проверяем на совпадение последней и первой вершины
                        Dim StartPt As Autodesk.AutoCAD.Geometry.Point2d = New Point2d(pLine.StartPoint.X, pLine.StartPoint.Y)
                        Dim EndPt As Autodesk.AutoCAD.Geometry.Point2d = pLine.GetPoint2dAt(pLine.NumberOfVertices - 1)
                        Dim dist As Double = StartPt.GetVectorTo(EndPt).Length
                        If dist <= DopuskXY And pLine.NumberOfVertices > 1 Then
                            pLine.RemoveVertexAt(pLine.NumberOfVertices - 1)
                            pLine.Closed = True
                            GoTo line1
                        End If
                        'делаем проверку для остальных вершин
                        For i As Integer = 0 To pLine.NumberOfVertices - 2
                            Dim seg As Autodesk.AutoCAD.Geometry.LineSegment2d = pLine.GetLineSegment2dAt(i)
                            If seg.Length <= DopuskXY Then
                                pLine.RemoveVertexAt(i + 1)
                                GoTo line1
                            End If
                        Next i
                        FuncKadastrPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineDelDubleVertex = False
                    End Try
                    '2d полилиния
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pLine2d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ObjEnt
                    Dim v2dPrew As Autodesk.AutoCAD.DatabaseServices.Vertex2d = Nothing
                    Dim countVert As Integer = 0
                    Try
line2:
                        For Each vid As ObjectId In pLine2d
                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            'если это первая вершина то берем ее значение и продолжаем цикл
                            If countVert = 0 Then
                                v2dPrew = v2d
                                countVert += 1
                                Continue For
                            End If
                            'сравниваем значение с предудущей вершиной
                            Dim seg As Double = v2d.Position.GetVectorTo(v2dPrew.Position).Length
                            If seg <= DopuskXY Then
                                v2d.Erase(True)
                            Else
                                v2dPrew = v2d
                                countVert += 1
                            End If
                        Next
                        'делаем проверку последней вершины
                        If v2dPrew.Position.X = pLine2d.StartPoint.X And v2dPrew.Position.Y = pLine2d.StartPoint.Y Then
                            v2dPrew.Erase(True)
                            pLine2d.Closed = True
                        End If
                        FuncKadastrPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineDelDubleVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                    Dim pLine3d As Autodesk.AutoCAD.DatabaseServices.Polyline3d = ObjEnt
                    Dim v3dPrew As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = Nothing
                    Dim countVert As Integer = 0
                    Try
line3:
                        For Each vid As ObjectId In pLine3d
                            Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            If countVert = 0 Then
                                v3dPrew = v3d
                                countVert += 1
                                Continue For
                            End If
                            Dim pt1 As Autodesk.AutoCAD.Geometry.Point2d = New Autodesk.AutoCAD.Geometry.Point2d(v3d.Position.X, v3d.Position.Y)
                            Dim pt2 As Autodesk.AutoCAD.Geometry.Point2d = New Autodesk.AutoCAD.Geometry.Point2d(v3dPrew.Position.X, v3dPrew.Position.Y)
                            Dim seg2d As Double = pt1.GetVectorTo(pt2).Length
                            Dim seg3d As Double = Math.Abs(v3d.Position.Z - v3dPrew.Position.Z)
                            If seg2d <= DopuskXY And seg3d <= seg3d Then
                                v3d.Erase()
                                countVert = 0
                                GoTo line3
                            Else
                                v3dPrew = v3d
                                countVert += 1
                            End If
                        Next
                        If v3dPrew.Position.X = pLine3d.StartPoint.X And v3dPrew.Position.Y = pLine3d.StartPoint.Y And v3dPrew.Position.Z = pLine3d.StartPoint.Z Then
                            v3dPrew.Erase(True)
                            pLine3d.Closed = True
                        End If
                        FuncKadastrPlineDelDubleVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineDelDubleVertex = False
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
    End Function
    '=========================================================================================
    'функция округляет вершины полилиний (всех 2д и 3д)
    Public Shared Function FuncKadastrPlineRoundVertex(ObjEnt As Autodesk.AutoCAD.DatabaseServices.Entity, Optional ByVal RoundZnXY As Integer = 2, Optional ByVal RoundZnH As Integer = 2) As Boolean
        FuncKadastrPlineRoundVertex = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If ObjEnt.Id.IsNull = False Or ObjEnt.Id.IsErased = False Then
                If TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pLine As Autodesk.AutoCAD.DatabaseServices.Polyline = ObjEnt
                    Try
                        For i As Integer = 0 To pLine.NumberOfVertices - 1
                            Dim x As Double = pLine.GetPoint2dAt(i).X
                            Dim y As Double = pLine.GetPoint2dAt(i).Y
                            Dim pt As Point2d = New Point2d(Math.Round(x, RoundZnXY), Math.Round(y, RoundZnXY))
                            pLine.SetPointAt(i, pt)
                        Next i
                        FuncKadastrPlineRoundVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineRoundVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pLine2d As Autodesk.AutoCAD.DatabaseServices.Polyline2d = ObjEnt
                    Try
                        For Each vid As ObjectId In pLine2d
                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            v2d.Position = New Point3d(Math.Round(v2d.Position.X, RoundZnXY), Math.Round(v2d.Position.Y, RoundZnXY), 0)
                        Next
                        FuncKadastrPlineRoundVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineRoundVertex = False
                    End Try
                ElseIf TypeOf ObjEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                    Dim pLine3d As Autodesk.AutoCAD.DatabaseServices.Polyline3d = ObjEnt
                    Try
                        For Each vid As ObjectId In pLine3d
                            Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForWrite)
                            v3d.Position = New Point3d(Math.Round(v3d.Position.X, RoundZnXY), Math.Round(v3d.Position.Y, RoundZnXY), Math.Round(v3d.Position.Z, RoundZnH))
                        Next
                        FuncKadastrPlineRoundVertex = True
                    Catch ex As Exception
                        FuncKadastrPlineRoundVertex = False
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
    End Function
    '========================================================================================
    'функция делает реверс полилинии
    Public Shared Function FuncKadastrReversPline(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncKadastrReversPline = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline = objEnt
                    Try
                        Dim vertData As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                        For i As Integer = 0 To pl.NumberOfVertices - 1
                            Dim pvd As PerVertexData = New PerVertexData()
                            pvd.bulge = (If(i > 0, pl.GetBulgeAt(i - 1), 0))
                            pvd.startWidth = (If(i > 0, pl.GetStartWidthAt(i - 1), 0))
                            pvd.endWidth = (If(i > 0, pl.GetEndWidthAt(i - 1), 0))
                            pvd.pt = pl.GetPoint2dAt(i)
                            vertData.Add(pvd)
                        Next
                        For i As Integer = 0 To pl.NumberOfVertices - 1
                            Dim pvd As PerVertexData = vertData(pl.NumberOfVertices - (i + 1))
                            pl.SetPointAt(i, pvd.pt)
                            pl.SetBulgeAt(i, -pvd.bulge)
                            pl.SetStartWidthAt(i, pvd.endWidth)
                            pl.SetEndWidthAt(i, pvd.startWidth)
                        Next
                    Catch ex As Exception
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
        FuncKadastrReversPline = True
    End Function
    '========================================================================================
    'функция делает начальной северозападный угол полилинии
    Public Shared Function FuncKadastrSVAngle(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncKadastrSVAngle = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline = objEnt
                    Try
                        'выбираем габаритный контейнер для полилинии
                        Dim minCoord As Autodesk.AutoCAD.Geometry.Point3d = pl.GeometricExtents.MinPoint
                        Dim maxCoord As Autodesk.AutoCAD.Geometry.Point3d = pl.GeometricExtents.MaxPoint
                        Try
                            Dim vertData As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                            Dim vertDataNew As List(Of PerVertexData) = New List(Of PerVertexData)(pl.NumberOfVertices)
                            Dim vSV2d As Point2d = New Point2d(minCoord.X, maxCoord.Y)
                            Dim numbVert As Integer = 0
                            Dim MinDist As Double = 1000000000
                            For i As Integer = 0 To pl.NumberOfVertices - 1
                                Dim pvd As PerVertexData = New PerVertexData()
                                pvd.bulge = (If(i > 0, pl.GetBulgeAt(i - 1), 0))
                                pvd.startWidth = (If(i > 0, pl.GetStartWidthAt(i - 1), 0))
                                pvd.endWidth = (If(i > 0, pl.GetEndWidthAt(i - 1), 0))
                                pvd.pt = pl.GetPoint2dAt(i)
                                Dim vdist As Double = pl.GetPoint2dAt(i).GetVectorTo(vSV2d).Length
                                If vdist < MinDist Then
                                    numbVert = i
                                    MinDist = vdist
                                End If
                                vertData.Add(pvd)
                            Next
                            If numbVert > 0 Then
                                '======================================================================
                                ' делаем сортировку вершин
                                Dim nVert As Integer = 0
                                For Each temppvd As PerVertexData In vertData
                                    If nVert >= numbVert Then
                                        vertDataNew.Add(temppvd)
                                        nVert += 1
                                    Else
                                        nVert += 1
                                    End If
                                Next
                                nVert = 0
                                For Each temppvd As PerVertexData In vertData
                                    If nVert < numbVert Then
                                        vertDataNew.Add(temppvd)
                                        nVert += 1
                                    End If
                                Next
                                '=====================================================================
                                'пересобираем полилинию
                                For i As Integer = 0 To pl.NumberOfVertices - 1
                                    Dim pvd As PerVertexData = vertDataNew(i)
                                    pl.SetPointAt(i, pvd.pt)
                                    pl.SetBulgeAt(i, -pvd.bulge)
                                    pl.SetStartWidthAt(i, pvd.endWidth)
                                    pl.SetEndWidthAt(i, pvd.startWidth)
                                Next
                            End If
                        Catch ex As Exception
                        End Try
                    Catch ex As Exception
                    End Try
                End If
            End If
            acTrans.Commit()
        End Using
    End Function
    '===============================================================================================
    'функция заменяет дуговые сегменты 2д полилинии на линейные
    Public Shared Function FuncKadastrReplaseArcToLineSegmentPline(ByRef objEnt As Autodesk.AutoCAD.DatabaseServices.Entity) As Boolean
        FuncKadastrReplaseArcToLineSegmentPline = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            If objEnt.Id.IsNull = False Or objEnt.Id.IsErased = False Then
                If TypeOf objEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    Dim pl As Autodesk.AutoCAD.DatabaseServices.Polyline2d = objEnt
                End If
            End If
            acTrans.Commit()
        End Using
        FuncKadastrReplaseArcToLineSegmentPline = True
    End Function
    '==================================================================================================================================================================================================
    'функция ищет полярные координаты по углу и расстоянию
    Public Shared Function FuncReturnPolarCoordinates(ByVal startPoint As Point2d, ByVal endPoint As Point2d, ByVal a As Double, ByVal l As Double) As Point2d
        FuncReturnPolarCoordinates = New Point2d(-1, -1)
        Try
            Dim DirLine As Double = startPoint.GetVectorTo(endPoint).Angle
            DirLine += a
            If DirLine < 0 Then DirLine = DirLine + Math.PI * 2
            If DirLine > 2 * Math.PI Then DirLine = DirLine - Math.PI * 2
            Dim X As Double = startPoint.X + l * Math.Cos(DirLine)
            Dim Y As Double = startPoint.Y + l * Math.Sin(DirLine)
            FuncReturnPolarCoordinates = New Point2d(X, Y)
        Catch ex As Exception
        End Try
    End Function
    'функция ищет полярные координаты по углу и расстоянию
    Public Shared Function FuncReturnPolarCoordinates2(ByVal startPoint As Point2d, ByVal a As Double, ByVal l As Double, Optional round As Integer = 3) As Point2d
        FuncReturnPolarCoordinates2 = New Point2d(-1, -1)
        Try
            Dim DirLine As Double = a
            Dim X As Double = startPoint.X + l * Math.Cos(DirLine)
            Dim Y As Double = startPoint.Y + l * Math.Sin(DirLine)
            FuncReturnPolarCoordinates2 = New Point2d(Math.Round(X, round), Math.Round(Y, round))
        Catch ex As Exception
        End Try
    End Function
    'функция ищет полярные координаты по углу и расстоянию
    Public Shared Function FuncReturnPolar3dCoordinates2(ByVal startPoint As Point2d, ByVal a As Double, ByVal l As Double, Optional round As Integer = 3) As Point3d
        FuncReturnPolar3dCoordinates2 = New Point3d(-1, -1, 0)
        Try
            Dim DirLine As Double = a
            Dim X As Double = startPoint.X + l * Math.Cos(DirLine)
            Dim Y As Double = startPoint.Y + l * Math.Sin(DirLine)
            FuncReturnPolar3dCoordinates2 = New Point3d(Math.Round(X, round), Math.Round(Y, round), 0)
        Catch ex As Exception
        End Try
    End Function
    'функция ищет полярные координаты по углу и расстоянию
    Public Shared Function FuncReturnContinueLineCoordinates(ByVal startPoint As Point3d, ByVal endPoint As Point3d, ByVal dl As Double, Optional round As Integer = 3) As Point3d
        FuncReturnContinueLineCoordinates = New Point3d(-1, -1, -1)
        Try
            Dim L As Double = startPoint.GetVectorTo(endPoint).Length
            Dim k As Double = 0
            If dl < 0 Then
                k = 1 + dl / L
            Else
                k = -1 * (dl / L)
            End If

            Dim X As Double = endPoint.X + k * (startPoint.X - endPoint.X)
            Dim Y As Double = endPoint.Y + k * (startPoint.Y - endPoint.Y)
            Dim z As Double = endPoint.Z + k * (startPoint.Z - endPoint.Z)
            FuncReturnContinueLineCoordinates = New Point3d(Math.Round(X, round), Math.Round(Y, round), Math.Round(z, round))
        Catch ex As Exception
        End Try
    End Function
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'площадь участка
    '=============================================================================================================================
    Public Shared Function GetArea(pt1 As Point2d, pt2 As Point2d, pt3 As Point2d) As Double
        Return (((pt2.X - pt1.X) * (pt3.Y - pt1.Y)) - ((pt3.X - pt1.X) * (pt2.Y - pt1.Y))) / 2
    End Function
    Public Shared Function GetArea(arc As CircularArc2d) As Double
        Dim rad As Double = arc.Radius
        Dim ang As Double = If(arc.IsClockWise, arc.StartAngle - arc.EndAngle, arc.EndAngle - arc.StartAngle)
        Return rad * rad * (ang - Math.Sin(ang)) / 2
    End Function
    Public Shared Function GetArea(pline As Polyline) As Double
        Dim arc As New CircularArc2d()
        Dim area As Double = 0
        Dim last As Integer = pline.NumberOfVertices - 1
        Dim p0 As Point2d = pline.GetPoint2dAt(0)
        If pline.GetBulgeAt(0) <> 0 Then
            area += GetArea(pline.GetArcSegment2dAt(0))
        End If
        For i As Integer = 1 To last - 1
            area += GetArea(p0, pline.GetPoint2dAt(i), pline.GetPoint2dAt(i + 1))
            If pline.GetBulgeAt(i) <> 0 Then
                area += GetArea(pline.GetArcSegment2dAt(i))
            End If
        Next
        If (pline.GetBulgeAt(last) <> 0) AndAlso pline.Closed Then
            area += GetArea(pline.GetArcSegment2dAt(last))
        End If
        Return area
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона по коллекции точек
    Public Shared Function FuncSQRTPolyline(ByVal pt3dCollection As Point3dCollection, Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolyline = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If pt3dCollection.Count > 2 Then
            Dim Xish As Double = Math.Round(pt3dCollection.Item(0).X, RoundVertZn)
            Dim Yish As Double = Math.Round(pt3dCollection.Item(0).Y, RoundVertZn)
            For i As Integer = 0 To pt3dCollection.Count - 2
                X1 = Math.Round(pt3dCollection.Item(i).X, RoundVertZn)
                Y1 = Math.Round(pt3dCollection.Item(i).Y, RoundVertZn)
                X2 = Math.Round(pt3dCollection.Item(i + 1).X, RoundVertZn)
                Y2 = Math.Round(pt3dCollection.Item(i + 1).Y, RoundVertZn)
                PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
            Next
            If Not (Xish = pt3dCollection.Item(pt3dCollection.Count - 1).X And Yish = pt3dCollection.Item(pt3dCollection.Count - 1).Y) Then
                PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
            End If
            PSqr = Math.Abs(PSqr)
            FuncSQRTPolyline = Math.Round(PSqr / 2, RoundAreatZn)
        End If
    End Function

    Public Shared Function FuncSQRTPolyline2(ByVal acPline As Polyline, Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolyline2 = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If acPline.NumberOfVertices > 2 Then
            Dim Xish As Double = Math.Round(acPline.GetPoint2dAt(0).X, RoundVertZn)
            Dim Yish As Double = Math.Round(acPline.GetPoint2dAt(0).Y, RoundVertZn)
            For i As Integer = 0 To acPline.NumberOfVertices - 2
                X1 = Math.Round(acPline.GetPoint2dAt(i).X, RoundVertZn)
                Y1 = Math.Round(acPline.GetPoint2dAt(i).Y, RoundVertZn)
                X2 = Math.Round(acPline.GetPoint2dAt(i + 1).X, RoundVertZn)
                Y2 = Math.Round(acPline.GetPoint2dAt(i + 1).Y, RoundVertZn)
                PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
            Next
            If Not (Xish = acPline.EndPoint.X And Yish = acPline.EndPoint.Y) Then
                PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
            ElseIf acPline.Closed = True Then
                If acPline.StartPoint.X = acPline.EndPoint.X And acPline.StartPoint.Y = acPline.EndPoint.Y Then
                    PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
                End If
            End If
            PSqr = Math.Abs(PSqr)
            FuncSQRTPolyline2 = Math.Round(PSqr / 2, RoundAreatZn)
        End If
    End Function

    '==========================================================================
    'функция преобразут форматирует число с нужным количеством знаков после запятой
    Public Shared Function FuncFormatNumb(ByVal Dist As Double, Optional ByVal r As Integer = 2) As String
        Try
            Dim ind As Integer = r - 1
            Dist = Math.Round(Dist, r)
            'преобразуем в строку
            Dim NumbStr As String = CStr(Dist)
            Dim pos1 As Integer = NumbStr.LastIndexOf(".")
            If pos1 > -1 Then
                Dim ost As String = Mid(NumbStr, pos1 + 2)
                'смотрим сколько символов после точки
                ind = r - ost.Length - 1
            Else
                NumbStr = NumbStr & "."
            End If
            If ind > -1 Then
                For i As Integer = 0 To ind
                    NumbStr = NumbStr & "0"
                Next i
            End If
            FuncFormatNumb = NumbStr
        Catch
            FuncFormatNumb = Dist
            Exit Function
        End Try
    End Function


    '============================================================================================================
    '==========================================================================================================
    'функция проверяет пересечение двух прямоугольников (Верхний левый угол 1 прямоугольника, нижний правый угол 2 прямоугольника)
    Public Shared Function FuncIntersectedRectangle(ByVal rect1TopLeftX As Double, ByVal rect1TopLeftY As Double, ByVal rect1BottomRightX As Double, ByVal rect1BottomRightY As Double, ByVal rect2TopLeftX As Double, ByVal rect2TopLeftY As Double, ByVal rect2BottomRightX As Double, ByVal rect2BottomRightY As Double) As Boolean
        If (rect1TopLeftX > rect2BottomRightX) Then
            Return False
        ElseIf (rect2TopLeftX > rect1BottomRightX) Then
            Return False
        ElseIf (rect1TopLeftY > rect2BottomRightY) Then
            Return False
        ElseIf (rect2TopLeftY > rect1BottomRightY) Then
            Return False
        Else
            Return True
        End If
    End Function
    '==========================================================================================================
    'функция проверяет пересечение двух прямоугольников (Верхний левый угол 1 прямоугольника, нижний правый угол 2 прямоугольника)
    Public Shared Function FuncIntersectedPolylineByBounds(ByVal firstEnt As Entity, ByVal secondEnt As Entity) As Boolean
        Dim pointCollection1 As Point3dCollection = New Point3dCollection()
        firstEnt.BoundingBoxIntersectWith(secondEnt, Intersect.OnBothOperands, New Plane(), pointCollection1, 1, 1)
        If pointCollection1.Count > 0 Then
            Return True
        End If
        Dim extEnt1 As Extents3d = firstEnt.Bounds.Value
        Dim extEnt2 As Extents3d = secondEnt.Bounds.Value
        Dim rect1 As New Rectangle(extEnt1.MinPoint.X, extEnt1.MaxPoint.Y, extEnt1.MaxPoint.X - extEnt1.MinPoint.X, extEnt1.MaxPoint.Y - extEnt1.MinPoint.Y)
        Dim rect2 As New Rectangle(extEnt2.MinPoint.X, extEnt2.MaxPoint.Y, extEnt2.MaxPoint.X - extEnt2.MinPoint.X, extEnt2.MaxPoint.Y - extEnt2.MinPoint.Y)
        ' Проверка пересечения по оси X
        Dim intersectX As Boolean = rect1.X < rect2.X + rect2.Width AndAlso rect1.X + rect1.Width > rect2.X
        ' Проверка пересечения по оси Y
        Dim intersectY As Boolean = rect1.Y < rect2.Y + rect2.Height AndAlso rect1.Y + rect1.Height > rect2.Y
        ' Если пересекаются по обеим осям, прямоугольники пересекаются
        Return intersectX AndAlso intersectY
    End Function

    ' Определение типа пересечения
    Public Shared Function GetIntersectionType(rect1 As Rectangle2, rect2 As Rectangle2) As Integer
        ' Проверка на отсутствие пересечения
        If (rect1.X2 < rect2.X1) OrElse (rect2.X2 < rect1.X1) OrElse
           (rect1.Y2 < rect2.Y1) OrElse (rect2.Y2 < rect1.Y1) Then
            Return 0
        End If

        ' Проверка, находится ли rect1 внутри rect2
        If (rect1.X1 >= rect2.X1) AndAlso (rect1.Y1 >= rect2.Y1) AndAlso
           (rect1.X2 <= rect2.X2) AndAlso (rect1.Y2 <= rect2.Y2) Then
            Return 2
        End If

        ' Проверка, находится ли rect2 внутри rect1
        If (rect2.X1 >= rect1.X1) AndAlso (rect2.Y1 >= rect1.Y1) AndAlso
           (rect2.X2 <= rect1.X2) AndAlso (rect2.Y2 <= rect1.Y2) Then
            Return 2
        End If

        ' Обычное пересечение
        Return 1
    End Function


    Public Shared Function IsPointOnLineSegment(startLinePoint As Point3d, endLinePoint As Point3d, pt As Point3d) As Boolean
        ' Проверка на совпадение с концами
        If pt.IsEqualTo(startLinePoint) OrElse pt.IsEqualTo(endLinePoint) Then
            Return False
        End If
        Try
            Dim fullLenght As Double = startLinePoint.Convert2d(New Plane).GetVectorTo(endLinePoint.Convert2d(New Plane)).Length
            If fullLenght = 0 Then
                Return False
            End If
            Dim lLenght As Double = startLinePoint.Convert2d(New Plane).GetVectorTo(pt.Convert2d(New Plane)).Length
            If lLenght < fullLenght Then
                Return True
            End If
        Catch ex As Exception
            ' Точка не лежит на линии
            Return False
        End Try
    End Function


    Public Shared Function TryConvertToNumber(ByVal strNumber As String) As Integer
        If String.IsNullOrWhiteSpace(strNumber) Then
            Return 0
        End If
        Dim trimmed As String = strNumber.Trim()
        Dim intValue As Integer
        If Integer.TryParse(trimmed, intValue) Then
            Return intValue
        End If
        ' Пробуем как Double и округляем/отбрасываем дробную часть
        Dim dblValue As Double
        If Double.TryParse(trimmed, dblValue) Then
            Return CInt(Math.Truncate(dblValue)) ' Отбрасываем дробную часть
            ' Или используйте CInt(dblValue) - округление
        End If
        ' Если ничего не вышло
        Return 0
    End Function




    '============================================================================================================================================================
    'функция удлинняет линию
    Public Shared Function FuncExtendLine(ByRef acline As Line, ByVal startLenght As Double, ByVal endLenght As Double, Optional round As Integer = 3) As Boolean
        FuncExtendLine = False
        Try
            Dim L As Double = acline.Length
            If L = 0 Then
                Return False
            End If
            If startLenght <> 0 Then
                Dim k1 As Double = -1 * (startLenght / L)
                Dim X1 As Double = acline.StartPoint.X + k1 * (acline.EndPoint.X - acline.StartPoint.X)
                Dim Y1 As Double = acline.StartPoint.Y + k1 * (acline.EndPoint.Y - acline.StartPoint.Y)
                Dim z1 As Double = acline.StartPoint.Z + k1 * (acline.EndPoint.Z - acline.StartPoint.Z)
                Dim startPoint As Point3d = New Point3d(X1, Y1, z1)
                acline.StartPoint = startPoint
            End If
            If endLenght <> 0 Then
                Dim k2 As Double = endLenght / L
                Dim X2 As Double = acline.EndPoint.X + k2 * (acline.EndPoint.X - acline.StartPoint.X)
                Dim Y2 As Double = acline.EndPoint.Y + k2 * (acline.EndPoint.Y - acline.StartPoint.Y)
                Dim z2 As Double = acline.EndPoint.Z + k2 * (acline.EndPoint.Z - acline.StartPoint.Z)
                Dim endPoint As Point3d = New Point3d(X2, Y2, z2)
                acline.EndPoint = endPoint
            End If
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function


    Public Shared Function FuncFindCrossCenterPointPolyline(ByVal Pline As Polyline) As Point3d
        FuncFindCrossCenterPointPolyline = New Point3d(-1, -1, -1)
        Try
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'определение центра полигона 
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '1-проверяем замнута ли полилиния
            Dim boolClosedPline As Boolean = False
            If Pline.Closed = True Then
                boolClosedPline = True
            ElseIf Pline.StartPoint.GetVectorTo(Pline.EndPoint).Length < 0.01 Then
                boolClosedPline = True
            Else
                FuncFindCrossCenterPointPolyline = Pline.GetPointAtDist(Pline.Length / 2)
                Exit Function
            End If
            '2-ищем габаритный контейнер
            Dim extPoint As Extents3d = Pline.GeometricExtents
            '3-проводим горизонтальную линию по середине габаритного контейнера (ГОРИЗОНТАЛЬНУЮ)
            Dim pt1 As Point2d = New Point2d(extPoint.MinPoint.X, (extPoint.MinPoint.Y + extPoint.MaxPoint.Y) / 2)
            Dim pt2 As Point2d = New Point2d(extPoint.MaxPoint.X, (extPoint.MinPoint.Y + extPoint.MaxPoint.Y) / 2)
            '3-заносим в массив все пересечения фигуры с этой горизонтальной линией
            Dim CountArrayPt As Integer = 0
            Dim ArrayPt As Double(,) = Nothing
            For i As Integer = 0 To Pline.NumberOfVertices - 1
                Dim LineSegment As LineSegment2d = Pline.GetLineSegment2dAt(i)
                If LineSegment.Length > 0 Then
                    Dim pt3 As Point2d = LineSegment.StartPoint
                    Dim pt4 As Point2d = LineSegment.EndPoint
                    Dim IntersectPoint As Point2d = New Point2d
                    Dim boolIntersectLine As Boolean = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, IntersectPoint)
                    If boolIntersectLine = True Then
                        ReDim Preserve ArrayPt(1, CountArrayPt)
                        ArrayPt(0, CountArrayPt) = IntersectPoint.X
                        ArrayPt(1, CountArrayPt) = IntersectPoint.Y
                        CountArrayPt += 1
                    End If
                End If
            Next
            '4-сортируем массив (в массиве должно быть по меньшей мере 2 точки)
            If ArrayPt.GetUpperBound(1) > 0 Then
                Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(ArrayPt, 0)
                Dim DistInsert As Double = 0 'точки должна быть максимально удалена от точек пересечения
                If ArrayPt.GetUpperBound(1) > 0 Then
                    For j As Integer = 0 To ArrayPt.GetUpperBound(1) - 1
                        Dim xStart As Double = ArrayPt(0, j)
                        Dim yStart As Double = ArrayPt(1, j)
                        Dim xEnd As Double = ArrayPt(0, j + 1)
                        Dim yEnd As Double = ArrayPt(1, j)
                        Dim ptStart As Point2d = New Point2d(xStart, yStart)
                        Dim ptEnd As Point2d = New Point2d(xEnd, yEnd)
                        Dim TempDist As Double = ptStart.GetVectorTo(ptEnd).Length
                        Dim FindCenterPoint As Point2d = New Point2d((ArrayPt(0, j) + ArrayPt(0, j + 1)) / 2, ArrayPt(1, j))
                        Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(Pline, FindCenterPoint)
                        If PointCrosPoligon = True Then
                            If TempDist > DistInsert Then
                                FuncFindCrossCenterPointPolyline = New Point3d(FindCenterPoint.X, FindCenterPoint.Y, 0)
                                DistInsert = TempDist
                            End If
                        End If
                    Next j
                End If
            End If
            '5-если неудача попробуем простой вариант
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                Dim FindCenterPoint As Point2d = New Point2d((extPoint.MaxPoint.X + extPoint.MinPoint.X) / 2, (extPoint.MaxPoint.Y + extPoint.MinPoint.Y) / 2)
                Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(Pline, FindCenterPoint)
                If PointCrosPoligon = True Then
                    FuncFindCrossCenterPointPolyline = New Point3d(FindCenterPoint.X, FindCenterPoint.Y, 0)
                End If
            End If
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                FuncFindCrossCenterPointPolyline = Pline.StartPoint
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '3d моделирование
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'смещение солида вдоль оси Z по точке
    Public Shared Function FuncMoveZSolid3d(ByVal userSolid3d As Solid3d, ByVal moveDist As Double) As Boolean
        FuncMoveZSolid3d = False
        If IsNothing(userSolid3d) Then
            If userSolid3d.Id.IsNull = False Then
                Dim p1 As Point3d = userSolid3d.MassProperties.Extents.MinPoint
                Dim p2 As Point3d = New Point3d(p1.X, p1.Y, p1.Z + moveDist)
                Dim v As Vector3d = p1.GetAsVector() - p2.GetAsVector()
                Dim m As Matrix3d = Matrix3d.Displacement(v)
                userSolid3d.TransformBy(m)
                FuncMoveZSolid3d = False
            End If
        End If
    End Function
    '===========================================================================================================
    'функция возвращает ближайшую точку пересечения тела и отрезка
    Public Shared Function FuncFindIntersectPointBySolid(ByVal acEnt As Autodesk.AutoCAD.DatabaseServices.Entity, ByVal startPoint As Point3d, ByVal endPoint As Point3d) As Point3d
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acLine As Line = New Line(startPoint, endPoint)
                acLine.SetDatabaseDefaults()
                acBlkTblRec.AppendEntity(acLine)
                acTrans.AddNewlyCreatedDBObject(acLine, True)
                Dim pointColl As Point3dCollection = New Point3dCollection()
                If TypeOf acEnt Is Solid3d Then
                    Dim acSolid3d As Solid3d = acEnt
                    acSolid3d.IntersectWith(acLine, Intersect.ExtendArgument, pointColl, 0, 0)
                ElseIf TypeOf acEnt Is SubDMesh Then
                    Dim acDMesh As SubDMesh = acEnt
                    acDMesh.IntersectWith(acLine, Intersect.ExtendThis, pointColl, IntPtr.Zero, IntPtr.Zero)

                End If
                acLine.Erase()
                If pointColl.Count > 0 Then
                    Dim Dist As Double = 0
                    For i As Integer = 0 To pointColl.Count - 1
                        Dim tempPt As Point3d = pointColl.Item(i)
                        If Dist = 0 Then
                            FuncFindIntersectPointBySolid = tempPt
                        Else
                            Dim tempDist As Double = tempPt.GetVectorTo(startPoint).Length
                            If tempDist < Dist Then
                                FuncFindIntersectPointBySolid = tempPt
                                Dist = tempDist
                            End If
                        End If
                    Next i
                End If
            End Using
        Catch ex As System.Exception
        End Try
    End Function

    Public Function FindSegmentCircleIntersections(circleCenter As Point2d, radius As Double, segmentStart As Point2d, segmentEnd As Point2d) As Point2d()

        Dim dx, dy, A, B, C, det, t1, t2 As Double
        Dim result As New System.Collections.Generic.List(Of Point2d)

        dx = segmentEnd.X - segmentStart.X
        dy = segmentEnd.Y - segmentStart.Y

        A = dx * dx + dy * dy
        B = 2 * (dx * (segmentStart.X - circleCenter.X) + dy * (segmentStart.Y - circleCenter.Y))
        C = (segmentStart.X - circleCenter.X) * (segmentStart.X - circleCenter.X) +
            (segmentStart.Y - circleCenter.Y) * (segmentStart.Y - circleCenter.Y) -
            radius * radius

        det = B * B - 4 * A * C

        If A <= 0.0000001 Or det < 0 Then
            ' Нет решений (отрезок и окружность не пересекаются)
            Return result.ToArray()
        ElseIf det = 0 Then
            ' Одно решение (касание)
            t1 = -B / (2 * A)
            If t1 >= 0 AndAlso t1 <= 1 Then
                ' Точка находится на отрезке
                result.Add(New Point2d(
                    segmentStart.X + t1 * dx,
                    segmentStart.Y + t1 * dy))
            End If
        Else
            ' Два решения
            t1 = (-B + Math.Sqrt(det)) / (2 * A)
            t2 = (-B - Math.Sqrt(det)) / (2 * A)

            ' Проверяем, находится ли точка на отрезке
            If t1 >= 0 AndAlso t1 <= 1 Then
                result.Add(New Point2d(
                    segmentStart.X + t1 * dx,
                    segmentStart.Y + t1 * dy))
            End If

            If t2 >= 0 AndAlso t2 <= 1 Then
                result.Add(New Point2d(
                    segmentStart.X + t2 * dx,
                    segmentStart.Y + t2 * dy))
            End If
        End If

        ' Если две точки совпадают (например, когда отрезок касается окружности)
        ' Оставляем только одну точку
        If result.Count = 2 AndAlso
           Math.Abs(result(0).X - result(1).X) < 0.0001 AndAlso
           Math.Abs(result(0).Y - result(1).Y) < 0.0001 Then
            result.RemoveAt(1)
        End If

        Return result.ToArray()
    End Function



    Public Shared Function FuncZoomToWindow(ByVal minPoint As Point2d, ByVal maxPoint As Point2d) As Boolean
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
        Dim view As ViewTableRecord = ed.GetCurrentView()
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            view.Width = maxPoint.X - minPoint.X
            view.Height = maxPoint.Y - minPoint.Y
            view.CenterPoint = New Point2d(minPoint.X + (view.Width / 2), minPoint.Y + (view.Height / 2))
            ed.SetCurrentView(view)
            trans.Commit()
        End Using
        Return True
    End Function

End Class



