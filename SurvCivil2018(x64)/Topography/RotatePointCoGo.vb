Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Public Class RotatePointCoGo
    '=============================================================================================
    'разворот высотных отметок вдоль полилинии
    <CommandMethod("RotatePointCoGo")>
    Public Sub RotatePointCoGo()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Try
            Dim strDWGName As String = acDoc.Name
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

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormRotate As RotPointCoGo = New RotPointCoGo
        FormRotate.ShowDialog()
        If FormRotate.boolRotate = False Then
            locker.Dispose()
            Exit Sub
        End If
        Dim boolMarker As Boolean = False
        Dim boolLabel As Boolean = False
        If FormRotate.RadioButton2.Checked = True Then
            boolMarker = True
        ElseIf FormRotate.RadioButton1.Checked = True Then
            boolLabel = True
        End If
        Dim Lenght As Double = FormRotate.NumericUpDown1.Value
        Dim dx As Double = FormRotate.NumericUpDown2.Value
        Dim dy As Double = FormRotate.NumericUpDown3.Value
        Dim k As Integer = 1
        If FormRotate.RadioButton4.Checked = True Then
            k = 1
        Else
            k = -1
        End If
        Do
            Using acTrans As Transaction = docBase.TransactionManager.StartTransaction()
                Dim NPoint As Integer = 0 'количество точек
                Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                Dim Pline2D As Polyline2d = Nothing
                Dim Pline3D As Polyline3d = Nothing
                Dim DLine As Line = Nothing
                'габаритный контейнер
                Dim MinCoord As Point3d = Nothing
                Dim MaxCoord As Point3d = Nothing
                Dim ArrayPointXY As Double(,) = Nothing 'массив для координат точек
                Dim SelectPline As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию вдоль которой требуется развернуть отметки:  ")
                If SelectPline.Status = PromptStatus.OK Then
                    Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectPline.ObjectId, OpenMode.ForRead)
                    '=================================================================================================================
                    'забираем точки перелома линии в массив
                    'полилиния
                    If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                        Pline = acEnt1
                        NPoint = Pline.NumberOfVertices
                        ReDim ArrayPointXY(1, Pline.NumberOfVertices)
                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                            Dim PointXY As Point2d = Pline.GetPoint2dAt(i)
                            ArrayPointXY(0, i) = PointXY.X
                            ArrayPointXY(1, i) = PointXY.Y
                        Next i
                        MinCoord = Pline.GeometricExtents.MinPoint
                        MaxCoord = Pline.GeometricExtents.MaxPoint
                        '===================================================================================
                        '2д полилиния
                    ElseIf TypeOf acEnt1 Is Polyline2d Then
                        Pline2D = acEnt1
                        For Each vtid As ObjectId In Pline2D
                            Dim vt As Vertex2d = acTrans.GetObject(vtid, OpenMode.ForRead)
                            ReDim Preserve ArrayPointXY(1, NPoint)
                            ArrayPointXY(0, NPoint) = vt.Position.X
                            ArrayPointXY(1, NPoint) = vt.Position.Y
                            NPoint += 1
                        Next
                        MinCoord = Pline2D.GeometricExtents.MinPoint
                        MaxCoord = Pline2D.GeometricExtents.MaxPoint
                        '===================================================================================
                        '3д полилиния
                    ElseIf TypeOf acEnt1 Is Polyline3d Then
                        Pline3D = acEnt1
                        For Each vtid As ObjectId In Pline3D
                            Dim v3d As PolylineVertex3d = acTrans.GetObject(vtid, OpenMode.ForRead)
                            ReDim Preserve ArrayPointXY(1, NPoint)
                            ArrayPointXY(0, NPoint) = v3d.Position.X
                            ArrayPointXY(1, NPoint) = v3d.Position.Y
                            NPoint += 1
                        Next
                        MinCoord = Pline3D.GeometricExtents.MinPoint
                        MaxCoord = Pline3D.GeometricExtents.MaxPoint
                        '=============================================================================
                        'линия
                    ElseIf TypeOf acEnt1 Is Line Then
                        DLine = acEnt1
                        ReDim ArrayPointXY(1, 1)
                        ArrayPointXY(0, 0) = DLine.StartPoint.X
                        ArrayPointXY(1, 0) = DLine.StartPoint.Y
                        ArrayPointXY(0, 1) = DLine.EndPoint.X
                        ArrayPointXY(1, 1) = DLine.EndPoint.Y
                        MinCoord = DLine.GeometricExtents.MinPoint
                        MaxCoord = DLine.GeometricExtents.MaxPoint
                        '===================================================================================
                    End If
                    '=======================================================================
                    'ищем кого точки которые находятся в габаритном контейнере полилинии
                    Dim ArrayPoint As Integer() = Nothing
                    Dim Count1 As Integer = 0
                    Dim CountPoint As Integer = CivilDoc.CogoPoints.Count
                    If CountPoint > 0 Then
                        For Each PointId As ObjectId In CivilDoc.CogoPoints
                            Dim PointCOGO1 As CogoPoint = PointId.GetObject(OpenMode.ForRead)
                            If PointCOGO1.Northing >= MinCoord.Y - 10 And PointCOGO1.Easting >= MinCoord.X - 10 Then
                                If PointCOGO1.Northing <= MaxCoord.Y + 10 And PointCOGO1.Easting <= MaxCoord.X + 10 Then
                                    ReDim Preserve ArrayPoint(Count1)
                                    ArrayPoint(Count1) = PointCOGO1.PointNumber
                                    Count1 += 1
                                End If
                            End If
                        Next
                    End If
                    'начинаем перебор вершин выбранного объекта
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'если это полилиния
                    If IsNothing(Pline) = False Then
                        'удаляем повторяющиеся вершины
                        Dim boolDelVertexPline As Boolean = MathFunction.FuncKadastrPlineDelDubleVertex(Pline, 0.01)
                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                            'координата узла полилинии
                            Dim PtTemp As Point3d = Pline.GetPoint3dAt(i)
                            'дистанция до узла полилиниии
                            Dim DistTemp As Double = Pline.GetDistAtPoint(PtTemp)
                            'координата смещения для второй точки
                            Dim Pttemp2 As Point3d = PtTemp
                            'угол поворота для метки по умолчанию                                                    
                            Dim RotateMarkerDisplay As Double = 0
                            'угол поворота (расчетный)
                            Dim RotateMarker As Double = Pline.GetFirstDerivative(PtTemp).Convert2d(New Plane).Angle
                            '=====================================================================================
                            Select Case RotateMarker
                                Case 0 To Math.PI / 2
                                    RotateMarkerDisplay = RotateMarker
                                Case Math.PI / 2 To Math.PI
                                    RotateMarkerDisplay = RotateMarker'RotateMarker + Math.PI
                                Case Math.PI To (3 * Math.PI) / 2
                                    RotateMarkerDisplay = RotateMarker 'RotateMarker - Math.PI
                                Case Else
                                    RotateMarkerDisplay = RotateMarker
                            End Select
                            'находим в узле cogo точку
                            If IsArray(ArrayPoint) = True Then
                                Dim Point As Point3d = New Point3d(PtTemp.X, PtTemp.Y, 0)
                                Dim PointCogoBKN As CogoPoint = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                If IsNothing(PointCogoBKN) = False Then
                                    Dim numb As Integer = PointCogoBKN.PointNumber
                                    If numb <> Val(FormRotate.Button2.Tag) Then
                                        If boolLabel = True Then
                                            PointCogoBKN.LabelRotation = RotateMarkerDisplay
                                        ElseIf boolMarker = True Then
                                            PointCogoBKN.MarkerRotation = RotateMarkerDisplay
                                        Else
                                            PointCogoBKN.LabelRotation = RotateMarkerDisplay
                                            PointCogoBKN.MarkerRotation = RotateMarkerDisplay
                                        End If

                                        Dim tempdx1 As Double = 0
                                        Dim tempdy1 As Double = 0
                                        Dim tempdx2 As Double = 0
                                        Dim tempdy2 As Double = 0
                                        If FormRotate.RadioButton4.Checked = True Then
                                            Dim fi As Double = RotateMarker + Math.PI / 2
                                            If fi < 0 Then fi = fi + Math.PI * 2
                                            If fi > 2 * Math.PI Then fi = fi - Math.PI * 2
                                            tempdx1 = dx * Math.Cos(RotateMarker)
                                            tempdy1 = dx * Math.Sin(RotateMarker)
                                            tempdx2 = dy * Math.Cos(fi)
                                            tempdy2 = dy * Math.Sin(fi)
                                        Else
                                            Dim fi As Double = RotateMarker - Math.PI / 2
                                            If fi < 0 Then fi = fi + Math.PI * 2
                                            If fi > 2 * Math.PI Then fi = fi - Math.PI * 2
                                            tempdx1 = dx * Math.Cos(RotateMarker)
                                            tempdy1 = dx * Math.Sin(RotateMarker)
                                            tempdx2 = dy * Math.Cos(fi)
                                            tempdy2 = dy * Math.Sin(fi)
                                        End If
                                        Dim PojntLabel As Point3d = PointCogoBKN.Location
                                        Dim NewPointLabel As Point3d = New Point3d(PtTemp.X + tempdx1 + tempdx2, PtTemp.Y + tempdy1 + tempdy2, PojntLabel.Z)
                                        PointCogoBKN.LabelLocation = NewPointLabel
                                    End If
                                End If
                            End If
                        Next
                        '===============================================================================================
                        '2d полилиния
                    ElseIf IsNothing(Pline2D) = False Then
                        'удаляем повторяющиеся вершины
                        Dim boolDelVertexPline As Boolean = MathFunction.FuncKadastrPlineDelDubleVertex(Pline2D, 0.01)
                        Dim CurvePline As Curve = Pline2D
                        For Each vtid As ObjectId In Pline2D
                            Try
                                Dim vn As Vertex2d = acTrans.GetObject(vtid, OpenMode.ForRead)
                                Dim RotateMarkerDisplay As Double = 0
                                Dim RotateMarker As Double = 0
                                Dim PtTemp As Point3d = vn.Position
                                PtTemp = CurvePline.GetClosestPointTo(PtTemp, True)
                                Dim DistTemp As Double = CurvePline.GetDistAtPoint(PtTemp)
                                Dim Pttemp2 As Point3d = PtTemp

                                DistTemp = DistTemp + dx
                                If DistTemp > Pline2D.Length Then
                                    DistTemp = DistTemp + 0.01
                                End If
                                Try
                                    If DistTemp < 0 Then
                                        RotateMarker = CurvePline.GetFirstDerivative(PtTemp).Convert2d(New Plane).Angle
                                    ElseIf DistTemp > Pline2D.Length Then
                                        RotateMarker = CurvePline.GetFirstDerivative(PtTemp).Convert2d(New Plane).Angle
                                    Else
                                        Pttemp2 = CurvePline.GetPointAtDist(DistTemp)
                                        Dim vect As Vector3d = PtTemp.GetVectorTo(Pttemp2)
                                        Dim xy As Plane = New Plane(Point3d.Origin, Vector3d.ZAxis)
                                        vect = PtTemp.GetVectorTo(Pttemp2)
                                        xy = New Plane(Point3d.Origin, Vector3d.ZAxis)
                                        RotateMarker = vect.AngleOnPlane(xy)
                                    End If
                                Catch ex As ArgumentException
                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                End Try
                                '=====================================================================================
                                Select Case RotateMarker
                                    Case 0 To Math.PI / 2
                                        RotateMarkerDisplay = RotateMarker
                                    Case Math.PI / 2 To Math.PI
                                        RotateMarkerDisplay = RotateMarker'RotateMarker + Math.PI
                                    Case Math.PI To (3 * Math.PI) / 2
                                        RotateMarkerDisplay = RotateMarker 'RotateMarker - Math.PI
                                    Case Else
                                        RotateMarkerDisplay = RotateMarker
                                End Select
                                'находим в узле cogo точку
                                If IsArray(ArrayPoint) = True Then
                                    Dim Point As Point3d = New Point3d(PtTemp.X, PtTemp.Y, 0)
                                    Dim PointCogoBKN As CogoPoint = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, Lenght)
                                    If IsNothing(PointCogoBKN) = False Then
                                        Dim numb As Integer = PointCogoBKN.PointNumber
                                        If numb <> Val(FormRotate.Button2.Tag) Then
                                            If boolLabel = True Then
                                                PointCogoBKN.LabelRotation = RotateMarkerDisplay
                                            ElseIf boolMarker = True Then
                                                PointCogoBKN.MarkerRotation = RotateMarkerDisplay
                                            Else
                                                PointCogoBKN.LabelRotation = RotateMarkerDisplay
                                                PointCogoBKN.MarkerRotation = RotateMarkerDisplay
                                            End If

                                            Dim tempdx1 As Double = 0
                                            Dim tempdy1 As Double = 0
                                            Dim tempdx2 As Double = 0
                                            Dim tempdy2 As Double = 0
                                            If FormRotate.RadioButton4.Checked = True Then
                                                Dim fi As Double = RotateMarker + Math.PI / 2
                                                If fi < 0 Then fi = fi + Math.PI * 2
                                                If fi > 2 * Math.PI Then fi = fi - Math.PI * 2
                                                tempdx1 = dx * Math.Cos(RotateMarker)
                                                tempdy1 = dx * Math.Sin(RotateMarker)
                                                tempdx2 = dy * Math.Cos(fi)
                                                tempdy2 = dy * Math.Sin(fi)
                                            Else
                                                Dim fi As Double = RotateMarker - Math.PI / 2
                                                If fi < 0 Then fi = fi + Math.PI * 2
                                                If fi > 2 * Math.PI Then fi = fi - Math.PI * 2
                                                tempdx1 = dx * Math.Cos(RotateMarker)
                                                tempdy1 = dx * Math.Sin(RotateMarker)
                                                tempdx2 = dy * Math.Cos(fi)
                                                tempdy2 = dy * Math.Sin(fi)
                                            End If
                                            Dim PojntLabel As Point3d = PointCogoBKN.Location
                                            Dim NewPointLabel As Point3d = New Point3d(PtTemp.X + tempdx1 + tempdx2, PtTemp.Y + tempdy1 + tempdy2, PojntLabel.Z)
                                            PointCogoBKN.LabelLocation = NewPointLabel
                                        End If
                                    End If
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
line1:
                    acTrans.Commit()
                Else
                    Exit Do
                End If
            End Using
        Loop
        locker.Dispose()
    End Sub
End Class
