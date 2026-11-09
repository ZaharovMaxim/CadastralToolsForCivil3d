Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.LinearReferencing
Imports Entity = Autodesk.AutoCAD.DatabaseServices.Entity
Imports Point = NetTopologySuite.Geometries.Point

Public Class ProjectFence
    ''' <summary>
    ''' Создаёт 3D-тело путём сдвига профиля (например, окружности) вдоль траектории.
    ''' Аналог команды СДВИГ (SWEEP).
    ''' </summary>
    ''' <param name="profileId">ObjectId профиля (замкнутая кривая: Circle, Polyline, Region)</param>
    ''' <param name="trajectoryId">ObjectId траектории (Polyline, Line, Arc, Spline)</param>
    ''' <param name="db">База данных чертежа</param>
    ''' <returns>ObjectId созданного тела или ObjectId.Null при ошибке</returns>
    Public Shared Function CreateSweptSolid(profileId As ObjectId, trajectoryId As ObjectId, db As Database) As Solid3d
        ' Проверка входных данных
        Dim result As Solid3d = Nothing
        If profileId.IsNull OrElse trajectoryId.IsNull Then
            Return Nothing
        End If
        Dim trans As Transaction = db.TransactionManager.StartTransaction()
        Using trans
            ' Получаем профиль и траекторию
            Dim profile As Entity = TryCast(trans.GetObject(profileId, OpenMode.ForRead), Entity)
            Dim path As Curve = TryCast(trans.GetObject(trajectoryId, OpenMode.ForRead), Curve)

            If profile Is Nothing OrElse path Is Nothing Then
                Return Nothing
            End If
            ' Профиль должен быть замкнутой кривой (Circle, замкнутая Polyline, Region)
            ' Траектория должна быть кривой (Curve)
            ' Создаём тело
            Dim solid As New Solid3d()
            ' Настраиваем параметры сдвига
            Dim sweepOptions As New SweepOptionsBuilder()
            ' Выравнивать профиль по нормали к траектории (как в команде СДВИГ)
            sweepOptions.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
            ' Отключаем "banking" (естественный поворот), чтобы избежать неожиданного вращения
            sweepOptions.Bank = False
            Try
                solid.CreateSweptSolid(profile, path, sweepOptions.ToSweepOptions())
            Catch ex As Exception
                ' Если не удалось создать тело (например, самопересечение), возвращаем Null
                solid.Dispose()
                Return Nothing
            End Try
            ' Добавляем тело в пространство модели
            Dim bt As BlockTable = TryCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = TryCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
            Dim solidId As ObjectId = ms.AppendEntity(solid)
            trans.AddNewlyCreatedDBObject(solid, True)
            trans.Commit()
            result = solid
        End Using
        Return result
    End Function
    'процедура трансформирует полилинию из вида сверху в вид СПЕРЕДИ
    Public Shared Sub TransformPolylineToFrontView(polylineId As ObjectId, db As Database)
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim polyline As Polyline = TryCast(trans.GetObject(polylineId, OpenMode.ForWrite), Polyline)
            If polyline Is Nothing Then Return
            ' 1. Перемещаем в точку (0,0,0) — за базовую точку берём первую вершину
            Dim firstPoint As Point3d = polyline.GetPoint3dAt(0)
            Dim moveVec As Vector3d = Point3d.Origin.GetVectorTo(firstPoint).Negate()
            polyline.TransformBy(Matrix3d.Displacement(moveVec))
            ' 2. Поворачиваем на 90° вокруг оси X (из XY в XZ)
            polyline.TransformBy(Matrix3d.Rotation(Math.PI / 2.0, Vector3d.XAxis, Point3d.Origin))
            trans.Commit()
        End Using
    End Sub
    'перемещение солида в нужную точку и его поворот в плоскости XY
    Public Shared Sub MoveAndRotateSolidByBasePoint(solidId As ObjectId, basePoint As Point3d, targetPoint As Point3d, angleDegrees As Double, db As Database)

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim solid As Solid3d = TryCast(trans.GetObject(solidId, OpenMode.ForWrite), Solid3d)
            If solid Is Nothing Then Return

            ' 1. Перемещение: basePoint → targetPoint
            Dim moveVec As Vector3d = targetPoint.GetVectorTo(basePoint).Negate()
            solid.TransformBy(Matrix3d.Displacement(moveVec))

            ' 2. Поворот в плоскости XY вокруг оси Z, базовая точка — targetPoint
            Dim angleRadians As Double = angleDegrees * Math.PI / 180.0
            solid.TransformBy(Matrix3d.Rotation(angleRadians, Vector3d.ZAxis, targetPoint))

            trans.Commit()
        End Using
    End Sub

    <CommandMethod("CreateProjectMeshFence")>
    Public Sub CreateProjectMeshFence()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Db As Database = acDoc.Database
        Dim ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                ' --- 1. Выбор траектории для профиля ---
                Dim acSSPr1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите траекторию: ")
                If acSSPr1.Status <> PromptStatus.OK Then Exit Sub
                Dim entPoly As Entity = acTrans.GetObject(acSSPr1.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entPoly Is Polyline) Then
                    MsgBox("Выбранный примитив не является Полилинией. Макро прервано.")
                    Exit Sub
                End If
                Dim acPoly As Polyline = DirectCast(entPoly, Polyline)
                ' --- 2. Работаем с КОПИЕЙ полилинии ---
                Dim profilePoly As Polyline = DirectCast(acPoly.Clone(), Polyline)
                acBlkTblRec.AppendEntity(profilePoly)
                acTrans.AddNewlyCreatedDBObject(profilePoly, True)
                ' Трансформируем копию: перемещаем в (0,0,0) и поворачиваем в плоскость XZ
                TransformPolylineToFrontView(profilePoly.ObjectId, Db)
                ' --- 3. Создаём окружность профиля в начале копии ---
                Dim startPt As Point3d = profilePoly.GetPoint3dAt(0)
                Dim centerPoint As New Point3d(startPt.X, startPt.Y, startPt.Z)
                Dim radius As Double = 0.002
                Dim acCircle As New Circle(centerPoint, Vector3d.ZAxis, radius)
                acBlkTblRec.AppendEntity(acCircle)
                acTrans.AddNewlyCreatedDBObject(acCircle, True)
                ' --- 4. Создаём тело сдвигом вдоль копии ---
                Dim userSolid As Solid3d = CreateSweptSolid(acCircle.ObjectId, profilePoly.ObjectId, Db)
                If userSolid Is Nothing Then
                    MsgBox("Не удалось создать тело. Проверьте профиль и траекторию.")
                    Exit Sub
                End If
                ' Базовая точка тела — начало выдавливания (центр профиля)
                Dim basePoint As Point3d = centerPoint
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ' --- 5. Выбор направления построения ограждения ---
                Dim acSSPr0 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию - ограждение: ")
                If acSSPr0.Status <> PromptStatus.OK Then Exit Sub
                Dim entPolyTr As Entity = acTrans.GetObject(acSSPr0.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entPolyTr Is Polyline) Then
                    MsgBox("Выбранный примитив не является Полилинией. Макро прервано.")
                    Exit Sub
                End If
                Dim traectoryPoly As Polyline = DirectCast(entPolyTr, Polyline)
                Dim acadNetTopologySuite As FuncNetTopologySuite = New FuncNetTopologySuite()
                Dim userLineString As LineString = acadNetTopologySuite.CreateLineString(traectoryPoly)
                'фильтруем блоки
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(Db, "A-Стойка металлическая сетчатого ограждения")
                Dim listPoint As List(Of Point) = listBlockReference.Values.ToList
                listPoint = CadastralTools.FuncNetTopologySuite.OrderPointsAlongLine(userLineString, listPoint)
                Dim SurfProject As TinSurface = Nothing
                Dim SurfCool As ObjectIdCollection = CivilDoc.GetSurfaceIds
                If SurfCool.Count > 0 Then
                    For Each idSurf As ObjectId In SurfCool
                        Dim Surf As Autodesk.Civil.DatabaseServices.TinSurface = idSurf.GetObject(OpenMode.ForRead)
                        SurfProject = Surf
                        Exit For
                    Next
                End If
                Dim currentDist As Double = 0.0
                If listPoint.Count > 1 Then
                    For i As Integer = 0 To listPoint.Count - 2
                        Dim firstPoint As Point = listPoint.Item(i)
                        Dim elevPoint1 As Double = SurfProject.FindElevationAtXY(firstPoint.X, firstPoint.Y) + 2.45
                        Dim secondPoint As Point = listPoint.Item(i + 1)
                        Dim elevPoint2 As Double = SurfProject.FindElevationAtXY(firstPoint.X, firstPoint.Y) + 2.45
                        Dim deltaH As Double = elevPoint2 - elevPoint1
                        If deltaH > 0 Then
                            elevPoint1 += deltaH
                        Else
                            elevPoint2 += Math.Abs(deltaH)
                        End If

                        Dim dist As Double = CadastralTools.MathFunction.GetDistanceRoundedToMillimeter(firstPoint, secondPoint)
                        For j As Double = 0 To dist Step 0.25
                            Try
                                ' Точка на траектории
                                Dim point As Point = CadastralTools.MathFunction.GetPointAtDistance(firstPoint, secondPoint, j)
                                Dim pointOnCurve As Point3d = New Point3d(point.X, point.Y, elevPoint1)
                                ' Клон тела
                                Dim cloneSolid As Solid3d = DirectCast(userSolid.Clone(), Solid3d)
                                acBlkTblRec.AppendEntity(cloneSolid)
                                acTrans.AddNewlyCreatedDBObject(cloneSolid, True)

                                ' Перемещение: basePoint → pointOnCurve
                                Dim moveVec As Vector3d = pointOnCurve.GetVectorTo(basePoint).Negate()
                                cloneSolid.TransformBy(Matrix3d.Displacement(moveVec))

                                ' Разворот СТРОГО ПЕРПЕНДИКУЛЯРНО направлению траектории
                                Dim dir As Vector3d = traectoryPoly.GetFirstDerivative(pointOnCurve).GetNormal()
                                Dim angleAlong As Double = Math.Atan2(dir.Y, dir.X)
                                Dim anglePerp As Double = angleAlong + Math.PI / 2.0
                                cloneSolid.TransformBy(Matrix3d.Rotation(anglePerp, Vector3d.ZAxis, pointOnCurve))

                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                MsgBox("Ошибка при создании тела на расстоянии " & currentDist.ToString() & ": " & ex.Message)
                            End Try

                        Next j
                    Next i
                End If
                ' --- 6. Расстановка клонов вдоль направления ---
                Dim stepDist As Double = 1.0
                Dim totalLength As Double = traectoryPoly.Length


                While currentDist <= totalLength


                    currentDist += stepDist
                End While

                ' --- 7. Удаление вспомогательных объектов ---
                ' Копия полилинии
                If profilePoly IsNot Nothing AndAlso Not profilePoly.IsErased Then
                    profilePoly.Erase()
                End If

                ' Окружность профиля
                If acCircle IsNot Nothing AndAlso Not acCircle.IsErased Then
                    acCircle.Erase()
                End If

                ' Исходное тело (шаблон для клонирования)
                If userSolid IsNot Nothing AndAlso Not userSolid.IsErased Then
                    userSolid.Erase()
                End If

                acTrans.Commit()
            End Using
        End Using
    End Sub

End Class
