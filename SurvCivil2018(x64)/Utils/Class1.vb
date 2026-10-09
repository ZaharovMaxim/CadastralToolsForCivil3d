Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Collections.Generic
Imports System
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Drawing.Graphics
Imports System.Drawing
Imports System.IO
Imports System.String
Imports System.Numerics
Public Class Class1
    <CommandMethod("MLEADERBLOCK")>
    Public Sub CreateMLeaderWithBlock()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ' 1. Запрос имени блока
            Dim blockNameOpt As PromptResult = ed.GetString(vbLf & "Введите имя блока для вставки: ")
            If blockNameOpt.Status <> PromptStatus.OK Then Return

            Dim blockName As String = blockNameOpt.StringResult

            ' Проверка существования блока
            If Not BlockExists(db, blockName) Then
                ed.WriteMessage(vbLf & $"Блок '{blockName}' не найден в чертеже.")
                Return
            End If

            ' 2. Запрос точек для мультивыноски
            Dim points As New List(Of Point3d)()

            ' Первая точка (начало выноски)
            Dim firstPointOpt As PromptPointResult = ed.GetPoint(vbLf & "Укажите точку начала выноски: ")
            If firstPointOpt.Status <> PromptStatus.OK Then Return
            points.Add(firstPointOpt.Value)

            ' Вторая точка (сгиб выноски)
            Dim secondPointOpt As PromptPointResult = ed.GetPoint(vbLf & "Укажите вторую точку выноски: ")
            If secondPointOpt.Status <> PromptStatus.OK Then Return
            points.Add(secondPointOpt.Value)

            ' Третья точка (положение блока)
            Dim thirdPointOpt As PromptPointResult = ed.GetPoint(vbLf & "Укажите положение блока: ")
            If thirdPointOpt.Status <> PromptStatus.OK Then Return
            points.Add(thirdPointOpt.Value)

            ' 3. Создание мультивыноски с блоком
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Создаем новую мультивыноску
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btrb As BlockTableRecord = DirectCast(trans.GetObject(bt(blockName), OpenMode.ForRead), BlockTableRecord)


                Dim mleader As New MLeader()
                ' Устанавливаем стиль мультивыноски
                SetMLeaderStyle(mleader, db)
                ' Добавляем точки выноски
                Dim leaderIndex As Integer = mleader.AddLeader()
                Dim leaderLineIndex As Integer = mleader.AddLeaderLine(leaderIndex)
                mleader.AddFirstVertex(leaderLineIndex, points(0))
                mleader.AddLastVertex(leaderLineIndex, points(1))
                ' Создаем блок-контент вместо текста
                CreateBlockContentForMLeader(mleader, blockName, points.Last(), db, trans)

                ' Добавляем мультивыноску в пространство модели
                Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim mleaderId As ObjectId = btr.AppendEntity(mleader)
                trans.AddNewlyCreatedDBObject(mleader, True)

                ' Настройка свойств мультивыноски
                ConfigureMLeaderProperties(mleader)
                mleader.BlockContentId = btrb.Id
                trans.Commit()

                ed.WriteMessage(vbLf & "Мультивыноска с блоком создана успешно!")
            End Using

        Catch ex As autodesk.AutoCAD.Runtime.Exception
            ed.WriteMessage(vbLf & $"Ошибка: {ex.Message}")
        End Try
    End Sub

    Private Function BlockExists(db As Database, blockName As String) As Boolean
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Return bt.Has(blockName)
        End Using
    End Function

    Private Sub SetMLeaderStyle(mleader As MLeader, db As Database)
        ' Используем стандартный стиль или создаем новый
        mleader.ContentType = ContentType.BlockContent
        mleader.TextStyleId = db.Textstyle
        mleader.TextHeight = 2.5
        mleader.ArrowSize = 2.5
        mleader.LandingGap = 1.0

        ' Настройка цвета и слоя
        mleader.ColorIndex = 256
        mleader.Layer = "0"
    End Sub

    Private Sub CreateBlockContentForMLeader(mleader As MLeader, blockName As String,
                                            insertionPoint As Point3d, db As Database,
                                            trans As Transaction)
        ' Создаем блок-контент для мультивыноски
        Dim blockContent As New MText()

        ' Вместо текста вставляем блок
        ' Но в AutoCAD мультивыноска не поддерживает блоки напрямую как контент
        ' Поэтому создадим альтернативное решение:

        ' Создаем анонимный блок, содержащий наш блок
        CreateBlockContent(mleader, blockName, insertionPoint, db, trans)
    End Sub

    Private Sub CreateBlockContent(mleader As MLeader, blockName As String,
                                  insertionPoint As Point3d, db As Database,
                                  trans As Transaction)
        ' Метод 1: Создание отдельного блока и его связь с выноской
        Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

        ' Создаем вставку блока
        Dim blockRef As New BlockReference(insertionPoint, bt(blockName))
        blockRef.Layer = "0"

        btr.AppendEntity(blockRef)
        trans.AddNewlyCreatedDBObject(blockRef, True)

        ' Устанавливаем связь между мультивыноской и блоком через словарь расширенных данных
        SetXDataForMLeader(mleader, blockRef.ObjectId)
    End Sub

    Private Sub SetXDataForMLeader(mleader As MLeader, blockId As ObjectId)
        ' Добавляем расширенные данные для связи с блоком
        Dim xdata As New ResultBuffer()
        xdata.Add(New TypedValue(1001, "MLEADER_BLOCK"))
        xdata.Add(New TypedValue(1005, blockId.Handle.ToString()))
        'mleader.XData = xdata
    End Sub

    Private Sub ConfigureMLeaderProperties(mleader As MLeader)
        ' Настройка дополнительных свойств мультивыноски
        'mleader.ArrowSymbolId = 0 ' Используем стандартную стрелку
        mleader.TextAttachmentType = TextAttachmentType.AttachmentMiddle
        mleader.EnableDogleg = True
        mleader.DoglegLength = 5.0
        'mleader.SetBlockContentId(0) ' Без блочного контента

    End Sub


    Public Function CalculatePerpendicularPoint(ByVal p1 As Vector3, ByVal p2 As Vector3, ByVal p3 As Vector3, ByVal length As Double) As Vector3
        ' 1. Вычисляем вектора от p2 к p1 и от p2 к p3
        Dim vectorP2P1 As Vector3 = p1 - p2
        Dim vectorP2P3 As Vector3 = p3 - p2
        ' 2. Вычисляем нормаль плоскости, образованной этими векторами
        Dim normal As Vector3 = Vector3.Cross(vectorP2P1, vectorP2P3)
        ' 3. Проверяем, не являются ли точки коллинеарными
        If normal.LengthSquared() < 0.00001 Then
            ' Если точки коллинеарны, используем перпендикуляр к отрезку p1-p3
            Dim vectorP1P3 As Vector3 = p3 - p1
            Dim perpendicularDir As Vector3
            ' Находим произвольный перпендикулярный вектор
            If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitY)
            End If
            normal = Vector3.Normalize(perpendicularDir)
            If normal.X > 0 And normal.Z < 0 Then
                vectorP1P3 = p1 - p3
                ' Находим произвольный перпендикулярный вектор
                If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                    perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitY)
                End If
                normal = Vector3.Normalize(perpendicularDir)
            End If
        Else
            ' Нормализуем вектор нормали
            normal = Vector3.Normalize(normal)
        End If
        ' 4. Восстанавливаем перпендикуляр от точки p2
        Dim resultPoint As Vector3 = p2 + normal * CSng(length)
        Return resultPoint
    End Function

    Public Function CalculatePerpendicularPointSimple2(ByVal p1 As Vector3, ByVal p2 As Vector3, ByVal p3 As Vector3, ByVal length As Double) As Vector3
        ' 1. Вычисляем вектора от p2 к p1 и от p2 к p3
        Dim vectorP2P1 As Vector3 = p1 - p2
        Dim vectorP2P3 As Vector3 = p3 - p2

        ' 2. Вычисляем нормаль плоскости
        Dim normal As Vector3 = Vector3.Cross(vectorP2P1, vectorP2P3)

        ' 3. Проверяем коллинеарность
        If normal.LengthSquared() < 0.00001 Then
            ' Коллинеарный случай
            Dim vectorP1P3 As Vector3 = p3 - p1
            Dim perpendicularDir As Vector3

            ' Выбираем ось для векторного произведения
            If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitY)
            Else
                perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitX)
            End If

            normal = Vector3.Normalize(perpendicularDir)

            ' Корректируем направление для гарантии Z < 0
            If normal.Z > 0 Then
                normal = -normal
            End If

            ' Дополнительная проверка по X
            If normal.X > 0 And normal.Z < 0 Then
                vectorP1P3 = p1 - p3
                If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                    perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitY)
                Else
                    perpendicularDir = Vector3.Cross(vectorP1P3, Vector3.UnitX)
                End If
                normal = Vector3.Normalize(perpendicularDir)

                ' Снова проверяем и корректируем Z
                If normal.Z > 0 Then
                    normal = -normal
                End If
            End If
        Else
            ' Нормализуем вектор нормали
            normal = Vector3.Normalize(normal)

            ' Корректируем направление для гарантии Z < 0
            If normal.Z > 0 Then
                normal = -normal
            End If
        End If

        ' 4. Восстанавливаем перпендикуляр
        Dim resultPoint As Vector3 = p2 + normal * CSng(length)

        ' 5. Финальная проверка и коррекция
        If resultPoint.Z >= p2.Z Then
            ' Если точка не ниже, принудительно опускаем
            Dim verticalDown As Vector3 = -Vector3.UnitZ
            resultPoint = p2 + verticalDown * CSng(length)
        End If

        Return resultPoint
    End Function

    <CommandMethod("writeVertexPline1")>
    Public Sub writeVertexPline1()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Try
                Dim prPointRes1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию:")
                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                If IsDBNull(entObjID) = False Then
                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForWrite)
                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                        Dim Pline As Polyline3d = entObj
                        Dim listPoint As List(Of Vector3) = New List(Of Vector3)
                        For Each vertexId As ObjectId In Pline
                            Dim v3d As PolylineVertex3d = tr.GetObject(vertexId, OpenMode.ForRead)
                            listPoint.Add(New Vector3(v3d.Position.X, v3d.Position.Y, v3d.Position.Z))
                        Next
                        If listPoint.Count > 3 Then
                            Dim p1 As Vector3 = Nothing
                            Dim p2 As Vector3 = Nothing
                            Dim p3 As Vector3 = Nothing
                            Dim newPline3d As Polyline3d = New Polyline3d
                            Dim polylineId As ObjectId = btr.AppendEntity(newPline3d)
                            tr.AddNewlyCreatedDBObject(newPline3d, True)
                            For i As Integer = 0 To listPoint.Count - 1
                                If i = 0 Then
                                    p1 = listPoint.Last
                                    p2 = listPoint(i)
                                    p3 = listPoint(i + 1)
                                ElseIf i = listPoint.Count - 1 Then
                                    p1 = listPoint(i - 1)
                                    p2 = listPoint(i)
                                    p3 = listPoint(0)
                                Else
                                    p1 = listPoint(i - 1)
                                    p2 = listPoint(i)
                                    p3 = listPoint(i + 1)
                                End If
                                Dim pt3d As Vector3 = CalculatePerpendicularPointSimple2(p1, p2, p3, 0.2)
                                Dim vertex As New PolylineVertex3d(New Point3d(pt3d.X, pt3d.Y, pt3d.Z))
                                newPline3d.AppendVertex(vertex)
                                tr.AddNewlyCreatedDBObject(vertex, True)
                            Next i
                            newPline3d.Closed = True
                        End If
                    End If
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub
End Class

Public Class IntersectionResult
    Public Property IntersectionPoint As Vector3?
    Public Property IsParallel As Boolean
    Public Property IsWithinSegment As Boolean
    Public Property DistanceFromStart As Single
    Public Property IntersectionParameter As Single
    Public Property ErrorMessage As String


    Public Function IntersectPlaneLineDetailed(
    ByVal planeP1 As Vector3,
    ByVal planeP2 As Vector3,
    ByVal planeP3 As Vector3,
    ByVal lineStart As Vector3,
    ByVal lineEnd As Vector3) As IntersectionResult

        Dim result As New IntersectionResult()

        ' 1. Вычисляем нормаль плоскости
        Dim vector1 As Vector3 = planeP2 - planeP1
        Dim vector2 As Vector3 = planeP3 - planeP1
        Dim planeNormal As Vector3 = Vector3.Cross(vector1, vector2)

        ' 2. Проверка определения плоскости
        If planeNormal.LengthSquared() < 0.00001 Then
            result.ErrorMessage = "Точки плоскости коллинеарны"
            Return result
        End If

        planeNormal = Vector3.Normalize(planeNormal)

        ' 3. Уравнение плоскости: n·(r - r0) = 0, где n - нормаль, r0 - точка на плоскости
        ' Или: n·r - n·r0 = 0 => n·r + D = 0, где D = -n·r0
        Dim D As Single = -Vector3.Dot(planeNormal, planeP1)

        ' 4. Параметрическое уравнение линии: r = p + t * d
        ' где p = lineStart, d = lineDirection
        Dim lineDirection As Vector3 = lineEnd - lineStart
        Dim lineLength As Single = lineDirection.Length()

        If lineLength < 0.00001 Then
            result.ErrorMessage = "Начало и конец линии совпадают"
            Return result
        End If

        ' 5. Проверка параллельности
        Dim denominator As Single = Vector3.Dot(planeNormal, lineDirection)

        If Math.Abs(denominator) < 0.00001 Then
            result.IsParallel = True

            ' Проверяем, лежит ли линия в плоскости
            Dim distanceToPlane As Single = Vector3.Dot(planeNormal, lineStart) + D
            If Math.Abs(distanceToPlane) < 0.00001 Then
                result.ErrorMessage = "Линия лежит в плоскости"
            Else
                result.ErrorMessage = "Линия параллельна плоскости, но не лежит в ней"
            End If

            Return result
        End If

        ' 6. Вычисляем параметр t
        Dim numerator As Single = -(Vector3.Dot(planeNormal, lineStart) + D)
        Dim t As Single = numerator / denominator

        result.IntersectionParameter = t
        result.DistanceFromStart = t * lineLength

        ' 7. Вычисляем точку пересечения
        Dim intersectionPoint As New Vector3(
        lineStart.X + t * lineDirection.X,
        lineStart.Y + t * lineDirection.Y,
        lineStart.Z + t * lineDirection.Z
    )

        result.IntersectionPoint = intersectionPoint

        ' 8. Проверяем, находится ли точка на отрезке
        result.IsWithinSegment = (t >= 0 AndAlso t <= 1)

        Return result
    End Function

    Public Function IsPointInTriangleBarycentric(
    ByVal p As Vector3,
    ByVal a As Vector3,
    ByVal b As Vector3,
    ByVal c As Vector3,
    Optional tolerance As Single = 0.0001F) As Boolean

        ' 1. Проектируем все точки на 2D плоскость для упрощения вычислений
        ' Выбираем наиболее стабильную проекцию
        Dim axisToDrop As Integer = GetDominantAxis(a, b, c)

        ' 2. Создаем 2D проекции точек
        Dim p2d As Vector2 = ProjectTo2D(p, axisToDrop)
        Dim a2d As Vector2 = ProjectTo2D(a, axisToDrop)
        Dim b2d As Vector2 = ProjectTo2D(b, axisToDrop)
        Dim c2d As Vector2 = ProjectTo2D(c, axisToDrop)

        ' 3. Вычисляем барицентрические координаты
        Dim v0 As Vector2 = b2d - a2d
        Dim v1 As Vector2 = c2d - a2d
        Dim v2 As Vector2 = p2d - a2d

        Dim dot00 As Single = Vector2.Dot(v0, v0)
        Dim dot01 As Single = Vector2.Dot(v0, v1)
        Dim dot11 As Single = Vector2.Dot(v1, v1)
        Dim dot20 As Single = Vector2.Dot(v2, v0)
        Dim dot21 As Single = Vector2.Dot(v2, v1)

        Dim invDenom As Single = 1.0F / (dot00 * dot11 - dot01 * dot01)

        ' Барицентрические координаты u и v
        Dim u As Single = (dot11 * dot20 - dot01 * dot21) * invDenom
        Dim v As Single = (dot00 * dot21 - dot01 * dot20) * invDenom

        ' 4. Проверка условий
        Return (u >= -tolerance) AndAlso (v >= -tolerance) AndAlso (u + v <= 1.0F + tolerance)
    End Function

    Private Function GetDominantAxis(a As Vector3, b As Vector3, c As Vector3) As Integer
        ' Находим ось с наибольшим изменением для стабильной проекции
        Dim normal As Vector3 = Vector3.Cross(b - a, c - a)

        ' Выбираем ось с наибольшим абсолютным значением в нормали
        If Math.Abs(normal.X) >= Math.Abs(normal.Y) AndAlso Math.Abs(normal.X) >= Math.Abs(normal.Z) Then
            Return 0 ' Отбрасываем X
        ElseIf Math.Abs(normal.Y) >= Math.Abs(normal.X) AndAlso Math.Abs(normal.Y) >= Math.Abs(normal.Z) Then
            Return 1 ' Отбрасываем Y
        Else
            Return 2 ' Отбрасываем Z
        End If
    End Function

    Private Function ProjectTo2D(point As Vector3, axisToDrop As Integer) As Vector2
        Select Case axisToDrop
            Case 0 ' Отбрасываем X
                Return New Vector2(point.Y, point.Z)
            Case 1 ' Отбрасываем Y
                Return New Vector2(point.X, point.Z)
            Case Else ' Отбрасываем Z
                Return New Vector2(point.X, point.Y)
        End Select
    End Function
    <CommandMethod("writeVertexPline2")>
    Public Sub writeVertexPline2()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Try
                Dim lstart As Vector3 = New Vector3
                Dim lend As Vector3 = New Vector3
                Dim prPointRes0 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите линию:")
                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes0.ObjectId
                If IsDBNull(entObjID) = False Then
                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForWrite)
                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Line Then
                        Dim Pline As Line = entObj
                        lstart = New Vector3(Pline.StartPoint.X, Pline.StartPoint.Y, Pline.StartPoint.Z)
                        lend = New Vector3(Pline.EndPoint.X, Pline.EndPoint.Y, Pline.EndPoint.Z)
                    End If
                End If
                Dim prPointRes1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите плоскость:")
                Dim entObjID1 As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                If IsDBNull(entObjID1) = False Then
                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID1, OpenMode.ForWrite)
                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                        Dim Pline As Polyline3d = entObj
                        Dim listPoint As List(Of Vector3) = New List(Of Vector3)
                        For Each vertexId As ObjectId In Pline
                            Dim v3d As PolylineVertex3d = tr.GetObject(vertexId, OpenMode.ForRead)
                            listPoint.Add(New Vector3(v3d.Position.X, v3d.Position.Y, v3d.Position.Z))
                        Next
                        If listPoint.Count > 3 Then
                            Dim p1 As Vector3 = Nothing
                            Dim p2 As Vector3 = Nothing
                            Dim p3 As Vector3 = Nothing
                            Dim newPline3d As Polyline3d = New Polyline3d
                            Dim polylineId As ObjectId = btr.AppendEntity(newPline3d)
                            tr.AddNewlyCreatedDBObject(newPline3d, True)
                            For i As Integer = 0 To listPoint.Count - 1
                                If i = 0 Then
                                    p1 = listPoint.Last
                                    p2 = listPoint(i)
                                    p3 = listPoint(i + 1)
                                ElseIf i = listPoint.Count - 1 Then
                                    p1 = listPoint(i - 1)
                                    p2 = listPoint(i)
                                    p3 = listPoint(0)
                                Else
                                    p1 = listPoint(i - 1)
                                    p2 = listPoint(i)
                                    p3 = listPoint(i + 1)
                                End If
                                Dim pt3d As IntersectionResult = IntersectPlaneLineDetailed(p1, p2, p3, lstart, lend)
                                Dim boolFind As Boolean = IsPointInTriangleBarycentric(p1, p2, p3, pt3d.IntersectionPoint.Value)
                            Next i
                            newPline3d.Closed = True
                        End If
                    End If
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub


    Public Function CreatePerpendicularAtPoint(
    ByVal lineStart As Vector3,
    ByVal lineEnd As Vector3,
    ByVal pointOnLine As Vector3,
    ByVal perpendicularLength As Single,
    Optional ByVal useRightHandRule As Boolean = True) As Vector3()

        ' Создает перпендикуляр к линии в заданной точке

        ' Вычисляем направление линии
        Dim lineDirection As Vector3 = lineEnd - lineStart

        If lineDirection.LengthSquared() < 0.00001 Then
            ' Точки совпадают
            Return {pointOnLine, pointOnLine}
        End If

        lineDirection = Vector3.Normalize(lineDirection)

        ' Находим перпендикулярное направление
        Dim perpendicularDir As Vector3

        If useRightHandRule Then
            ' Используем правило правой руки
            perpendicularDir = Vector3.Cross(lineDirection, Vector3.UnitY)

            ' Если получился нулевой вектор, пробуем другую ось
            If perpendicularDir.LengthSquared() < 0.00001 Then
                perpendicularDir = Vector3.Cross(lineDirection, Vector3.UnitZ)
            End If
        Else
            ' Используем правило левой руки
            perpendicularDir = Vector3.Cross(Vector3.UnitY, lineDirection)

            If perpendicularDir.LengthSquared() < 0.00001 Then
                perpendicularDir = Vector3.Cross(Vector3.UnitZ, lineDirection)
            End If
        End If

        ' Нормализуем перпендикуляр
        If perpendicularDir.LengthSquared() < 0.00001 Then
            ' Не удалось найти перпендикуляр
            Return {pointOnLine, pointOnLine}
        End If

        perpendicularDir = Vector3.Normalize(perpendicularDir)

        ' Создаем точки перпендикуляра
        Dim perpendicularPoint1 As Vector3 = pointOnLine + perpendicularDir * perpendicularLength
        Dim perpendicularPoint2 As Vector3 = pointOnLine - perpendicularDir * perpendicularLength

        Return {perpendicularPoint1, perpendicularPoint2}
    End Function

    Public Enum PerpendicularPlane
        Vertical    ' Вертикальный перпендикуляр (по умолчанию)
        Horizontal  ' Горизонтальный перпендикуляр
        Any         ' Любой перпендикуляр (старое поведение)
    End Enum

    Public Function CreatePerpendicularAtPoint2(
    ByVal lineStart As Vector3,
    ByVal lineEnd As Vector3,
    ByVal pointOnLine As Vector3,
    ByVal perpendicularLength As Single,
    Optional ByVal useRightHandRule As Boolean = True,
    Optional ByVal plane As PerpendicularPlane = PerpendicularPlane.Vertical,
    Optional ByVal referenceVector As Vector3 = Nothing) As Vector3()

        ' Создает перпендикуляр к линии в заданной точке
        ' plane: определяет плоскость, в которой лежит перпендикуляр
        ' referenceVector: опорный вектор для определения ориентации

        ' Вычисляем направление линии
        Dim lineDirection As Vector3 = lineEnd - lineStart

        If lineDirection.LengthSquared() < 0.00001 Then
            ' Точки совпадают
            Return {pointOnLine, pointOnLine}
        End If

        lineDirection = Vector3.Normalize(lineDirection)

        ' Находим перпендикулярное направление в зависимости от выбранной плоскости
        Dim perpendicularDir As Vector3 = Vector3.Zero

        Select Case plane
            Case PerpendicularPlane.Vertical
                ' ВЕРТИКАЛЬНЫЙ перпендикуляр (сохраняет вертикальность)
                perpendicularDir = GetVerticalPerpendicular(lineDirection, useRightHandRule)

            Case PerpendicularPlane.Horizontal
                ' ГОРИЗОНТАЛЬНЫЙ перпендикуляр (лежит в горизонтальной плоскости)
                perpendicularDir = GetHorizontalPerpendicular(lineDirection, pointOnLine, useRightHandRule, referenceVector)

            Case PerpendicularPlane.Any
                ' ЛЮБОЙ перпендикуляр (старое поведение)
                perpendicularDir = GetAnyPerpendicular(lineDirection, useRightHandRule)

        End Select

        ' Нормализуем перпендикуляр
        If perpendicularDir.LengthSquared() < 0.00001 Then
            ' Не удалось найти перпендикуляр
            Return {pointOnLine, pointOnLine}
        End If

        perpendicularDir = Vector3.Normalize(perpendicularDir)

        ' Создаем точки перпендикуляра
        Dim perpendicularPoint1 As Vector3 = pointOnLine + perpendicularDir * perpendicularLength
        Dim perpendicularPoint2 As Vector3 = pointOnLine - perpendicularDir * perpendicularLength

        Return {perpendicularPoint1, perpendicularPoint2}
    End Function

    Private Function GetVerticalPerpendicular(
    ByVal lineDirection As Vector3,
    ByVal useRightHandRule As Boolean) As Vector3

        ' ВЕРТИКАЛЬНЫЙ перпендикуляр: перпендикулярен линии и сохраняет вертикальность
        ' (имеет нулевую проекцию на горизонтальную плоскость)

        ' Проекция направления линии на горизонтальную плоскость
        Dim horizontalDirection As New Vector2(lineDirection.X, lineDirection.Y)

        If horizontalDirection.LengthSquared() < 0.00001 Then
            ' Линия вертикальна - используем ось X для перпендикуляра
            Return If(useRightHandRule, Vector3.UnitX, -Vector3.UnitX)
        End If

        ' Нормализуем горизонтальное направление
        horizontalDirection = Vector2.Normalize(horizontalDirection)

        ' Перпендикуляр в горизонтальной плоскости
        Dim horizontalPerpendicular As Vector2

        If useRightHandRule Then
            ' Вправо: (-Y, X)
            horizontalPerpendicular = New Vector2(-horizontalDirection.Y, horizontalDirection.X)
        Else
            ' Влево: (Y, -X)
            horizontalPerpendicular = New Vector2(horizontalDirection.Y, -horizontalDirection.X)
        End If

        ' Создаем 3D вектор с нулевым Z
        Return New Vector3(horizontalPerpendicular.X, horizontalPerpendicular.Y, 0)
    End Function

    Private Function GetHorizontalPerpendicular(
    ByVal lineDirection As Vector3,
    ByVal pointOnLine As Vector3,
    ByVal useRightHandRule As Boolean,
    ByVal referenceVector As Vector3) As Vector3

        ' ГОРИЗОНТАЛЬНЫЙ перпендикуляр: лежит в горизонтальной плоскости
        ' (перпендикулярен проекции линии на горизонтальную плоскость)

        ' 1. Проекция линии на горизонтальную плоскость (игнорируем Z)
        Dim horizontalLineDir As New Vector2(lineDirection.X, lineDirection.Y)

        If horizontalLineDir.LengthSquared() < 0.00001 Then
            ' Линия вертикальна в XY - перпендикуляр может быть любым горизонтальным вектором
            ' Используем referenceVector или ось X по умолчанию
            If referenceVector.LengthSquared() > 0.00001 Then
                Dim refHorizontal As New Vector2(referenceVector.X, referenceVector.Y)
                If refHorizontal.LengthSquared() > 0.00001 Then
                    refHorizontal = Vector2.Normalize(refHorizontal)
                    Return New Vector3(refHorizontal.X, refHorizontal.Y, 0)
                End If
            End If
            Return New Vector3(1, 0, 0) ' Ось X по умолчанию
        End If

        ' 2. Нормализуем горизонтальное направление линии
        horizontalLineDir = Vector2.Normalize(horizontalLineDir)

        ' 3. Вычисляем перпендикуляр в горизонтальной плоскости
        Dim horizontalPerpendicular As Vector2

        If useRightHandRule Then
            ' Вправо: (-Y, X)
            horizontalPerpendicular = New Vector2(-horizontalLineDir.Y, horizontalLineDir.X)
        Else
            ' Влево: (Y, -X)
            horizontalPerpendicular = New Vector2(horizontalLineDir.Y, -horizontalLineDir.X)
        End If

        ' 4. Создаем 3D вектор (Z = 0, так как горизонтальный)
        Return New Vector3(horizontalPerpendicular.X, horizontalPerpendicular.Y, 0)
    End Function

    Private Function GetAnyPerpendicular(
    ByVal lineDirection As Vector3,
    ByVal useRightHandRule As Boolean) As Vector3

        ' ЛЮБОЙ перпендикуляр (старая логика)

        If useRightHandRule Then
            ' Используем правило правой руки
            Dim perpendicularDir As Vector3 = Vector3.Cross(lineDirection, Vector3.UnitY)

            ' Если получился нулевой вектор, пробуем другую ось
            If perpendicularDir.LengthSquared() < 0.00001 Then
                perpendicularDir = Vector3.Cross(lineDirection, Vector3.UnitZ)
            End If

            Return perpendicularDir
        Else
            ' Используем правило левой руки
            Dim perpendicularDir As Vector3 = Vector3.Cross(Vector3.UnitY, lineDirection)

            If perpendicularDir.LengthSquared() < 0.00001 Then
                perpendicularDir = Vector3.Cross(Vector3.UnitZ, lineDirection)
            End If

            Return perpendicularDir
        End If
    End Function

    ' Дополнительная функция для получения перпендикуляра в произвольной плоскости
    Public Function GetPerpendicularInPlane(
    ByVal lineDirection As Vector3,
    ByVal planeNormal As Vector3,
    Optional ByVal useRightHandRule As Boolean = True) As Vector3

        ' Возвращает перпендикуляр к линии, лежащий в заданной плоскости

        ' Нормализуем векторы
        lineDirection = Vector3.Normalize(lineDirection)
        planeNormal = Vector3.Normalize(planeNormal)

        ' Проверяем, что линия не параллельна нормали плоскости
        Dim dot As Single = Vector3.Dot(lineDirection, planeNormal)
        If Math.Abs(dot) > 0.999 Then
            ' Линия перпендикулярна плоскости - любой вектор в плоскости будет перпендикулярен
            Return GetAnyPerpendicularInPlane(planeNormal)
        End If

        ' Вычисляем перпендикуляр через двойное векторное произведение
        ' Формула: (planeNormal × lineDirection) × lineDirection
        Dim cross1 As Vector3 = Vector3.Cross(planeNormal, lineDirection)
        Dim perpendicular As Vector3 = Vector3.Cross(cross1, lineDirection)

        ' Нормализуем и корректируем направление
        perpendicular = Vector3.Normalize(perpendicular)

        ' Проверяем ориентацию
        If Not useRightHandRule Then
            perpendicular = -perpendicular
        End If

        Return perpendicular
    End Function

    Private Function GetAnyPerpendicularInPlane(ByVal planeNormal As Vector3) As Vector3
        ' Находит любой вектор в заданной плоскости

        ' Пробуем разные базисные векторы
        Dim candidates As Vector3() = {Vector3.UnitX, Vector3.UnitY, Vector3.UnitZ}

        For Each candidate As Vector3 In candidates
            Dim cross As Vector3 = Vector3.Cross(planeNormal, candidate)
            If cross.LengthSquared() > 0.00001 Then
                Return Vector3.Normalize(cross)
            End If
        Next

        ' Если все кандидаты не подошли
        Return Vector3.UnitX
    End Function





    <CommandMethod("writeVertexPline3")>
    Public Sub writeVertexPline3()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Try
                Dim lstart As Vector3 = New Vector3
                Dim lend As Vector3 = New Vector3
                Dim prPointRes0 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите линию:")
                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes0.ObjectId
                If IsDBNull(entObjID) = False Then
                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForWrite)
                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Line Then
                        Dim Pline As Line = entObj
                        lstart = New Vector3(Pline.StartPoint.X, Pline.StartPoint.Y, Pline.StartPoint.Z)
                        lend = New Vector3(Pline.EndPoint.X, Pline.EndPoint.Y, Pline.EndPoint.Z)
                    End If
                End If
                Dim Point As Vector3() = CreatePerpendicularAtPoint2(lstart, lend, lstart, 5, True, PerpendicularPlane.Vertical)
                Dim Point2 As Vector3() = CreatePerpendicularAtPoint2(lstart, lend, lstart, 5, True, PerpendicularPlane.Horizontal)
                MsgBox("   Результат: X=")
                'MsgBox("   Результат: X=" & Point.X & ", Y=" & Point.Y & ", Z=" & Point.Z)

            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub










End Class