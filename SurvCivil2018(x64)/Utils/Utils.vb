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
Public Class Utils
    'сделать все По слою
    <CommandMethod("writeVertexPline")>
    Public Sub writeVertexPline()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim prPointRes1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию:")
                Dim ptIntRez As PromptIntegerResult = ed.GetInteger("Введите шаг разбивки кривой в мм: ")
                If ptIntRez.Value > 0 Then
                    Dim OshShag As Integer = ptIntRez.Value
                    If prPointRes1.Status = PromptStatus.OK Then
                        Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                        If IsDBNull(entObjID) = False Then
                            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForWrite)
                            If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Pline As Polyline = entObj
                                If Pline.Closed = True Then
                                    Dim boolConv As Boolean = MathFunction.FuncKadastrPlineDelDubleVertex(Pline)
line1:
                                    For i As Integer = 0 To Pline.NumberOfVertices - 1
                                        Dim userBulge As Double = Pline.GetBulgeAt(i)
                                        If userBulge <> 0 Then
                                            Dim seg As Autodesk.AutoCAD.Geometry.CircularArc2d = Pline.GetArcSegment2dAt(i)
                                            Dim stPoint3d As Point3d = New Point3d(seg.StartPoint.X, seg.StartPoint.Y, 0)
                                            Dim enPoint3d As Point3d = New Point3d(seg.EndPoint.X, seg.EndPoint.Y, 0)
                                            Dim startDist As Double = Pline.GetDistAtPoint(stPoint3d)
                                            Dim endDist As Double = Pline.GetDistAtPoint(enPoint3d)
                                            Dim dDist As Double = endDist - startDist
                                            If dDist > OshShag Then
                                                Dim sh As Double = dDist / OshShag
                                                sh = MathFunction.FuncTrimDigitCall(sh, 0)
                                                sh = dDist / sh
                                                For j As Double = startDist To endDist Step sh
                                                    If Math.Abs(j - startDist) <= 0.001 Then
                                                        i += 1
                                                        Continue For
                                                    End If
                                                    If Math.Abs(j - endDist) <= 0.001 Then
                                                        i += 1
                                                        Exit For
                                                    End If
                                                    Dim pt As Point3d = Pline.GetPointAtDist(j)
                                                    Pline.AddVertexAt(i, pt.Convert2d(New Plane), 0, 0, 0)
                                                    i += 1
                                                Next
                                            End If
                                        End If
                                    Next i
                                    Dim strFolder As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку")
                                    If IO.Directory.Exists(strFolder) = True Then
                                        Dim exportTxtVert As StreamWriter = New StreamWriter(strFolder & "\" & "vert" & ".txt", False, System.Text.Encoding.Default)
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim userBulge As Double = Pline.GetBulgeAt(i)
                                            If userBulge <> 0 Then
                                                Dim seg As Autodesk.AutoCAD.Geometry.CircularArc2d = Pline.GetArcSegment2dAt(i)
                                                Dim stPoint3d As Point3d = New Point3d(seg.StartPoint.X, seg.StartPoint.Y, 0)
                                                Dim enPoint3d As Point3d = New Point3d(seg.EndPoint.X, seg.EndPoint.Y, 0)
                                                Dim startDist As Double = Pline.GetDistAtPoint(stPoint3d)
                                                Dim endDist As Double = Pline.GetDistAtPoint(enPoint3d)
                                                Dim dDist As Double = endDist - startDist
                                                If dDist > OshShag * 2 Then
                                                    For j As Double = startDist To endDist Step OshShag
                                                        If Math.Abs(j - endDist) <= 0.01 Then
                                                            i += 1
                                                            Exit For
                                                        End If
                                                        Dim pt As Point3d = Pline.GetPointAtDist(j)
                                                        Pline.AddVertexAt(i, pt.Convert2d(New Plane), 0, 0, 0)
                                                        i += 1
                                                    Next
                                                End If
                                            Else
                                                Dim vert1 As Point2d = Pline.GetPoint2dAt(i)
                                                Dim vert2 As Point2d = Nothing
                                                If i = Pline.NumberOfVertices - 1 Then
                                                    vert2 = Pline.GetPoint2dAt(0)
                                                Else
                                                    vert2 = Pline.GetPoint2dAt(i + 1)
                                                End If
                                                Dim x1 As Double = Math.Round(vert1.X, 3)
                                                Dim y1 As Double = Math.Round(vert1.Y, 3)
                                                Dim x2 As Double = Math.Round(vert2.X, 3)
                                                Dim y2 As Double = Math.Round(vert2.Y, 3)
                                            End If





                                            'exportTxtVert.WriteLine(x1 & " " & y1 & " " & x2 & " " & y2)
                                        Next i
                                        exportTxtVert.Close()
                                    End If
                                Else
                                    MsgBox("Выбранная полилиния не замкнута. Макро прервано!!!")
                                End If
                            End If
                        End If
                    End If
                End If
                tr.Commit()
            Catch ex As System.Exception
            End Try
        End Using
    End Sub

    Public Class XRefSelection
        <CommandMethod("SELECTXREFOBJECT")>
        Public Sub SelectXRefObject()
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Try
                ' Запрос выбора объекта
                Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект во внешней ссылке: ")
                If per.Status <> PromptStatus.OK Then Return

                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    ' Получаем выбранный объект
                    Dim objId As ObjectId = per.ObjectId

                    ' Проверяем, находится ли объект в XRef
                    If IsObjectInXRef(objId, tr) Then
                        ed.WriteMessage(vbLf & "Выбран объект во внешней ссылке!")

                        ' Получаем информацию об объекте
                        Dim obj As Entity = tr.GetObject(objId, OpenMode.ForRead)
                        ed.WriteMessage(vbLf & "Тип объекта: " & obj.GetType().Name)

                        ' Подсвечиваем объект
                        obj.Highlight()
                    Else
                        ed.WriteMessage(vbLf & "Объект не находится во внешней ссылке!")
                    End If

                    tr.Commit()
                End Using

            Catch ex As System.Exception
                ed.WriteMessage(vbLf & "Ошибка: " & ex.Message)
            End Try
        End Sub

        Private Function IsObjectInXRef(objId As ObjectId, tr As Transaction) As Boolean
            Dim blkRef As BlockReference = TryCast(tr.GetObject(objId, OpenMode.ForRead), BlockReference)

            If blkRef IsNot Nothing Then
                Dim btr As BlockTableRecord = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead)

                ' Проверяем, является ли блок внешней ссылкой
                Return btr.IsFromExternalReference
            End If

            Return False
        End Function
    End Class

    ' Альтернативная команда с графическим выбором
    <CommandMethod("DeleteInBlockGraphic")>
    Public Sub DeleteEntityInsideBlockReferenceGraphic()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        ' Выбор BlockReference
        Dim blockRefId As ObjectId = SelectBlockReference(ed)
        If blockRefId.IsNull Then
            Return
        End If

        Dim idsEnt As List(Of ObjectId) = New List(Of ObjectId)
        Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию: ")
        If rs.Status = PromptStatus.OK Then
            Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
            ids.Reverse()
            ids.Add(rs.ObjectId)
            For Each id As ObjectId In ids
                Dim idStr As String = id.ObjectClass.Name
                If idStr Like "AcDbLine" Then
                    idsEnt.Add(id)
                End If
            Next
        End If
        If idsEnt.Count > 0 Then
            DeleteEntityFromBlock(db, blockRefId, idsEnt(0))
        End If
    End Sub
    <CommandMethod("DeleteInBlock")>
    Public Sub DeleteEntityInsideBlockReference()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database

        Try
            ' Шаг 1: Выбор BlockReference
            Dim blockRefId As ObjectId = SelectBlockReference(ed)
            If blockRefId.IsNull Then
                ed.WriteMessage(vbCrLf & "BlockReference не выбран.")
                Return
            End If

            ' Шаг 2: Выбор примитива внутри BlockReference
            Dim entityId As ObjectId = SelectEntityInBlock(ed, blockRefId)
            If entityId.IsNull Then
                ed.WriteMessage(vbCrLf & "Примитив не выбран.")
                Return
            End If

            ' Шаг 3: Подтверждение удаления
            Dim result As PromptKeywordOptions = New PromptKeywordOptions(vbCrLf & "Удалить выбранный примитив? [Да/Нет]")
            result.Keywords.Add("Да")
            result.Keywords.Add("Нет")
            result.Keywords.Default = "Нет"

            Dim keywordResult As PromptResult = ed.GetKeywords(result)
            If keywordResult.StringResult = "Да" Then
                ' Шаг 4: Удаление примитива
                If DeleteEntityFromBlock(db, blockRefId, entityId) Then
                    ed.WriteMessage(vbCrLf & "Примитив успешно удален из BlockReference.")
                Else
                    ed.WriteMessage(vbCrLf & "Ошибка при удалении примитива.")
                End If
            Else
                ed.WriteMessage(vbCrLf & "Удаление отменено.")
            End If

        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            ed.WriteMessage(vbCrLf & "Ошибка: " & ex.Message)
        End Try
    End Sub

    Private Function SelectBlockReference(ed As Editor) As ObjectId
        Dim opts As PromptEntityOptions = New PromptEntityOptions(vbCrLf & "Выберите BlockReference: ")
        opts.SetRejectMessage(vbCrLf & "Выберите только BlockReference.")
        opts.AddAllowedClass(GetType(BlockReference), False)

        Dim result As PromptEntityResult = ed.GetEntity(opts)
        If result.Status = PromptStatus.OK Then
            Return result.ObjectId
        End If

        Return ObjectId.Null
    End Function
    Private Function SelectEntityInBlock(ed As Editor, blockRefId As ObjectId) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Получаем BlockReference
                Dim blockRef As BlockReference = trans.GetObject(blockRefId, OpenMode.ForRead)

                ' Получаем BlockTableRecord (определение блока)
                Dim btr As BlockTableRecord = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead)

                ' Выводим информацию о содержимом блока
                ed.WriteMessage(vbCrLf & "Содержимое блока '" & btr.Name & "':")

                Dim entityIds As New List(Of ObjectId)
                Dim entityInfo As New List(Of String)

                ' Собираем все примитивы в блоке
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim entity As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                        If entity IsNot Nothing Then
                            entityIds.Add(id)
                            entityInfo.Add(GetEntityInfo(entity))
                        End If
                    End If
                Next

                ' Если в блоке нет примитивов
                If entityIds.Count = 0 Then
                    ed.WriteMessage(vbCrLf & "Блок не содержит примитивов.")
                    Return ObjectId.Null
                End If

                ' Выводим список примитивов
                For i As Integer = 0 To entityIds.Count - 1
                    ed.WriteMessage(vbCrLf & $"{i + 1}. {entityInfo(i)}")
                Next

                ' Запрашиваем выбор примитива
                Dim opts As PromptIntegerOptions = New PromptIntegerOptions(vbCrLf & vbCrLf & "Введите номер примитива для удаления: ")
                opts.LowerLimit = 1
                opts.UpperLimit = entityIds.Count
                opts.UseDefaultValue = False

                Dim intResult As PromptIntegerResult = ed.GetInteger(opts)
                If intResult.Status = PromptStatus.OK Then
                    Dim selectedIndex As Integer = intResult.Value - 1

                    ' Показываем предварительный просмотр выбранного примитива
                    Dim selectedEntity As Entity = trans.GetObject(entityIds(selectedIndex), OpenMode.ForRead)
                    ed.WriteMessage(vbCrLf & "Выбран: " & GetEntityInfo(selectedEntity))

                    ' Подтверждение выбора
                    Dim confirmOpts As PromptKeywordOptions = New PromptKeywordOptions(vbCrLf & "Это правильный примитив? [Да/Нет]")
                    confirmOpts.Keywords.Add("Да")
                    confirmOpts.Keywords.Add("Нет")
                    confirmOpts.Keywords.Default = "Да"

                    Dim confirmResult As PromptResult = ed.GetKeywords(confirmOpts)
                    If confirmResult.StringResult = "Да" Then
                        Return entityIds(selectedIndex)
                    End If
                End If

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                ed.WriteMessage(vbCrLf & "Ошибка при выборе примитива: " & ex.Message)
            End Try
        End Using

        Return ObjectId.Null
    End Function

    Private Function DeleteEntityFromBlock(db As Database, blockRefId As ObjectId, entityId As ObjectId) As Boolean
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Получаем BlockReference
                Dim blockRef As BlockReference = trans.GetObject(blockRefId, OpenMode.ForRead)

                ' Получаем BlockTableRecord
                Dim btr As BlockTableRecord = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite)

                ' Получаем примитив для удаления
                Dim entityToDelete As Entity = trans.GetObject(entityId, OpenMode.ForWrite)

                ' Удаляем примитив
                entityToDelete.Erase()

                trans.Commit()
                Return True

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                trans.Abort()
                Return False
            End Try
        End Using
    End Function

    ' Функция для получения информации о примитиве
    Private Function GetEntityInfo(entity As Entity) As String
        Dim entityType As String = entity.GetType().Name

        ' Для разных типов примитивов добавляем специфическую информацию
        Select Case entityType
            Case "Line"
                Dim line As Line = DirectCast(entity, Line)
                Return $"Линия: Длина = {Math.Round(line.Length, 2)}"

            Case "Circle"
                Dim circle As Circle = DirectCast(entity, Circle)
                Return $"Круг: Радиус = {Math.Round(circle.Radius, 2)}"

            Case "Polyline"
                Dim pline As Polyline = DirectCast(entity, Polyline)
                Return $"Полилиния: Вершин = {pline.NumberOfVertices}"

            Case "DBText"
                Dim text As DBText = DirectCast(entity, DBText)
                Return $"Текст: '{text.TextString}'"

            Case "MText"
                Dim mtext As MText = DirectCast(entity, MText)
                Return $"Мультитекст: '{mtext.Contents}'"

            Case Else
                Return $"{entityType} (ID: {entity.ObjectId})"
        End Select
    End Function
End Class
