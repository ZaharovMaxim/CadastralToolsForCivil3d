Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports System.Text.RegularExpressions
Public Class ConvertPlane
    'сделать все По слою
    <CommandMethod("makeDrawingGray")>
    Public Sub makeDrawingGray()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
        Dim dictLayer As Dictionary(Of String, Integer) = AcadFunction.FuncGetLayerColors()
        Dim nameObject As String = ""
        Using tr
            Try
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            nameObject = acEnt.GetType.ToString & " handle " & acEnt.Handle.ToString
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is MText Then
                                    Dim acMtext As MText = acEnt
                                    Dim str As String = acMtext.Contents
                                    acMtext.Contents = Regex.Replace(acMtext.Contents, "\\C\d+;", MText.ColorChange & "256;")
                                    acMtext.Contents = Regex.Replace(acMtext.Contents, "\\c\d+;", MText.ColorChange & "256;")
                                    Dim pt As Point3d = acMtext.Location
                                    If pt.Z <> 0 Then
                                        acMtext.Location = New Point3d(pt.X, pt.Y, 0)
                                    End If
                                ElseIf TypeOf acEnt Is DBText Then
                                    Dim text As DBText = acEnt
                                    Dim pt As Point3d = text.Position
                                    If pt.Z <> 0 Then
                                        text.Position = New Point3d(pt.X, pt.Y, 0)
                                    End If
                                ElseIf TypeOf acEnt Is line Then
                                    Dim acLine As Line = acEnt
                                    If acLine.StartPoint.Z <> 0 Then
                                        Dim newPt As Point3d = New Point3d(acLine.StartPoint.X, acLine.StartPoint.Y, 0)
                                        acLine.StartPoint = newPt
                                    End If
                                    If acLine.EndPoint.Z <> 0 Then
                                        Dim newPt As Point3d = New Point3d(acLine.EndPoint.X, acLine.EndPoint.Y, 0)
                                        acLine.EndPoint = newPt
                                    End If
                                ElseIf TypeOf acEnt Is Polyline Then
                                    Dim acPline As Polyline = acEnt
                                    If acPline.Elevation <> 0 Then
                                        acPline.Elevation = 0
                                    End If
                                ElseIf TypeOf acEnt Is Polyline2d Then
                                    Dim acPline As Polyline2d = acEnt
                                    If acPline.Elevation <> 0 Then
                                        acPline.Elevation = 0
                                    End If
                                ElseIf TypeOf acEnt Is Circle Then
                                    Dim acCirc As Circle = acEnt
                                    If acCirc.Center.Z <> 0 Then
                                        Dim newPt As Point3d = New Point3d(acCirc.StartPoint.X, acCirc.StartPoint.Y, 0)
                                        acCirc.Center = newPt
                                    End If
                                ElseIf TypeOf acEnt Is Hatch Then
                                    Dim acHatch As Hatch = acEnt
                                    If acHatch.Elevation <> 0 Then
                                        acHatch.Elevation = 0
                                    End If
                                ElseIf TypeOf acEnt Is DBPoint Then
                                    Dim acPoint As DBPoint = acEnt
                                    If acPoint.Position.Z <> 0 Then
                                        Dim newPt As Point3d = New Point3d(acPoint.Position.X, acPoint.Position.Y, 0)
                                        acPoint.Position = newPt
                                    End If
                                ElseIf TypeOf acEnt Is BlockReference Then
                                    Dim blk As BlockReference = acEnt
                                    If blk.AttributeCollection.Count > 0 Then
                                        For Each attId As ObjectId In blk.AttributeCollection
                                            Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForWrite)
                                            attRef.ColorIndex = 256
                                        Next
                                    End If
                                    If blk.Position.Z <> 0 Then
                                        Dim newPt As Point3d = New Point3d(blk.Position.X, blk.Position.Y, 0)
                                        blk.Position = newPt
                                    End If
                                End If
                                    acEnt.ColorIndex = 256
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
                MsgBox(ex.Message & " " & nameObject)
            End Try
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'перебираем блоки
            Dim bt = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            For Each btrId As ObjectId In bt
                Dim btr = CType(tr.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                If btr.IsLayout Then Continue For
                For Each id As ObjectId In btr
                    Try
                        If id.IsNull = False Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = id.GetObject(OpenMode.ForWrite)
                            If TypeOf acEnt Is Hatch Then
                                acEnt.Layer = "0"
                            Else
                                acEnt.ColorIndex = 256
                                acEnt.Layer = "0"
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
            Next
            tr.Commit()
        End Using
    End Sub
    'сделать все По слою
    <CommandMethod("ReplaseTextToMtext")>
    Public Sub ReplaseTextToMtext()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'сохраняем документ
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
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Try
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is DBText Then
                                    Dim acText As DBText = acEnt
                                    Dim atPoint As AttachmentPoint = acText.Justify

                                    Dim acNewMText As MText = New MText()
                                    acNewMText.SetDatabaseDefaults()
                                    acNewMText.Contents = acText.TextString
                                    acNewMText.Location = acText.Position
                                    acNewMText.Attachment = atPoint
                                    acNewMText.Layer = acText.Layer
                                    acNewMText.Color = acText.Color
                                    acNewMText.BackgroundFill = True
                                    acNewMText.BackgroundScaleFactor = 1
                                    Dim acBound As Extents3d = acText.Bounds
                                    Dim ach As Double = acBound.MaxPoint.X - acBound.MinPoint.X
                                    Dim acw As Double = acBound.MaxPoint.Y - acBound.MinPoint.Y
                                    acNewMText.Height = acText.Height
                                    acNewMText.Rotation = acText.Rotation
                                    acNewMText.TextHeight = acText.Height
                                    acNewMText.TextStyleId = acText.TextStyleId
                                    acNewMText.Width = acw
                                    acNewMText.Height = ach
                                    acBlkTblRec.AppendEntity(acNewMText)
                                    tr.AddNewlyCreatedDBObject(acNewMText, True)
                                    acText.Erase()
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub

    <CommandMethod("convReg")>
    Public Sub convReg()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Try
                '=============================================================
                'выбираем трассу
                Dim pEntRes As PromptEntityResult = ed.GetEntity("Выберите полосу отвода: ")
                Dim parentEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pEntRes.ObjectId, OpenMode.ForRead)
                If TypeOf parentEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim parentPline As Autodesk.AutoCAD.DatabaseServices.Polyline = parentEnt
                    Dim tempCollection As DBObjectCollection = New DBObjectCollection()
                    tempCollection.Add(parentPline)
                    Dim parentRegionCollection As DBObjectCollection = New DBObjectCollection()
                    parentRegionCollection = Region.CreateFromCurves(tempCollection)
                    If parentRegionCollection.Count > 0 Then
                        Dim parentRegion As Region = parentRegionCollection.Item(0)
                        If parentRegion.Area > 0 Then
                            'выбираем участки
                            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                            If acSSPrompt.Status = PromptStatus.OK Then
                                Dim acSSet As SelectionSet = acSSPrompt.Value
                                For Each acSSObj As SelectedObject In acSSet
                                    If Not IsDBNull(acSSObj) Then
                                        Dim childEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                        If Not IsDBNull(childEnt) Then
                                            If TypeOf childEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                If parentRegion.Area = 0 Then
                                                    tempCollection = New DBObjectCollection()
                                                    tempCollection.Add(parentPline)
                                                    parentRegionCollection = New DBObjectCollection()
                                                    parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                                    If parentRegionCollection.Count > 0 Then
                                                        parentRegion = parentRegionCollection.Item(0)
                                                    End If
                                                End If
                                                Dim tempChildCollection As DBObjectCollection = New DBObjectCollection()
                                                tempChildCollection.Add(childEnt)
                                                Dim childCollection As DBObjectCollection = New DBObjectCollection()
                                                childCollection = Region.CreateFromCurves(tempChildCollection)
                                                If childCollection.Count > 0 Then
                                                    Dim childRegion As Region = childCollection.Item(0)
                                                    If childRegion.Area > 0 Then
                                                        childRegion.BooleanOperation(BooleanOperationType.BoolIntersect, parentRegion)
                                                        Dim points As Point2d() = Nothing
                                                        Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(childRegion)
                                                            points = brep.Edges.[Select](Function(e) e.Vertex1.Point.Convert2d(New Plane)).ToArray()
                                                        End Using
                                                        Dim normal = childRegion.Normal
                                                        Using pline = New Polyline(points.Length)
                                                            For i As Integer = 0 To points.Length - 1
                                                                pline.AddVertexAt(i, points(i), 0.0, 0.0, 0.0)
                                                            Next
                                                            pline.Closed = True
                                                            pline.Normal = normal
                                                            Dim curSpace = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                                                            curSpace.AppendEntity(pline)
                                                            tr.AddNewlyCreatedDBObject(pline, True)
                                                        End Using
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub





    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'конвертирует чертеж в метры
    <CommandMethod("ConvertDrawingMMtoM")>
    Public Sub ConvertDrawingMMtoM()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
            ed.WriteMessage(vbCrLf & "Начало преобразования мм → м...")

            ' 1. Запрос подтверждения
            Dim confirmOpts As New PromptKeywordOptions(vbCrLf & "Преобразовать весь чертеж из мм в метры? [Да/Нет] <Да>: ")
            confirmOpts.Keywords.Add("Да")
            confirmOpts.Keywords.Add("Нет")
            confirmOpts.Keywords.Default = "Да"

            Dim confirmRes As PromptResult = ed.GetKeywords(confirmOpts)

            If confirmRes.StringResult = "Нет" Then
                ed.WriteMessage(vbCrLf & "Преобразование отменено.")
                Return
            End If

            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' 2. Получаем BlockTable для работы с объектами
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                ' 3. Масштабируем все объекты в ModelSpace
                Dim scaleFactor As Double = 0.001 ' 1 мм = 0.001 м
                Dim scaleMatrix As Matrix3d = Matrix3d.Scaling(scaleFactor, Point3d.Origin)

                Dim objectCount As Integer = 0
                Dim textCount As Integer = 0
                Dim dimCount As Integer = 0

                ' 4. Перебираем все объекты
                For Each objId As ObjectId In btr
                    If Not objId.IsErased Then
                        Dim entity As Autodesk.AutoCAD.DatabaseServices.Entity = trans.GetObject(objId, OpenMode.ForWrite)

                        ' Масштабируем геометрию объекта
                        entity.TransformBy(scaleMatrix)

                        ' Обработка текста и размеров (особый случай)
                        If TypeOf entity Is DBText Then
                            Dim text As DBText = DirectCast(entity, DBText)
                            ScaleText(text, scaleFactor)
                            textCount += 1

                        ElseIf TypeOf entity Is MText Then
                            Dim mtext As MText = DirectCast(entity, MText)
                            ScaleMText(mtext, scaleFactor)
                            textCount += 1

                        ElseIf TypeOf entity Is Dimension Then
                            Dim dimObj As Dimension = DirectCast(entity, Dimension)
                            ScaleDimension(dimObj, scaleFactor)
                            dimCount += 1
                        End If

                        objectCount += 1
                    End If
                Next

                ' 5. Обновляем системные переменные
                'UpdateSystemVariables(db, scaleFactor)

                ' 6. Обновляем аннотативные масштабы
                'UpdateAnnotationScales(db, scaleFactor)

                trans.Commit()

                ed.WriteMessage(vbCrLf & "Преобразование завершено!")
                ed.WriteMessage(vbCrLf & $"Обработано объектов: {objectCount}")
                ed.WriteMessage(vbCrLf & $"Текстов: {textCount}")
                ed.WriteMessage(vbCrLf & $"Размеров: {dimCount}")
                ed.WriteMessage(vbCrLf & "Масштаб: 1 мм = 0.001 м")

            End Using

        Catch ex As Exception
            ed.WriteMessage(vbCrLf & $"Ошибка: {ex.Message}")
            Application.ShowAlertDialog($"Ошибка преобразования: {ex.Message}")
        End Try
    End Sub

    Private Sub ScaleText(text As DBText, scaleFactor As Double)
        ' Масштабируем высоту текста
        If text.Height > 0 Then
            text.Height *= scaleFactor
        End If

        ' Масштабируем ширину (если задана)
        If text.WidthFactor > 0 Then
            ' Ширина обычно не масштабируется
        End If
    End Sub

    Private Sub ScaleMText(mtext As MText, scaleFactor As Double)
        ' Масштабируем высоту MText
        If mtext.TextHeight > 0 Then
            mtext.TextHeight *= scaleFactor
        End If
    End Sub

    Private Sub ScaleDimension(dimObj As Dimension, scaleFactor As Double)
        ' Масштабируем размеры
        If dimObj.Dimscale > 0 Then
            dimObj.Dimscale *= scaleFactor
        End If

        ' Масштабируем текст размеров
        If dimObj.Dimtxt > 0 Then
            dimObj.Dimtxt *= scaleFactor
        End If

        ' Масштабируем выносные линии
        If dimObj.Dimasz > 0 Then
            dimObj.Dimasz *= scaleFactor
        End If

        If dimObj.Dimdle > 0 Then
            dimObj.Dimdle *= scaleFactor
        End If

        If dimObj.Dimexo > 0 Then
            dimObj.Dimexo *= scaleFactor
        End If

        ' Для размеров с текстом
        If TypeOf dimObj Is AlignedDimension Then
            Dim alignedDim As AlignedDimension = DirectCast(dimObj, AlignedDimension)
            ' Дополнительная обработка при необходимости
        End If
    End Sub

End Class
