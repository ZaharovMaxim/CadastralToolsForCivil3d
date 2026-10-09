Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Public Class AirPowerLines
    <CommandMethod("CreateAirPowerLines")>
    Public Sub AirPowerLines()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim FormAirPowerLine As AirPowerLine = New AirPowerLine
        FormAirPowerLine.ShowDialog()
        If FormAirPowerLine.boolShow = False Then
            Exit Sub
        End If
        Dim ArrayPoint As Integer() = Nothing 'массив с точками
        'ищем пути для таблиц
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim NameLayer As String = "0" 'слой начальный
        Dim indColor As Integer = 7 'цвет по блоку
        Dim MScale As Double = 1 'начальный масштаб (не правильный)
        Dim RotBlk As Double = 0 'угол поворота блока
        Dim NameBlk As String = "Условный знак ЛЭП высокого напряжения"
        Dim TableName As String = "Воздушные линии ЛЭП"
        Dim ArrayPropBlk As String(,) = Nothing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Try
                If FormAirPowerLine.RadioButton1.Checked = True Then
                    NameBlk = "Условный знак ЛЭП высокого напряжения"
                    TableName = "Воздушные линии ЛЭП"
                ElseIf FormAirPowerLine.RadioButton2.Checked = True Then
                    NameBlk = "Условный знак ЛЭП низкого напряжения"
                    TableName = "Воздушные линии ЛЭП"
                ElseIf FormAirPowerLine.RadioButton3.Checked = True Then
                    NameBlk = "Условный знак воздушной проводной линии связи"
                    TableName = "Воздушные линии связи"
                Else
                    Dim SelectBLK As PromptEntityResult = ed.GetEntity(vbLf & "Выберите образец блока: ")
                    If SelectBLK.Status = PromptStatus.Cancel Then
                        Throw New System.Exception("Прервано пользователем")
                        Exit Sub 'если нажали ESC то выходим из программы
                    End If
                    Dim SelectObj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectBLK.ObjectId, OpenMode.ForRead)
                    If TypeOf SelectObj Is BlockReference Then
                        Dim blkRef As BlockReference = SelectObj
                        NameLayer = blkRef.Layer 'слой
                        MScale = blkRef.ScaleFactors.X 'масштаб
                        indColor = blkRef.ColorIndex 'цвет
                        'ищем имя блока
                        If blkRef.IsDynamicBlock = True Then
                            Dim blockDyn As BlockTableRecord = acTrans.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead)
                            NameBlk = blockDyn.Name
                        Else
                            NameBlk = blkRef.Name
                        End If
                    Else
                        MsgBox("Выбранный объект не является блоком. Мокрос завершен")
                        Exit Sub
                    End If
                    GoTo Line2
                End If
                Dim TemplateFile As String = FormAirPowerLine.ComboBox1.Tag & FormAirPowerLine.ComboBox1.Text
                If File.Exists(TemplateFile) = True Then
                    Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(TemplateFile, NameBlk, ArrayPropBlk, "Блок", TableName)
                    If IsArray(ArrayPropBlk) = True Then
                        For i As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                            Dim NameField As String = ArrayPropBlk(0, i)
                            If NameField.Trim Like "Name" Then
                                NameBlk = ArrayPropBlk(1, i)
                                If IsNothing(NameBlk) = False Then
                                    If NameBlk.Trim.Length > 0 Then
                                        Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(TemplateFile, NameBlk)
                                        If acBlkTbl.Has(NameBlk) = False Then
                                            MsgBox("Определение блока в шаблоне не найдено. Макро прервано!!!")
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            ElseIf NameField.Trim Like "Color" Then
                                indColor = Val(ArrayPropBlk(1, i))
                            ElseIf NameField.Trim Like "Layer" Then
                                NameLayer = Val(ArrayPropBlk(1, i))
                            ElseIf NameField.Trim Like "Scale" Then
                                Dim TempScaleBlk As Double = Val(ArrayPropBlk(1, i))
                                MScale = TempScaleBlk * (FormAirPowerLine.NumericUpDown1.Value / 1000)
                            End If
                        Next i
                    Else
                        MsgBox("Характеристики блока не найдены. Макро прервано!!!")
                        Exit Sub
                    End If
                End If
                acTrans.Commit()
            Catch
                MsgBox("Ошибка. Макро прервано!!!")
            End Try
        End Using
Line2:
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '======================================================================================================================
        'выбор расстановки знаков (вручную или автоматически)
        If FormAirPowerLine.RadioButton5.Checked = True Then
            '==============================================================
            'выбираем сегмент  линии, полилинии, 2d полилинии, 3d полилинии
            Dim hgEnt As String = ""
            Do
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim options As New PromptNestedEntityOptions(vbLf & "Выберите сегмент на полилинии: ")
                    options.AllowNone = False
                    Dim result As PromptNestedEntityResult = ed.GetNestedEntity(options)
                    If result.Status = PromptStatus.Cancel Then
                        GoTo line1 'если нажали ESC то выходим из программы
                    Else
                        Dim wcsPickedPoint As Point3d = result.PickedPoint.TransformBy(ed.CurrentUserCoordinateSystem) 'координаты точки указания
                        Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(result.ObjectId, OpenMode.ForRead)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'полилиния
                        If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim StPoint As Point2d = Nothing 'начальная точка сегмента
                            Dim EndPoint As Point2d = Nothing 'конечная точка сегмента
                            Dim pline As Autodesk.AutoCAD.DatabaseServices.Polyline = DirectCast(acTrans.GetObject(result.ObjectId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                            Dim Elev As Double = pline.Elevation
                            Dim pointOnPline As Point3d = If(result.GetContainers().Length = 0, pline.GetClosestPointTo(wcsPickedPoint, False), pline.GetClosestPointTo(wcsPickedPoint.TransformBy(result.Transform.Inverse()), False))
                            'получаем нужный сегмент
                            Dim DistPline As Double = pline.GetDistAtPoint(pointOnPline)
                            For i As Integer = 0 To pline.NumberOfVertices - 2
                                Dim Dist1 As Double = pline.GetDistAtPoint(pline.GetPoint3dAt(i))
                                Dim Dist2 As Double = pline.GetDistAtPoint(pline.GetPoint3dAt(i + 1))
                                If DistPline > Dist1 And DistPline < Dist2 Then
                                    StPoint = pline.GetPoint2dAt(i)
                                    EndPoint = pline.GetPoint2dAt(i + 1)
                                    Exit For
                                End If
                            Next
                            Dim BasePoint As Point2d = New Point2d(wcsPickedPoint.X, wcsPickedPoint.Y)
                            Dim vect1 As Vector2d = StPoint.GetVectorTo(BasePoint)
                            Dim vect2 As Vector2d = EndPoint.GetVectorTo(BasePoint)
                            If vect1.Length < vect2.Length Then
                                Dim Angle As Double = StPoint.GetVectorTo(EndPoint).Angle
                                Dim insPlkPoint As Point3d = New Point3d(StPoint.X, StPoint.Y, Elev)
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(insPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            Else
                                Dim Angle As Double = EndPoint.GetVectorTo(StPoint).Angle
                                Dim insPlkPoint As Point3d = New Point3d(EndPoint.X, EndPoint.Y, Elev)
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(insPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            End If
                            hgEnt = acEnt1.Handle.ToString
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '2д полилиния
                        ElseIf TypeOf acEnt1 Is Vertex2d Then
                            Dim StPoint As Point3d = Nothing 'начальная точка сегмента
                            Dim EndPoint As Point3d = Nothing 'конечная точка сегмента
                            Dim pline As Polyline2d = DirectCast(acTrans.GetObject(result.ObjectId, OpenMode.ForRead), Polyline2d)
                            Dim pointOnPline As Point3d = If(result.GetContainers().Length = 0, pline.GetClosestPointTo(wcsPickedPoint, False), pline.GetClosestPointTo(wcsPickedPoint.TransformBy(result.Transform.Inverse()), False))
                            'получаем нужный сегмент
                            Dim DistPline As Double = pline.GetDistAtPoint(pointOnPline)
                            Dim ArrayVertex As Double(,) = Nothing
                            Dim CountArrayVertex As Integer = 0
                            For Each VertID As ObjectId In pline
                                Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForRead)
                                ReDim Preserve ArrayVertex(2, CountArrayVertex)
                                ArrayVertex(0, CountArrayVertex) = v2d.Position.X
                                ArrayVertex(1, CountArrayVertex) = v2d.Position.Y
                                ArrayVertex(2, CountArrayVertex) = v2d.Position.Z
                            Next
                            For i As Integer = 0 To ArrayVertex.GetUpperBound(1) - 1
                                Dim pt1 As Point3d = New Point3d(ArrayVertex(0, i), ArrayVertex(1, i), ArrayVertex(2, i))
                                Dim pt2 As Point3d = New Point3d(ArrayVertex(0, i + 1), ArrayVertex(1, i + 1), ArrayVertex(2, i + 1))
                                Dim Dist1 As Double = pline.GetDistAtPoint(pt1)
                                Dim Dist2 As Double = pline.GetDistAtPoint(pt2)
                                If DistPline > Dist1 And DistPline < Dist2 Then
                                    StPoint = New Point3d(ArrayVertex(0, i), ArrayVertex(1, i), ArrayVertex(2, i))
                                    EndPoint = New Point3d(ArrayVertex(0, i + 1), ArrayVertex(1, i + 1), ArrayVertex(2, i + 1))
                                    Exit For
                                End If
                            Next
                            Dim BasePoint As Point2d = New Point2d(wcsPickedPoint.X, wcsPickedPoint.Y)
                            Dim st1 As Point2d = New Point2d(StPoint.X, StPoint.Y)
                            Dim st2 As Point2d = New Point2d(EndPoint.X, EndPoint.Y)
                            Dim vect1 As Vector2d = st1.GetVectorTo(BasePoint)
                            Dim vect2 As Vector2d = st2.GetVectorTo(BasePoint)
                            If vect1.Length < vect2.Length Then
                                Dim Angle As Double = st1.GetVectorTo(st2).Angle
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(StPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            Else
                                Dim Angle As Double = st2.GetVectorTo(st1).Angle
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(EndPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            End If
                            hgEnt = acEnt1.Handle.ToString
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '3д полилиния
                        ElseIf TypeOf acEnt1 Is PolylineVertex3d Then
                            Dim StPoint As Point3d = Nothing 'начальная точка сегмента
                            Dim EndPoint As Point3d = Nothing 'конечная точка сегмента
                            Dim pline1 As PolylineVertex3d = DirectCast(acTrans.GetObject(result.ObjectId, OpenMode.ForRead), PolylineVertex3d)

                            Dim pline As Polyline3d = DirectCast(acTrans.GetObject(result.ObjectId, OpenMode.ForRead), Polyline3d)
                            Dim pointOnPline As Point3d = If(result.GetContainers().Length = 0, pline.GetClosestPointTo(wcsPickedPoint, False), pline.GetClosestPointTo(wcsPickedPoint.TransformBy(result.Transform.Inverse()), False))
                            'получаем нужный сегмент
                            Dim DistPline As Double = pline.GetDistAtPoint(pointOnPline)
                            Dim ArrayVertex As Double(,) = Nothing
                            Dim CountArrayVertex As Integer = 0
                            For Each VertID As ObjectId In pline
                                Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForRead)
                                ReDim Preserve ArrayVertex(2, CountArrayVertex)
                                ArrayVertex(0, CountArrayVertex) = v2d.Position.X
                                ArrayVertex(1, CountArrayVertex) = v2d.Position.Y
                                ArrayVertex(2, CountArrayVertex) = v2d.Position.Z
                            Next
                            For i As Integer = 0 To ArrayVertex.GetUpperBound(1) - 1
                                Dim pt1 As Point3d = New Point3d(ArrayVertex(0, i), ArrayVertex(1, i), ArrayVertex(2, i))
                                Dim pt2 As Point3d = New Point3d(ArrayVertex(0, i + 1), ArrayVertex(1, i + 1), ArrayVertex(2, i + 1))
                                Dim Dist1 As Double = pline.GetDistAtPoint(pt1)
                                Dim Dist2 As Double = pline.GetDistAtPoint(pt2)
                                If DistPline > Dist1 And DistPline < Dist2 Then
                                    StPoint = New Point3d(ArrayVertex(0, i), ArrayVertex(1, i), ArrayVertex(2, i))
                                    EndPoint = New Point3d(ArrayVertex(0, i + 1), ArrayVertex(1, i + 1), ArrayVertex(2, i + 1))
                                    Exit For
                                End If
                            Next
                            Dim BasePoint As Point2d = New Point2d(wcsPickedPoint.X, wcsPickedPoint.Y)
                            Dim st1 As Point2d = New Point2d(StPoint.X, StPoint.Y)
                            Dim st2 As Point2d = New Point2d(EndPoint.X, EndPoint.Y)
                            Dim vect1 As Vector2d = st1.GetVectorTo(BasePoint)
                            Dim vect2 As Vector2d = st2.GetVectorTo(BasePoint)
                            If vect1.Length < vect2.Length Then
                                Dim Angle As Double = st1.GetVectorTo(st2).Angle
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(StPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            Else
                                Dim Angle As Double = st2.GetVectorTo(st1).Angle
                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(EndPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                            End If
                            hgEnt = acEnt1.Handle.ToString
                        End If
                    End If
                    acTrans.Commit()
                End Using
            Loop
        Else
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'автоматически
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim PointsCogoCollection As CogoPointCollection = CivilDoc.CogoPoints
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем полилинии
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                    Dim LayerBline As String = Pline.Layer.ToString
                                    'создаем габаритный контейнер полилинии для фильтрации точек
                                    Dim MinCoord As Point3d = Pline.GeometricExtents.MinPoint
                                    Dim MaxCoord As Point3d = Pline.GeometricExtents.MaxPoint
                                    Dim Count1 As Integer = 0
                                    Dim CountPoint As Integer = PointsCogoCollection.Count
                                    If CountPoint > 0 Then
                                        For Each IDCogoPoint As ObjectId In PointsCogoCollection
                                            Dim PointCOGO1 As CogoPoint = IDCogoPoint.GetObject(OpenMode.ForRead)
                                            If PointCOGO1.Northing >= MinCoord.Y - 10 And PointCOGO1.Easting >= MinCoord.X - 10 Then
                                                If PointCOGO1.Northing <= MaxCoord.Y + 10 And PointCOGO1.Easting <= MaxCoord.X + 10 Then
                                                    ReDim Preserve ArrayPoint(Count1)
                                                    ArrayPoint(Count1) = PointCOGO1.PointNumber
                                                    Count1 += 1
                                                End If
                                            End If
                                        Next
                                    End If
                                    'начинаем перебор вершин полилинии
                                    Dim vn As Integer = Pline.NumberOfVertices
                                    For i As Integer = 0 To vn - 1
                                        Dim HPoint As Double = 0
                                        Dim PointCogoVert As CogoPoint = Nothing
                                        Dim Point As Point3d = Pline.GetPoint3dAt(i)
                                        If IsArray(ArrayPoint) = True Then
                                            PointCogoVert = FuncSurvey.FuncFindCogoPointByCoordinates(Point, ArrayPoint, 0.2)
                                            If IsNothing(PointCogoVert) = False Then
                                                If PointCogoVert.Id.IsNull = False Then
                                                    Try
                                                        Dim pt As Point2d = New Point2d(PointCogoVert.Easting, PointCogoVert.Northing)
                                                        If PointCogoVert.Easting <> Point.X Or PointCogoVert.Northing <> Point.Y Then
                                                            Pline.SetPointAt(i, pt)
                                                            Point = Pline.GetPoint3dAt(i)
                                                        End If
                                                    Catch ex As System.NullReferenceException
                                                    End Try
                                                End If
                                            End If
                                        End If
                                        If i = 0 Then
                                            Dim Point2 As Point3d = Pline.GetPoint3dAt(i + 1)
                                            Dim Angle As Double = Point.Convert2d(New Plane).GetVectorTo(Point2.Convert2d(New Plane)).Angle
                                            If FormAirPowerLine.CheckBox1.Checked = True Then
                                                Try
                                                    PointCogoVert.MarkerRotation = Angle
                                                    HPoint = PointCogoVert.Elevation
                                                Catch
                                                End Try
                                            End If
                                            Dim InsPlkPoint As Point3d = New Point3d(Point.X, Point.Y, HPoint)
                                            Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                                        ElseIf i = vn - 1 Then
                                            Dim Point0 As Point2d = Pline.GetPoint2dAt(i - 1)
                                            Dim Angle As Double = Point.Convert2d(New Plane).GetVectorTo(Point0).Angle
                                            If FormAirPowerLine.CheckBox1.Checked = True Then
                                                Try
                                                    PointCogoVert.MarkerRotation = Angle
                                                    HPoint = PointCogoVert.Elevation
                                                Catch
                                                End Try
                                            End If
                                            Dim InsPlkPoint As Point3d = New Point3d(Point.X, Point.Y, HPoint)
                                            Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle)
                                        Else
                                            Dim Point0 As Point2d = Pline.GetPoint2dAt(i - 1)
                                            Dim Point2 As Point2d = Pline.GetPoint2dAt(i + 1)
                                            Dim Angle1 As Double = Point.Convert2d(New Plane).GetVectorTo(Point0).Angle
                                            Dim Angle2 As Double = Point.Convert2d(New Plane).GetVectorTo(Point2).Angle
                                            Dim Angle As Double = (Angle1 + Angle2) / 2
                                            If FormAirPowerLine.CheckBox1.Checked = True Then
                                                Try
                                                    PointCogoVert.MarkerRotation = Angle
                                                    HPoint = PointCogoVert.Elevation
                                                Catch
                                                End Try
                                            End If
                                            Dim InsPlkPoint As Point3d = New Point3d(Point.X, Point.Y, HPoint)
                                            Dim InsBlk1 As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle1)
                                            Dim InsBlk2 As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPlkPoint, NameBlk, NameLayer, indColor, MScale, 20, Angle2)
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    Next
                End If
                acTrans.Commit()
            End Using
        End If
line1:
    End Sub
End Class
