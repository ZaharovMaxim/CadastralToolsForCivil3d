Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Public Class ProjectBridge
    <CommandMethod("ProjectBalka")>
    Public Sub ProjectBalka()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'исходные данные
        Dim formBearm As FormCreateBridge = New FormCreateBridge
        formBearm.ShowDialog()
        If formBearm.showDialog1 = False Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim upWidthBearm As Double = formBearm.NumericUpDown1.Value / 1000
        Dim downWidthBearm As Double = formBearm.NumericUpDown2.Value / 1000
        Dim heightBearm As Double = formBearm.NumericUpDown4.Value / 1000
        Dim lenghtPointPrStart As Double = formBearm.NumericUpDown8.Value / 1000
        Dim lenghtPointPrEnd As Double = formBearm.NumericUpDown8.Value / 1000
        Dim dLenghtEarth As Double = formBearm.NumericUpDown7.Value / 1000
        Dim offsetFirstBearm As Double = formBearm.NumericUpDown3.Value / 1000
        Dim offsetLastBearm As Double = formBearm.NumericUpDown5.Value / 1000
        Dim countRowsBearm As Double = formBearm.NumericUpDown6.Value
        Dim zazor As Double = formBearm.NumericUpDown9.Value / 1000
        Dim lenghtBearm As Double = formBearm.NumericUpDown10.Value / 1000

        Dim layerAxis As String = "Балка-Ось"
        Dim InsLayerID As ObjectId = AcadFunction.FuncAddLayer(layerAxis, 10,, 40)

        Dim layerBalkaCounter As String = "Вспомогательные построения"
        Dim InsLayerID2 As ObjectId = AcadFunction.FuncAddLayer(layerBalkaCounter)

        Dim layerBalka As String = "Балка"
        Dim InsLayerID3 As ObjectId = AcadFunction.FuncAddLayer(layerBalka)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayTablePS As String(,) = Nothing
        ReDim arrayTablePS(1, 11)
        arrayTablePS(0, 0) = "LABEL"
        arrayTablePS(1, 0) = "Название объекта"
        arrayTablePS(0, 1) = "Description"
        arrayTablePS(1, 1) = "Описание объекта"
        arrayTablePS(0, 2) = "LenghBearm"
        arrayTablePS(1, 2) = "Длина балки"
        arrayTablePS(0, 3) = "HeightBearm"
        arrayTablePS(1, 3) = "Высота балки"
        arrayTablePS(0, 4) = "WidthDownBearm"
        arrayTablePS(1, 4) = "Ширина балки по низу"
        arrayTablePS(0, 5) = "WidthUpBearm"
        arrayTablePS(1, 5) = "Ширина балки по верху"
        arrayTablePS(0, 6) = "StartLenghtPr"
        arrayTablePS(1, 6) = "Расстоние от начала балки до точки опирания"
        arrayTablePS(0, 7) = "EndLenghtPr"
        arrayTablePS(1, 7) = "Расстоние от конца балки до точки опирания"
        arrayTablePS(0, 8) = "Zazor"
        arrayTablePS(1, 8) = "Зазор между балками"
        arrayTablePS(0, 9) = "GLOBALID"
        arrayTablePS(1, 9) = "Глобальный ID"
        arrayTablePS(0, 10) = "LOCALID"
        arrayTablePS(1, 10) = "Локальный ID"
        arrayTablePS(0, 11) = "NOTE"
        arrayTablePS(1, 11) = "Примечание"

        Dim idPSTable As ObjectId = FuncPropertySet.FuncCreatedSetProperties("PROJECT_BEARM", arrayTablePS, "Балка мостовая")
        Dim rng As Random = New Random
        Dim gIG As String = rng.Next
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim Align As Alignment = Nothing 'трасса
        Dim surf As TinSurface = Nothing 'поверхность
        Dim startSection As Line = Nothing 'начальное сечение
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim mvbStyleDict As Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition(db)

            Dim alignCollectionID As ObjectIdCollection = CivilDoc.GetAlignmentIds()
            If alignCollectionID.Count = 0 Then
                MsgBox("В чертеже нет объектов типа ТРАССА. Макро прервано.")
                Exit Sub
            ElseIf alignCollectionID.Count = 1 Then
                Dim idAl As ObjectId = alignCollectionID.Item(0)
                Dim entAlign As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idAl, OpenMode.ForRead)
                Align = entAlign
            Else
                Dim acSSAlign As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Трассу: ")
                If Not (acSSAlign.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entAlign As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSAlign.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entAlign Is Alignment) Then
                    MsgBox("Выбранный примитив не является Трассой. Макро прервано.")
                    Exit Sub
                End If
                Align = entAlign
            End If

            Dim surfCollectionID As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If surfCollectionID.Count = 0 Then
                MsgBox("В чертеже нет объектов типа ПОВЕРХНОСТЬ. Макро прервано.")
                Exit Sub
            ElseIf surfCollectionID.Count = 1 Then
                Dim idSf As ObjectId = surfCollectionID.Item(0)
                Dim entSurf As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idSf, OpenMode.ForRead)
                surf = entSurf
            Else
                Dim acSurf As PromptEntityResult = ed.GetEntity(vbLf & "Выберите поверхность: ")
                If Not (acSurf.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entSurf As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSurf.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entSurf Is TinSurface) Then
                    MsgBox("Выбранный примитив не является поверхностью. Макро прервано.")
                    Exit Sub
                End If
                surf = entSurf
            End If

            Dim acStartSect As PromptEntityResult = ed.GetEntity(vbLf & "Выберите сечение: ")
            If Not (acStartSect.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim entLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acStartSect.ObjectId, OpenMode.ForRead)
            If Not (TypeOf entLine Is Line) Then
                MsgBox("Выбранный примитив не является Линией. Макро прервано.")
                Exit Sub
            End If
            startSection = entLine
            '==================================================================================================================
            'преобразовываем трассу в полилинию
            Dim alignPlineID As ObjectId = Align.GetPolyline
            Dim alignPline As Polyline = tr.GetObject(alignPlineID, OpenMode.ForRead)
            '==================================================================================================================
            'распараллеливаем трассы
            Dim arrayAxis As Double() = Nothing
            Dim countArrayAxis As Integer = 0
            ReDim Preserve arrayAxis(countArrayAxis)
            arrayAxis(countArrayAxis) = offsetFirstBearm
            countArrayAxis += 1
            ReDim Preserve arrayAxis(countArrayAxis)
            arrayAxis(countArrayAxis) = -1 * offsetFirstBearm
            countArrayAxis += 1
            'расчет длины остальных балок
            If countRowsBearm > 1 Then
                Dim deltaLenghtRow As Double = offsetLastBearm / (countRowsBearm - 1)
                For i As Integer = 1 To countRowsBearm - 1
                    Dim tempDist As Double = offsetFirstBearm + deltaLenghtRow * i
                    ReDim Preserve arrayAxis(countArrayAxis)
                    arrayAxis(countArrayAxis) = tempDist
                    countArrayAxis += 1
                    ReDim Preserve arrayAxis(countArrayAxis)
                    arrayAxis(countArrayAxis) = -1 * tempDist
                    countArrayAxis += 1
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IsNothing(alignPline) = False And IsArray(arrayAxis) = True Then
                For i As Integer = 0 To arrayAxis.Length - 1
                    Dim offsetLenghtAxis As Double = arrayAxis(i)
                    Dim dbAxisPlineCollection As DBObjectCollection = alignPline.GetOffsetCurves(offsetLenghtAxis)
                    If dbAxisPlineCollection.Count = 1 Then
                        Dim axisPline As Polyline = dbAxisPlineCollection.Item(0)
                        acBlkTblRec.AppendEntity(axisPline)
                        tr.AddNewlyCreatedDBObject(axisPline, True)
                        If IsNothing(axisPline) = False Then
                            If axisPline.Id.IsNull = False Then
                                'задаем 4 точки шкафной стенки
                                Dim startSectionPoint3d As Point3dCollection = New Point3dCollection()
                                startSectionPoint3d.Add(New Point3d(startSection.StartPoint.X, startSection.StartPoint.Y, 0))
                                startSectionPoint3d.Add(New Point3d(startSection.EndPoint.X, startSection.EndPoint.Y, 0))
                                startSectionPoint3d.Add(New Point3d(startSection.StartPoint.X, startSection.StartPoint.Y, 0))
                                startSectionPoint3d.Add(New Point3d(startSection.EndPoint.X, startSection.EndPoint.Y, 0))
                                Dim pointCollection As Point3dCollection = New Point3dCollection()
                                'находим точку пересечения трассы и сечения
                                axisPline.IntersectWith(startSection, Intersect.ExtendBoth, New Plane(), pointCollection, 0, 0)
                                If pointCollection.Count = 1 Then
                                    'начальная точка раскладки болок
                                    Dim startAlignPoint As Point3d = pointCollection.Item(0)
                                    'находим расстояние по полилинии от начала полилинии до точки пересечения
                                    Dim startAlignDist As Double = axisPline.GetDistAtPoint(startAlignPoint)
                                    For j As Integer = 0 To formBearm.DataGridView1.RowCount - 2
                                        Dim endSectionPoint3d As Point3dCollection = New Point3dCollection()
                                        Dim tempZazor As String = formBearm.DataGridView1.Rows(j).Cells(0).Value
                                        If IsNothing(tempZazor) = False Then
                                            If IsNumeric(tempZazor) = True Then
                                                zazor = Val(tempZazor) / 1000
                                            End If
                                        End If
                                        Dim tempLenghtBearm As String = formBearm.DataGridView1.Rows(j).Cells(1).Value
                                        If IsNothing(tempLenghtBearm) = False Then
                                            If IsNumeric(tempLenghtBearm) = True Then
                                                lenghtBearm = Val(tempLenghtBearm) / 1000
                                            End If
                                        End If
                                        Dim tempLenghtStartPr As String = formBearm.DataGridView1.Rows(j).Cells(2).Value
                                        If IsNothing(tempLenghtStartPr) = False Then
                                            If IsNumeric(tempLenghtStartPr) = True Then
                                                lenghtPointPrStart = Val(tempLenghtStartPr) / 1000
                                            End If
                                        End If
                                        Dim tempLenghtEndPr As String = formBearm.DataGridView1.Rows(j).Cells(3).Value
                                        If IsNothing(tempLenghtEndPr) = False Then
                                            If IsNumeric(tempLenghtEndPr) = True Then
                                                lenghtPointPrEnd = Val(tempLenghtEndPr) / 1000
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'находим отметку низа начала балки (с зазором)
                                        Dim ElevStartAlignPoint As Double = 0
                                        Try
                                            ElevStartAlignPoint = surf.FindElevationAtXY(startAlignPoint.X, startAlignPoint.Y) - heightBearm - dLenghtEarth
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            Exit For
                                        End Try
                                        startAlignPoint = New Point3d(startAlignPoint.X, startAlignPoint.Y, ElevStartAlignPoint)
                                        'находим первую (начальную точку опирания) 
                                        Dim newPoint1 As Point3d = FuncBredge.FuncFindPointBearm(axisPline, startAlignPoint, lenghtPointPrStart + zazor, surf, heightBearm, dLenghtEarth)
                                        'находим вторую точку опирания
                                        Dim distFirstPointPr As Double = axisPline.GetDistAtPoint(newPoint1)
                                        Dim newPoint2 As Point3d = FuncBredge.FuncFindPointBearm(axisPline, newPoint1, lenghtBearm - lenghtPointPrStart - lenghtPointPrEnd, surf, heightBearm, dLenghtEarth)
                                        'получаем осевую линию по низу балки за минусом рассояний до точек опирания
                                        Dim lineDownBalka As Line = New Line(newPoint1, newPoint2)
                                        acBlkTblRec.AppendEntity(lineDownBalka)
                                        tr.AddNewlyCreatedDBObject(lineDownBalka, True)
                                        lineDownBalka.LayerId = InsLayerID
                                        '======================================================================================================================
                                        'корректируем длину балки за наклон
                                        For k As Integer = 0 To 1000
                                            Dim boolBalka As Boolean = FuncBredge.FuncCorrBalka(lineDownBalka, startSectionPoint3d, heightBearm, downWidthBearm, upWidthBearm, lenghtPointPrStart, lenghtPointPrEnd, zazor, endSectionPoint3d)
                                            If boolBalka = True Then
                                                Exit For
                                            End If
                                        Next k
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'рисуем вспомогательные линии
                                        Dim line1 As Line = New Line(endSectionPoint3d.Item(0), startSectionPoint3d.Item(0))
                                        acBlkTblRec.AppendEntity(line1)
                                        tr.AddNewlyCreatedDBObject(line1, True)
                                        line1.LayerId = InsLayerID2
                                        line1.ColorIndex = 3
                                        line1.LineWeight = 40

                                        Dim line2 As Line = New Line(endSectionPoint3d.Item(1), startSectionPoint3d.Item(1))
                                        acBlkTblRec.AppendEntity(line2)
                                        tr.AddNewlyCreatedDBObject(line2, True)
                                        line2.LayerId = InsLayerID2
                                        line2.ColorIndex = 3
                                        line2.LineWeight = 40

                                        Dim line3 As Line = New Line(endSectionPoint3d.Item(2), startSectionPoint3d.Item(2))
                                        acBlkTblRec.AppendEntity(line3)
                                        tr.AddNewlyCreatedDBObject(line3, True)
                                        line3.LayerId = InsLayerID2
                                        line3.ColorIndex = 5
                                        line3.LineWeight = 40

                                        Dim line4 As Line = New Line(endSectionPoint3d.Item(3), startSectionPoint3d.Item(3))
                                        acBlkTblRec.AppendEntity(line4)
                                        tr.AddNewlyCreatedDBObject(line4, True)
                                        line4.LayerId = InsLayerID2
                                        line4.ColorIndex = 5
                                        line4.LineWeight = 40
                                        '=======================================================================================================================================
                                        'вычисляем нове начало раскладки балок
                                        Dim tempEndPr As Point3d = line1.EndPoint
                                        startAlignPoint = axisPline.GetClosestPointTo(tempEndPr, True)
                                        startAlignDist = axisPline.GetDistAtPoint(startAlignPoint)
                                        'startSectionPoint3d = endSectionPoint3d
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim arrayData As String(,) = Nothing
                                        ReDim arrayData(2, 11)
                                        arrayData(0, 0) = "PROJECT_BEARM"
                                        arrayData(1, 0) = "LABEL"
                                        arrayData(2, 0) = "Исскуственные сооружения"

                                        arrayData(0, 1) = "PROJECT_BEARM"
                                        arrayData(1, 1) = "Description"
                                        arrayData(2, 1) = "Балка мостовая"

                                        arrayData(0, 2) = "PROJECT_BEARM"
                                        arrayData(1, 2) = "LenghBearm"
                                        arrayData(2, 2) = lenghtBearm

                                        arrayData(0, 3) = "PROJECT_BEARM"
                                        arrayData(1, 3) = "HeightBearm"
                                        arrayData(2, 3) = heightBearm

                                        arrayData(0, 4) = "PROJECT_BEARM"
                                        arrayData(1, 4) = "WidthDownBearm"
                                        arrayData(2, 4) = downWidthBearm

                                        arrayData(0, 5) = "PROJECT_BEARM"
                                        arrayData(1, 5) = "WidthUpBearm"
                                        arrayData(2, 5) = upWidthBearm

                                        arrayData(0, 6) = "PROJECT_BEARM"
                                        arrayData(1, 6) = "StartLenghtPr"
                                        arrayData(2, 6) = lenghtPointPrStart

                                        arrayData(0, 7) = "PROJECT_BEARM"
                                        arrayData(1, 7) = "EndLenghtPr"
                                        arrayData(2, 7) = lenghtPointPrEnd

                                        arrayData(0, 8) = "PROJECT_BEARM"
                                        arrayData(1, 8) = "Zazor"
                                        arrayData(2, 8) = zazor

                                        arrayData(0, 9) = "PROJECT_BEARM"
                                        arrayData(1, 9) = "GLOBALID"
                                        arrayData(2, 9) = gIG

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next

                                        arrayData(0, 11) = "PROJECT_BEARM"
                                        arrayData(1, 11) = "NOTE"
                                        arrayData(2, 11) = "-"
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем многовидовой блок
                                        Try
                                            Dim nameMVBBlock As String = formBearm.DataGridView1.Rows(j).Cells(4).Value
                                            Dim idmvbstyledict As ObjectId = mvbStyleDict.GetAt(nameMVBBlock)
                                            If IsNothing(idmvbstyledict) = False Then
                                                If idmvbstyledict.IsNull = False Then
                                                    Dim mvb As Autodesk.Aec.DatabaseServices.MultiViewBlockReference = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()
                                                    mvb.BlockDefId = idmvbstyledict
                                                    mvb.Location = lineDownBalka.StartPoint
                                                    mvb.LayerId = InsLayerID3
                                                    mvb.Scale = New Scale3d(1, 1, 1)
                                                    Dim vectNormal As Vector3d = lineDownBalka.StartPoint.GetVectorTo(lineDownBalka.EndPoint)
                                                    mvb.Direction = vectNormal
                                                    mvb.SetDatabaseDefaults()
                                                    acBlkTblRec.AppendEntity(mvb)
                                                    tr.AddNewlyCreatedDBObject(mvb, True)
                                                    Dim userDict2 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                    If idPSTable.IsNull = False Then
                                                        Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(mvb, idPSTable)
                                                        Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(mvb, arrayData)
                                                    End If
                                                    'userDict2.Add("Name", "Многовидовой блок")
                                                    'userDict2.Add("GLOBALID", gIG)
                                                    'userDict2.Add("LOCALID", rng.Next)
                                                    'Dim boolWriteData1 As Boolean = FuncXRecords.FuncWriteXData(line1, "PROJECT_BEARM", userDict2)
                                                End If
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'формируем XData
                                        'Dim nameDictionary As String = "PROJECT_BEARM"
                                        'Dim userDict As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                        'userDict.Add("Name", "Axis")
                                        'userDict.Add("Zazor", zazor)
                                        'userDict.Add("LenghBearm", lenghtBearm)
                                        'userDict.Add("HeightBearm", heightBearm)
                                        'userDict.Add("WidthDownBearm", downWidthBearm)
                                        'userDict.Add("StartLenghtPr", lenghtPointPrStart)
                                        'userDict.Add("EndLenghtPr", lenghtPointPrEnd)
                                        'userDict.Add("GLOBALID", gIG)
                                        'userDict.Add("LOCALID", rng.Next)
                                        'Dim boolWriteData As Boolean = FuncXRecords.FuncWriteXData(lineDownBalka, nameDictionary, userDict)

                                        'Dim userDict1 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                        'userDict1.Add("Name", "TempAxis")
                                        'userDict1.Add("GLOBALID", gIG)
                                        'userDict1.Add("LOCALID", rng.Next)
                                        'boolWriteData = FuncXRecords.FuncWriteXData(line1, nameDictionary, userDict1)
                                        'boolWriteData = FuncXRecords.FuncWriteXData(line2, nameDictionary, userDict1)
                                        'boolWriteData = FuncXRecords.FuncWriteXData(line3, nameDictionary, userDict1)
                                        'boolWriteData = FuncXRecords.FuncWriteXData(line4, nameDictionary, userDict1)
                                        arrayData(0, 1) = "PROJECT_BEARM"
                                        arrayData(1, 1) = "Description"
                                        arrayData(2, 1) = "Ось балки"

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next
                                        If idPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(lineDownBalka, idPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(lineDownBalka, arrayData)
                                        End If

                                        arrayData(0, 1) = "PROJECT_BEARM"
                                        arrayData(1, 1) = "Description"
                                        arrayData(2, 1) = "Вспомогательная ось"

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next
                                        If idPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(line1, idPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(line1, arrayData)
                                        End If

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next
                                        If idPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(line2, idPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(line2, arrayData)
                                        End If

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next
                                        If idPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(line3, idPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(line3, arrayData)
                                        End If

                                        arrayData(0, 10) = "PROJECT_BEARM"
                                        arrayData(1, 10) = "LOCALID"
                                        arrayData(2, 10) = rng.Next
                                        If idPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(line4, idPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(line4, arrayData)
                                        End If
                                    Next j
                                End If
                            End If
                            If axisPline.IsWriteEnabled = False Then
                                axisPline.UpgradeOpen()
                                axisPline.Erase()
                            Else
                                axisPline.Erase()
                            End If
                        End If
                    End If
                Next
                If alignPline.IsWriteEnabled = False Then
                    alignPline.UpgradeOpen()
                    alignPline.Erase()
                End If
            End If
            tr.Commit()
        End Using
    End Sub
End Class
