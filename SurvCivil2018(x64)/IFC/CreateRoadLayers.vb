Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.EditorInput
Imports System.Runtime.InteropServices
Imports Autodesk.Aec.Modeler
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.BoundaryRepresentation

Public Class CreateRoadLayers

    Public Shared Function GetSubentity(ByVal ed As Editor, ByVal tr As Transaction, ByVal ss As SelectionSet, ByVal type As SubentityType, <Out> ByRef solId As ObjectId) As SubentityId
        Dim retId As SubentityId = New SubentityId()
        Dim so As SelectedObject = ss(0)
        solId = so.ObjectId
        Dim solid3d As Solid3d = TryCast(tr.GetObject(so.ObjectId, OpenMode.ForRead), Solid3d)
        Dim sso As SelectedSubObject() = so.GetSubentities()
        Dim subentityType As SubentityType = sso(0).FullSubentityPath.SubentId.Type
        If subentityType = type Then
            retId = sso(0).FullSubentityPath.SubentId
        Else
            ed.WriteMessage(vbLf & "Invalid Subentity Type: {0}, " & "Expecting a {1} to be selected." & vbLf, subentityType, type.ToString())
        End If
        Return retId
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==============================================================================================================================
    'создание слоев дорожной одежды (нижний слой с выступом)
    <CommandMethod("CreateRoadLayersBySurface")>
    Public Sub CreateRoadLayersBySurface()
        '' Get the current document and database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        'запускаем форму
        Dim formSurface As FormCreateRoadLayers = New FormCreateRoadLayers()
        formSurface.ShowDialog()
        If formSurface.boolOk = False Then
            Exit Sub
        End If
        '=====================================================================================================================
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
            Dim nameLayerSolid3d As String = "0"
            'доступ к поверхности
            Dim tinS As TinSurface = Nothing
            Dim nameSurf As String = formSurface.ComboBox1.Text
            If IsNothing(nameSurf) = False Then
                If nameSurf.Trim.Length > 0 Then
                    Dim idSurfCollection As ObjectIdCollection = CivilDoc.GetSurfaceIds()
                    If idSurfCollection.Count > 0 Then
                        For Each id As ObjectId In idSurfCollection
                            Dim surf As TinSurface = acTrans.GetObject(id, OpenMode.ForRead)
                            If surf.Name Like nameSurf Then
                                tinS = surf
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
            If IsNothing(tinS) = True Then
                MsgBox("Поверхность с выбранным именем не найденв. Макро прервано!!!")
                Exit Sub
            End If
            If tinS.Id.IsNull = True Then
                MsgBox("Поверхность с выбранным именем не найденв. Макро прервано!!!")
                Exit Sub
            End If
            '====================================================================================================================
            'извлекаем границу поверхности в виде 3d полилинии
            Dim pLIne3dTop As Polyline3d = Nothing
            Dim boundaresCollection As ObjectIdCollection = tinS.ExtractBorder(SurfaceExtractionSettingsType.Plan)
            If boundaresCollection.Count > 0 Then
                Dim MaxLenghtPline As Double = 0
                For i As Integer = 0 To boundaresCollection.Count - 1
                    Dim idEnt As ObjectId = boundaresCollection.Item(i)
                    If idEnt.IsNull = False Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(idEnt, OpenMode.ForRead)
                        If TypeOf acEnt Is Polyline3d Then
                            Dim temp3dPoly As Polyline3d = acEnt
                            If temp3dPoly.Length > MaxLenghtPline Then
                                pLIne3dTop = acEnt
                                MaxLenghtPline = temp3dPoly.Length
                            End If
                        End If
                    End If
                Next
            End If
            If IsNothing(pLIne3dTop) = True Then
                MsgBox("Не удалось извлеч границу поверхности. Макро прервано!!!")
                Exit Sub
            End If
            If pLIne3dTop.Id.IsNull = True Then
                MsgBox("Не удалось извлеч границу поверхности. Макро прервано!!!")
                Exit Sub
            End If
            Dim TempLayer As String = "Temp_Construction"
            Dim idLayerTemp As ObjectId = AcadFunction.FuncAddLayer(TempLayer, 0,, 20)
            If IsNothing(idLayerTemp) = True Then
                TempLayer = "0"
            End If
            If idLayerTemp.IsNull = True Then
                TempLayer = "0"
            End If
            '==========================================================================================================
            'вычленяем строки из таблицы
            Dim ArrayTableRezult As String(,) = Nothing
            Dim countArrayTableRezult As Integer = 0
            Dim indexColor As Integer = 0
            Dim offsetTop As Double = 0
            Dim offsetDown As Double = 0
            Dim rng As Random = New Random()
            For i As Integer = 0 To formSurface.DataGridView1.RowCount - 1
                'толщина слоя
                Dim strLayerThickness As String = formSurface.DataGridView1.Rows(i).Cells(0).Value
                If IsNothing(strLayerThickness) Then Continue For
                If strLayerThickness.Trim.Length = 0 Then Continue For
                Dim layerThickness As Double = Val(strLayerThickness.Trim)
                layerThickness = Math.Abs(layerThickness / 100)
                If layerThickness <= 0 Then Continue For
                'имя слоя
                Dim cmbbox As DataGridViewComboBoxCell = formSurface.DataGridView1.Rows(i).Cells(1)
                Dim LayerVal As String = cmbbox.Value
                If IsNothing(LayerVal) Then LayerVal = nameLayerSolid3d
                If LayerVal.Trim.Length = 0 Then LayerVal = nameLayerSolid3d
                'цвет
                Dim strColor As String = formSurface.DataGridView1.Rows(i).Cells(2).Value
                If IsNothing(strColor) Then indexColor = rng.Next(1, 10)
                If strColor.Trim.Length = 0 Then indexColor = rng.Next(1, 10)
                indexColor = Val(strColor)
                If indexColor < 0 Then indexColor = 0
                If indexColor > 255 Then indexColor = 256
                'длина выступа
                Dim strOffsetTop As String = formSurface.DataGridView1.Rows(i).Cells(3).Value
                If IsNothing(strOffsetTop) Then offsetTop = 0
                If strOffsetTop.Trim.Length = 0 Then offsetTop = 0
                offsetTop = Val(strOffsetTop)
                offsetTop = Math.Abs(offsetTop / 100)
                'длина откоса
                Dim strOffsetDown As String = formSurface.DataGridView1.Rows(i).Cells(4).Value
                If IsNothing(strOffsetDown) Then offsetDown = 0
                If strOffsetDown.Trim.Length = 0 Then offsetDown = 0
                offsetDown = Val(strOffsetDown)
                offsetDown = Math.Abs(offsetDown / 100)
                'PS
                Dim cmbboxPS As DataGridViewComboBoxCell = formSurface.DataGridView1.Rows(i).Cells(5)
                Dim PSVal As String = cmbboxPS.Value
                If IsNothing(PSVal) Then PSVal = ""

                'заполняем массив со значениями
                ReDim Preserve ArrayTableRezult(5, countArrayTableRezult)
                ArrayTableRezult(0, countArrayTableRezult) = layerThickness
                ArrayTableRezult(1, countArrayTableRezult) = LayerVal
                ArrayTableRezult(2, countArrayTableRezult) = indexColor
                ArrayTableRezult(3, countArrayTableRezult) = offsetTop
                ArrayTableRezult(4, countArrayTableRezult) = offsetDown
                ArrayTableRezult(5, countArrayTableRezult) = PSVal
                countArrayTableRezult += 1
            Next i
            If IsArray(ArrayTableRezult) = False Then
                MsgBox("Данные введены не корректно!!! Макро прервано.")
                Exit Sub
            End If
            'давим поверхности
            Dim levelValue As Double = formSurface.NumericUpDown1.Value
            Dim valueDept As Double = 0
            Dim idPsTable As ObjectId = Nothing
            For i As Integer = 0 To ArrayTableRezult.GetUpperBound(1)
                Try
                    'берем толщину слоя
                    valueDept = Val(ArrayTableRezult(0, i))
                    valueDept = Math.Abs(valueDept)
                    If valueDept <= 0 Then Continue For
                    'читаем слой и создаем его если требуется
                    Dim nameLayer As String = ArrayTableRezult(1, i)
                    Dim idLayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer, 0,, 20)
                    If IsNothing(idLayer) = True Then
                        nameLayer = "0"
                    End If
                    If idLayer.IsNull = True Then
                        nameLayer = "0"
                    End If
                    'цвет солида
                    Dim indColor As Integer = Val(ArrayTableRezult(2, i))
                    'считывем полку
                    offsetTop = Val(ArrayTableRezult(3, i))
                    'считываем откос
                    offsetDown = Val(ArrayTableRezult(4, i))
                    'таблица PS
                    Dim strPs As String = ArrayTableRezult(5, i)
                    If strPs.Trim.Length > 0 Then
                        idPsTable = FuncPropertySet.FuncReturnIDDictSetProperties(strPs)
                    End If
                    '===============================================================================================
                    'строим солиды
                    Dim SolidSurf As Solid3d = Nothing
                    'надо разблокировать строку
                    Dim idSolidCollection As ObjectIdCollection '= tinS.CreateSolidsAtDepth(valueDept, nameLayer, indColor)
                    If idSolidCollection.Count > 0 Then
                        For Each idTempSurf As ObjectId In idSolidCollection
                            If idTempSurf.IsNull = False Then
                                SolidSurf = idTempSurf.GetObject(OpenMode.ForWrite)
                                Dim DistMoveSolid As Double = -1 * (valueDept - levelValue)
                                Dim boolMoveSurf3d As Boolean = MathFunction.FuncMoveZSolid3d(SolidSurf, DistMoveSolid)
                                '=================================================================================================
                                'коллекция для верхней поверхности
                                Dim idcoll3dPolineTop As ObjectIdCollection = New ObjectIdCollection
                                'коллекция для нижней поверхности
                                Dim idcoll3dPlineDown As ObjectIdCollection = New ObjectIdCollection
                                'коллекция для внешней границы поверхности
                                Dim idcoll3dPlineOutBoundary As ObjectIdCollection = New ObjectIdCollection
                                'коллекция для внутренней границы поверхности
                                Dim idcoll3dPlineHideBoundary As ObjectIdCollection = New ObjectIdCollection
                                'делаем внутренний островок
                                If offsetTop > 0 OrElse offsetDown > 0 Then
                                    '==================================================================================================
                                    'забираем в коллекция верх 3д полилинии предварительно пересчитав отметки вершин
                                    Dim pLIne3doffsetTop As Polyline3d = pLIne3dTop.Clone()
                                    acBlkTblRec.AppendEntity(pLIne3doffsetTop)
                                    acTrans.AddNewlyCreatedDBObject(pLIne3doffsetTop, True)
                                    'меняем отметки
                                    For Each vertID As ObjectId In pLIne3doffsetTop
                                        Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                        vert.Position = New Point3d(vert.Position.X, vert.Position.Y, vert.Position.Z - levelValue)
                                    Next
                                    pLIne3doffsetTop.Closed = True
                                    pLIne3doffsetTop.Layer = TempLayer
                                    pLIne3doffsetTop.ColorIndex = 3
                                    idcoll3dPolineTop.Add(pLIne3doffsetTop.Id) 'добавляем в коллекцию для верхней поверхности
                                    '====================================================================================================
                                    'низ последней поверхности без полки
                                    Dim pLIne3doffseDown As Polyline3d = pLIne3dTop.Clone()
                                    acBlkTblRec.AppendEntity(pLIne3doffseDown)
                                    acTrans.AddNewlyCreatedDBObject(pLIne3doffseDown, True)
                                    'меняем отметки
                                    For Each vertID As ObjectId In pLIne3doffseDown
                                        Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                        vert.Position = New Point3d(vert.Position.X, vert.Position.Y, vert.Position.Z - levelValue - valueDept)
                                    Next
                                    pLIne3doffseDown.Closed = True
                                    pLIne3doffseDown.Layer = TempLayer
                                    pLIne3doffseDown.ColorIndex = 5
                                    idcoll3dPlineDown.Add(pLIne3doffseDown.Id) 'добавляем в коллекцию для верхней поверхности
                                End If
                                '========================================================================================================
                                'полка
                                If offsetTop > 0 Then
                                    'создаем новую полилинию для функции offset
                                    Dim plineTemp As Polyline = New Polyline()
                                    Dim plineOffsetTemp As Polyline = New Polyline()
                                    Dim countVert As Integer = 0
                                    Dim ind As Integer = 1
                                    For Each vertID As ObjectId In pLIne3dTop
                                        Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                        plineTemp.AddVertexAt(countVert, vert.Position.Convert2d(New Plane), 0, 0, 0)
                                        countVert += 1
                                    Next
                                    acBlkTblRec.AppendEntity(plineTemp)
                                    acTrans.AddNewlyCreatedDBObject(plineTemp, True)
                                    idcoll3dPlineHideBoundary.Add(plineTemp.Id) 'добавляем в коллекцию с островком
                                    'смотрим ее направление
                                    If plineTemp.NumberOfVertices > 2 Then
                                        Dim area As Double = MathFunction.GetArea(plineTemp)
                                        If area > 0 Then
                                            ind = 1
                                        Else
                                            ind = -1
                                        End If
                                    End If
Line1:
                                    offsetTop = offsetTop * ind
                                    'делаем ofset
                                    Dim dbPlinecollection As DBObjectCollection = plineTemp.GetOffsetCurves(offsetTop)
                                    Dim maxPlineLenght As Double = 0
                                    If dbPlinecollection.Count > 0 Then
                                        For k As Integer = 0 To dbPlinecollection.Count - 1
                                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = dbPlinecollection.Item(k)
                                            If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                Dim TempPline As Polyline = ent
                                                If TempPline.Length > maxPlineLenght Then
                                                    plineOffsetTemp = ent
                                                End If
                                            End If
                                        Next k
                                    End If
                                    If plineOffsetTemp.Area < plineTemp.Area Then
                                        ind = ind * -1
                                        GoTo Line1
                                    End If
                                    '============================================================================================
                                    'назначаем X и Y координаты
                                    If IsNothing(plineOffsetTemp) = False Then
                                        acBlkTblRec.AppendEntity(plineOffsetTemp)
                                        acTrans.AddNewlyCreatedDBObject(plineOffsetTemp, True)
                                        Dim point1Collection As New Point3dCollection()
                                        Dim point2Collection As New Point3dCollection()
                                        For k1 As Integer = 0 To plineOffsetTemp.NumberOfVertices - 1
                                            Dim pt3dtemp As Point3d = plineOffsetTemp.GetPoint3dAt(k1)
                                            Dim pt3dtemp1 As Point3d = Nothing
                                            Dim pt3dtemp2 As Point3d = Nothing
                                            Dim MinLenDist As Double = 999999999
                                            For Each vertID As ObjectId In pLIne3dTop
                                                Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                                Dim lenVerttemp As Double = pt3dtemp1.Convert2d(New Plane).GetVectorTo(vert.Position.Convert2d(New Plane)).Length
                                                If lenVerttemp < MinLenDist Then
                                                    MinLenDist = lenVerttemp
                                                    pt3dtemp1 = New Point3d(pt3dtemp.X, pt3dtemp.Y, vert.Position.Z - levelValue)
                                                    pt3dtemp2 = New Point3d(pt3dtemp.X, pt3dtemp.Y, vert.Position.Z - levelValue - valueDept)
                                                End If
                                            Next
                                            point1Collection.Add(pt3dtemp1)
                                            point2Collection.Add(pt3dtemp2)
                                        Next k1
                                        If point1Collection.Count > 1 Then
                                            Dim pLine3dTopOffset2 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, point1Collection, True)
                                            pLine3dTopOffset2.Layer = TempLayer
                                            pLine3dTopOffset2.ColorIndex = 3
                                            acBlkTblRec.AppendEntity(pLine3dTopOffset2)
                                            acTrans.AddNewlyCreatedDBObject(pLine3dTopOffset2, True)
                                            idcoll3dPolineTop.Add(pLine3dTopOffset2.Id)
                                            'если откоса нет заносим эту линию как внешняя граница поверхности
                                            If offsetDown = 0 Then
                                                idcoll3dPlineOutBoundary.Add(plineOffsetTemp.Id)
                                            Else
                                                If plineOffsetTemp.IsWriteEnabled = False Then
                                                    plineOffsetTemp.UpgradeOpen()
                                                End If
                                                plineOffsetTemp.Erase()
                                            End If
                                            Dim pLine3dDownOffset2 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, point2Collection, True)
                                            pLine3dDownOffset2.Layer = TempLayer
                                            pLine3dDownOffset2.ColorIndex = 5
                                            acBlkTblRec.AppendEntity(pLine3dDownOffset2)
                                            acTrans.AddNewlyCreatedDBObject(pLine3dDownOffset2, True)
                                            idcoll3dPlineDown.Add(pLine3dDownOffset2.Id)
                                        End If
                                    End If
                                End If
                                '===================================================================================================================
                                'рисуем низ
                                If Math.Abs(offsetDown) > 0 Then
                                    offsetDown = Math.Abs(offsetTop) + Math.Abs(offsetDown)
                                    If offsetDown > offsetTop Then
                                        'создаем новую полилинию для функции offset
                                        Dim plineTemp2 As Polyline = New Polyline()
                                        Dim plineOffsetTemp2 As Polyline = New Polyline()
                                        Dim countVert As Integer = 0
                                        Dim ind As Integer = 1
                                        For Each vertID As ObjectId In pLIne3dTop
                                            Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                            plineTemp2.AddVertexAt(countVert, vert.Position.Convert2d(New Plane), 0, 0, 0)
                                            countVert += 1
                                        Next
                                        acBlkTblRec.AppendEntity(plineTemp2)
                                        acTrans.AddNewlyCreatedDBObject(plineTemp2, True)
                                        'смотрим ее направление
                                        If plineTemp2.NumberOfVertices > 2 Then
                                            Dim area As Double = MathFunction.GetArea(plineTemp2)
                                            If area > 0 Then
                                                ind = 1
                                            Else
                                                ind = -1
                                            End If
                                        End If
Line2:
                                        offsetDown = offsetDown * ind
                                        'делаем ofset
                                        Dim dbPlinecollection As DBObjectCollection = plineTemp2.GetOffsetCurves(offsetDown)
                                        Dim maxPlineLenght As Double = 0
                                        If dbPlinecollection.Count > 0 Then
                                            For k As Integer = 0 To dbPlinecollection.Count - 1
                                                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = dbPlinecollection.Item(k)
                                                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                                    Dim TempPline As Polyline = ent
                                                    If TempPline.Length > maxPlineLenght Then
                                                        plineOffsetTemp2 = ent
                                                    End If
                                                End If
                                            Next k
                                        End If
                                        If plineOffsetTemp2.Area < plineTemp2.Area Then
                                            ind = ind * -1
                                            GoTo Line2
                                        End If
                                        '============================================================================================
                                        'назначаем X и Y координаты
                                        If IsNothing(plineOffsetTemp2) = False Then
                                            acBlkTblRec.AppendEntity(plineOffsetTemp2)
                                            acTrans.AddNewlyCreatedDBObject(plineOffsetTemp2, True)
                                            Dim point2Collection As New Point3dCollection()
                                            For k1 As Integer = 0 To plineOffsetTemp2.NumberOfVertices - 1
                                                Dim pt3dtemp As Point3d = plineOffsetTemp2.GetPoint3dAt(k1)
                                                Dim pt3dtemp1 As Point3d = Nothing
                                                Dim pt3dtemp2 As Point3d = Nothing
                                                Dim MinLenDist As Double = 999999999
                                                For Each vertID As ObjectId In pLIne3dTop
                                                    Dim vert As PolylineVertex3d = acTrans.GetObject(vertID, OpenMode.ForWrite)
                                                    Dim lenVerttemp As Double = pt3dtemp1.Convert2d(New Plane).GetVectorTo(vert.Position.Convert2d(New Plane)).Length
                                                    If lenVerttemp < MinLenDist Then
                                                        MinLenDist = lenVerttemp
                                                        pt3dtemp2 = New Point3d(pt3dtemp.X, pt3dtemp.Y, vert.Position.Z - levelValue - valueDept)
                                                    End If
                                                Next
                                                point2Collection.Add(pt3dtemp2)
                                            Next k1
                                            If point2Collection.Count > 1 Then
                                                Dim pLine3dDownOffset2 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, point2Collection, True)
                                                pLine3dDownOffset2.Layer = TempLayer
                                                pLine3dDownOffset2.ColorIndex = 3
                                                acBlkTblRec.AppendEntity(pLine3dDownOffset2)
                                                acTrans.AddNewlyCreatedDBObject(pLine3dDownOffset2, True)
                                                idcoll3dPolineTop.Add(pLine3dDownOffset2.Id)
                                                idcoll3dPlineDown.Add(pLine3dDownOffset2.Id)
                                                idcoll3dPlineOutBoundary.Add(plineOffsetTemp2.Id)
                                            End If
                                        End If
                                        If plineTemp2.IsWriteEnabled = False Then
                                            plineTemp2.UpgradeOpen()
                                        End If
                                        plineTemp2.Erase()
                                    End If
                                End If
                                '===========================================================================================
                                'создаем новые поверхности для вычитания
                                If idcoll3dPolineTop.Count > 0 And idcoll3dPlineDown.Count > 0 Then
                                    'поверхность для верха
                                    Dim bound As SurfaceBoundaryType = SurfaceBoundaryType.Outer
                                    Dim bound1 As SurfaceBoundaryType = SurfaceBoundaryType.Hide
                                    Dim surfaceStyleId As ObjectId = CivilDoc.Styles.SurfaceStyles(0)
                                    Dim nameTempSurf As String = "tinV" & rng.Next
                                    Dim surfaceId As ObjectId = TinSurface.Create(nameTempSurf, surfaceStyleId)
                                    Dim tin1 As TinSurface = acTrans.GetObject(surfaceId, OpenMode.ForWrite)
                                    tin1.BreaklinesDefinition.AddStandardBreaklines(idcoll3dPolineTop, 10, 5, 5, 0)
                                    tin1.BoundariesDefinition.AddBoundaries(idcoll3dPlineOutBoundary, 1, bound, True)
                                    tin1.BoundariesDefinition.AddBoundaries(idcoll3dPlineHideBoundary, 1, bound1, True)
                                    '============================================================================================
                                    'нижняя поверхность
                                    Dim surfaceStyleId2 As ObjectId = CivilDoc.Styles.SurfaceStyles(0)
                                    Dim nameTempSurf2 As String = "tinN" & rng.Next
                                    Dim surfaceId2 As ObjectId = TinSurface.Create(nameTempSurf2, surfaceStyleId2)
                                    Dim tin2 As TinSurface = acTrans.GetObject(surfaceId2, OpenMode.ForWrite)
                                    tin2.BreaklinesDefinition.AddStandardBreaklines(idcoll3dPlineDown, 10, 5, 5, 0)
                                    tin2.BoundariesDefinition.AddBoundaries(idcoll3dPlineOutBoundary, 1, bound, True)
                                    'новая коллекция для результирующего слоя
                                    Dim idNewSurfColl As ObjectIdCollection = New ObjectIdCollection()
                                    'перед использованием надо разблокировать нижнюю строку
                                    'idNewSurfColl = tin1.CreateSolidsAtSurface(surfaceId2, nameLayer, indColor)
                                    Dim idNewSolidTemp As ObjectId = Nothing
                                    If idNewSurfColl.Count > 0 Then
                                        idNewSolidTemp = idNewSurfColl.Item(0)
                                    End If
                                    Dim newSurfaceSurf As Solid3d = acTrans.GetObject(idNewSolidTemp, OpenMode.ForWrite)
                                    SolidSurf.BooleanOperation(BooleanOperationType.BoolUnite, newSurfaceSurf)
                                    '==================================================================================
                                    'удаляем поверхности
                                    tin1.Erase()
                                    tin2.Erase()
                                    If formSurface.CheckBox1.Checked = False Then
                                        If idcoll3dPolineTop.Count > 0 Then
                                            For Each id As ObjectId In idcoll3dPolineTop
                                                Dim plinet As Polyline3d = id.GetObject(OpenMode.ForWrite)
                                                plinet.Erase()
                                            Next
                                        End If
                                        If idcoll3dPlineDown.Count > 0 Then
                                            For Each id As ObjectId In idcoll3dPlineDown
                                                Dim plinet As Polyline3d = id.GetObject(OpenMode.ForWrite)
                                                plinet.Erase()
                                            Next
                                        End If
                                    End If
                                End If
                                If idcoll3dPlineHideBoundary.Count > 0 Then
                                    For Each id As ObjectId In idcoll3dPlineHideBoundary
                                        Dim plinet As Polyline = id.GetObject(OpenMode.ForWrite)
                                        plinet.Erase()
                                    Next
                                End If
                                If idcoll3dPlineOutBoundary.Count > 0 Then
                                    For Each id As ObjectId In idcoll3dPlineOutBoundary
                                        Dim plinet As Polyline = id.GetObject(OpenMode.ForWrite)
                                        plinet.Erase()
                                    Next
                                End If
                            End If
                        Next
                    End If
                    If IsNothing(idPsTable) = False Then
                        If idPsTable.IsNull = False Then
                            Dim boolInsPs As Boolean = FuncPropertySet.FuncAddDataSetProperties(SolidSurf, idPsTable)
                        End If
                    End If
                    If pLIne3dTop.IsWriteEnabled = False Then
                        pLIne3dTop.UpgradeOpen()
                        pLIne3dTop.Erase()
                    End If
                Catch ex As System.Exception
                End Try
                levelValue = levelValue + valueDept
            Next i
            acTrans.Commit()
        End Using
    End Sub


    <CommandMethod("EC")>
    Public Sub EdgeChamfer()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim pso As PromptSelectionOptions = New PromptSelectionOptions()
        pso.MessageForAdding = vbLf & "Выберите ребро для снятия фаски"
        pso.SingleOnly = True
        pso.ForceSubSelections = True
        Dim psrEdge As PromptSelectionResult = ed.GetSelection(pso)
        If psrEdge.Status <> PromptStatus.OK Then Return
        pso.MessageForAdding = vbLf & "Выберите грань, которой принадлежит ребро"
        pso.SinglePickInSpace = True
        Dim psrFace As PromptSelectionResult = ed.GetSelection(pso)
        If psrFace.Status <> PromptStatus.OK Then Return
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim edgeSolId, faceSolId As ObjectId
                Dim edgeId As SubentityId = GetSubentity(ed, tr, psrEdge.Value, SubentityType.Edge, edgeSolId)
                Dim faceId As SubentityId = GetSubentity(ed, tr, psrFace.Value, SubentityType.Face, faceSolId)
                If edgeSolId = faceSolId AndAlso edgeSolId <> ObjectId.Null Then
                    Dim solid As Solid3d = TryCast(tr.GetObject(edgeSolId, OpenMode.ForWrite), Solid3d)
                    If solid IsNot Nothing Then
                        Dim edgeIds As SubentityId() = New SubentityId() {edgeId}
                        solid.ChamferEdges(edgeIds, faceId, 0.5, 1.0) 'фаска
                        solid.ExtrudeFaces(edgeIds, 0.5, 0) 'Удлинить
                    End If

                    Try
                        Using brp As New Brep(solid)
                            Dim length As Integer = 0
                            For Each edg As Autodesk.AutoCAD.BoundaryRepresentation.Edge In brp.Edges
                                Dim crv As ExternalCurve3d = edg.Curve
                                If crv.IsCircularArc Then
                                    Dim line As CircularArc3d = crv.NativeCurve
                                    length = length + line.GetLength(line.StartAngle, line.EndAngle, 0.01)
                                ElseIf crv.IsEllipticalArc Then
                                    Dim line As EllipticalArc3d = crv.NativeCurve
                                    length = length + line.GetLength(line.StartAngle, line.EndAngle, 0.01)
                                ElseIf crv.IsLineSegment Then
                                    Dim line As LineSegment3d = crv.NativeCurve
                                    length = length + line.Length()
                                End If
                                crv.Dispose()
                            Next
                            MsgBox("Total: " & (length / 2))
                        End Using
                    Catch ex As Autodesk.AutoCAD.BoundaryRepresentation.Exception
                        MsgBox(ex.Message & vbLf & ex.ErrorStatus)
                    End Try


                    tr.Commit()
                End If
            End Using
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            ed.WriteMessage(ex.Message)
        End Try
    End Sub




End Class
