Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

Public Class LandCadastralReport
    '=========================================================================================================
    'таблица с координатами участков постоянного и временного отводов в MS Excel, Word, Autocad
    <CommandMethod("ListCoordinatesCadastralObject")>
    Public Sub ListCoordinatesCadastralObject()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim SelectParcel As PromptSelectionResult = ed.GetSelection()
                If SelectParcel.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                    Dim acSSetCadastralObject As SelectionSet = SelectParcel.Value
                    'фильтуем объекты
                    Dim dictCadastralAllObject As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Geometry)) = SelectCadastralObject.getCadastralObjectBySelectionSet(db, acSSetCadastralObject)
                    If dictCadastralAllObject.Count = 0 Then
                        tr.Abort()
                        locker.Dispose()
                        MsgBox("Не удалось найти кадастровые объекты в выбранных примитивах.")
                        Exit Sub
                    End If
                    Dim listCategoryObject As List(Of String) = dictCadastralAllObject.Keys.ToList
                    'вызываем диалоговое окно 
                    Dim formCatalogCoordinates As CadastralTools.FormReportCatalogCoordinatesCadastralObject = New CadastralTools.FormReportCatalogCoordinatesCadastralObject(listCategoryObject)
                    formCatalogCoordinates.ShowDialog()
                    If formCatalogCoordinates.boolShow = False Then
                        tr.Abort()
                        locker.Dispose()
                        Exit Sub
                    End If
                    Dim deskObject As String = formCatalogCoordinates.deskObject
                    If dictCadastralAllObject.ContainsKey(deskObject) = True Then
                        'получаем нужный тип объектов (если это полигоны, то еще с учетом дырок, если они там есть
                        Dim dictCadastralObject As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Geometry)) = CadastralTools.SelectCadastralObject.getCadastralObjectByDesk(dictCadastralAllObject, deskObject)
                        '====================================================================================================
                        'фильтуем точки чертежа в пределах видимости
                        Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(db)
                        '====================================================================================================
                        'запускаем программу по выводу координат
                        Dim classCatalogCoordinates As CadastralTools.CatalogCoordinates = New CadastralTools.CatalogCoordinates
                        classCatalogCoordinates.CatalogExportCoordinates(dictCadastralObject, listBlockReference, formCatalogCoordinates, AddressOf DrawCatalogTableCAD, AddressOf ExportParcelDwgCAD)
                    End If
                End If
                tr.Commit()
            End Using
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод каталога координат в виде таблицы AutoCAD (делегат для CadastralTools.CatalogCoordinates)
    'таблицы размещаются последовательно с шагом 100 мм по оси X (аналог PointInsTbl1.X + 100 эталона)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private pointInsertTableCAD As Autodesk.AutoCAD.Geometry.Point3d = Nothing
    Private boolAskPointInsert As Boolean = False
    'папка для сохранения DWG и флаг её выбора (ветка "Участок в отдельный файл")
    Private NameDirExportDwg As String = ""
    Private boolAskFolderDwg As Boolean = False
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'экспорт примитивов объекта в отдельный DWG через шаблон (делегат для CadastralTools.CatalogCoordinates)
    'перенесено максимально похоже на ветку "Участок в отдельный файл" эталона
    '(Territory_planning_project\CatalogExportCoordinates.vb, строки 939-1023)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub ExportParcelDwgCAD(ByVal objCatalog As CadastralTools.CatalogCoordinatesObject)
        If IsNothing(objCatalog) = True Then Return
        If IsNothing(objCatalog.ListHandles) = True Then Return
        If objCatalog.ListHandles.Count = 0 Then Return
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        'указываем папку для сохранения (один раз на команду, аналог j = 0 эталона)
        If boolAskFolderDwg = False Then
            NameDirExportDwg = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите директорию", 1)
            If IO.Directory.Exists(NameDirExportDwg) = False Then
                Return
            End If
            boolAskFolderDwg = True
        End If
        Dim NameDoc As String = objCatalog.ConditionNumber


        NameDoc = NameDoc.Replace("_pxqc;", "")
        NameDoc = NameDoc.Replace("\", "_")
        NameDoc = NameDoc.Replace("/", "_")
        NameDoc = NameDoc.Replace(":", "_")
        NameDoc = NameDoc.Replace("*", "_")
        NameDoc = NameDoc.Replace(".", "_")
        NameDoc = NameDoc.Replace(",", "_")
        Dim strSave As String = NameDirExportDwg & "\" & NameDoc & ".dwg"
        'защита от перезаписи существующего файла (аналог эталона)
        If IO.File.Exists(strSave) = True Then
            Dim rng As Random = New Random
            NameDoc = NameDoc & rng.Next
            strSave = NameDirExportDwg & "\" & NameDoc & ".dwg"
        End If
        'заносим объекты в коллекцию по handle'ам контура
        Dim idColl As ObjectIdCollection = New ObjectIdCollection
        For Each handleObject As String In objCatalog.ListHandles
            If IsNothing(handleObject) = True Then Continue For
            If handleObject.Trim.Length = 0 Then Continue For
            Dim idObject As ObjectId = AcadFunction.FuncReturnObjectID(handleObject)
            If IsNothing(idObject) = False Then
                If idObject.IsNull = False Then
                    idColl.Add(idObject)
                End If
            End If
        Next
        'если коллекция не пустая тогда копируем объекты
        If idColl.Count > 0 Then
            'новый чертеж по стандартному шаблону acadiso
            Using newDb As Database = New Database(True, True)
                Dim newDbMsId As ObjectId = SymbolUtilityServices.GetBlockModelSpaceId(newDb)
                Dim idMap As IdMapping = New IdMapping()
                db.WblockCloneObjects(idColl, newDbMsId, idMap, DuplicateRecordCloning.Replace, False)
                Dim prevDb As Database = HostApplicationServices.WorkingDatabase
                HostApplicationServices.WorkingDatabase = newDb
                newDb.UpdateExt(True)
                newDb.TileMode = True
                Using vTab As ViewportTable = TryCast(newDb.ViewportTableId.Open(OpenMode.ForRead), ViewportTable)
                    Dim acVptId As ObjectId = vTab("*Active")
                    Using vpTabRec As ViewportTableRecord = TryCast(acVptId.Open(OpenMode.ForWrite), ViewportTableRecord)
                        Dim scrRatio As Double = (vpTabRec.Width / vpTabRec.Height)
                        Dim matWCS2DCS As Matrix3d = Matrix3d.PlaneToWorld(vpTabRec.ViewDirection)
                        matWCS2DCS = Matrix3d.Displacement(vpTabRec.Target - Point3d.Origin) * matWCS2DCS
                        matWCS2DCS = Matrix3d.Rotation(-vpTabRec.ViewTwist, vpTabRec.ViewDirection, vpTabRec.Target) * matWCS2DCS
                        matWCS2DCS = matWCS2DCS.Inverse()
                        Dim extents As Extents3d = New Extents3d(newDb.Extmin, newDb.Extmax)
                        extents.TransformBy(matWCS2DCS)
                        Dim width As Double = (extents.MaxPoint.X - extents.MinPoint.X)
                        Dim height As Double = (extents.MaxPoint.Y - extents.MinPoint.Y)
                        Dim center As Point2d = New Point2d((extents.MaxPoint.X + extents.MinPoint.X) * 0.5, (extents.MaxPoint.Y + extents.MinPoint.Y) * 0.5)
                        If width > (height * scrRatio) Then height = width / scrRatio
                        vpTabRec.Height = height
                        vpTabRec.Width = height * scrRatio
                        vpTabRec.CenterPoint = center
                    End Using
                End Using
                HostApplicationServices.WorkingDatabase = prevDb
                newDb.SaveAs(strSave, DwgVersion.AC1021)
            End Using
        End If
    End Sub

    Private Sub DrawCatalogTableCAD(ByVal objCatalog As CadastralTools.CatalogCoordinatesObject)
        If IsNothing(objCatalog) = True Then Return
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim tb As New Autodesk.AutoCAD.DatabaseServices.Table()
            Dim ms As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            'количество колонок (аналог CheckBox8 эталона); 6-я колонка "Описание" —
            'дополнение по нашему требованию (в эталоне расширенных колонок 5)
            Dim countColumns As Integer = 3
            If objCatalog.ExtendedProperties = True Then countColumns = 6
            tb.TableStyle = db.Tablestyle
            tb.SetSize(3, countColumns)
            tb.SetRowHeight(8)
            tb.Rows(0).TextHeight = 2
            tb.Rows(1).TextHeight = 2
            tb.Rows(2).TextHeight = 2
            If objCatalog.ExtendedProperties = True Then
                tb.Columns(0).Width = 15 'номер точки
                tb.Columns(1).Width = 20 'X
                tb.Columns(2).Width = 20 'Y
                tb.Columns(3).Width = 40 'Метод определения
                tb.Columns(4).Width = 20 'Погрешность
                tb.Columns(5).Width = 45 'Описание обозначения точки
            Else
                tb.Columns(0).Width = 25 'номер точки
                tb.Columns(1).Width = 20 'X
                tb.Columns(2).Width = 20 'Y
            End If
            'точка вставки таблицы (аналог j = 0 / PointInsTbl1.X + 100 эталона)
            If boolAskPointInsert = False Then
                Dim ppr As PromptPointResult = ed.GetPoint("Укажите точку вставки таблицы: ")
                If ppr.Status <> PromptStatus.OK Then Return
                pointInsertTableCAD = ppr.Value
                boolAskPointInsert = True
            Else
                pointInsertTableCAD = New Autodesk.AutoCAD.Geometry.Point3d(pointInsertTableCAD.X + 100, pointInsertTableCAD.Y, 0)
            End If
            tb.Position = pointInsertTableCAD
            'шапка таблицы (аналог range/range1 эталона)
            Dim range As CellRange = CellRange.Create(tb, 1, 0, 2, 0)
            tb.MergeCells(range)
            tb.Cells(1, 0).Value = "Обозначение характерных точек границ"
            Dim range1 As CellRange = CellRange.Create(tb, 1, 1, 1, 2)
            tb.MergeCells(range1)
            tb.Cells(1, 1).Value = "Координаты, м"
            tb.Cells(2, 1).Value = "X, м"
            tb.Cells(2, 2).Value = "Y, м"
            If objCatalog.ExtendedProperties = True Then
                tb.Cells(1, 3).Value = "Метод определения координат"
                Dim range2 As CellRange = CellRange.Create(tb, 1, 3, 2, 3)
                tb.MergeCells(range2)
                tb.Cells(1, 4).Value = "Средняя квадратическая погрешность положения характерных точек (Mt), м"
                Dim range3 As CellRange = CellRange.Create(tb, 1, 4, 2, 4)
                tb.MergeCells(range3)
                tb.Cells(1, 5).Value = "Описание обозначения точки на местности (при наличии)"
                Dim range4 As CellRange = CellRange.Create(tb, 1, 5, 2, 5)
                tb.MergeCells(range4)
            End If
            ms.AppendEntity(tb)
            tr.AddNewlyCreatedDBObject(tb, True)
            'присваиваем таблице набор характеристик CADASTRAL_OBJECT:
            'Имя таблицы (LABEL) = "Каталог координат", Описание (Number) = условный номер участка
            Try
                Dim nameTablePS As String = "CADASTRAL_OBJECT"
                Dim idTablePS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(nameTablePS)
                If idTablePS.IsNull = True Then
                    'определение набора характеристик отсутствует — пробуем скопировать из шаблона
                    Dim asmPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                    Dim parentPatch1 As String = IO.Directory.GetParent(asmPath).FullName
                    Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                    Dim dirTemplates As String = parentPatch2 & "\FileResources\Sample\"
                    If IO.Directory.Exists(dirTemplates) = True Then
                        For Each fileDwg As String In IO.Directory.GetFiles(dirTemplates, "*.dwg")
                            If AcadFunction.FuncCopyTablePSDestDraw(fileDwg, nameTablePS) = True Then Exit For
                        Next
                    End If
                    idTablePS = FuncPropertySet.FuncReturnIDDictSetProperties(nameTablePS)
                End If
                If idTablePS.IsNull = False Then
                    Dim arrayPS As String(,) = Nothing
                    ReDim Preserve arrayPS(2, 1)
                    arrayPS(0, 0) = nameTablePS
                    arrayPS(1, 0) = "LABEL"
                    arrayPS(2, 0) = "Каталог координат"
                    arrayPS(0, 1) = nameTablePS
                    arrayPS(1, 1) = "Number"
                    arrayPS(2, 1) = objCatalog.ConditionNumber
                    FuncPropertySet.FuncAddDataSetProperties(tb, idTablePS)
                    FuncPropertySet.FuncWriteAllDataSetProperties(tb, arrayPS)
                End If
            Catch ex As System.Exception
            End Try
            'заголовок таблицы (аналог Cells(0, 0) эталона)
            tb.Cells(0, 0).Value = "Каталог координат поворотных точек объекта с условным номером: " & objCatalog.ConditionNumber
            'жирный шрифт заголовка (строка 0) — текстовый стиль с Bold = True
            Try
                Dim nameBoldStyle As String = "CatalogCoordinates_Bold"
                Dim idBoldStyle As ObjectId = ObjectId.Null
                Dim tst As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                If tst.Has(nameBoldStyle) Then
                    idBoldStyle = tst.Item(nameBoldStyle)
                Else
                    'создаем стиль на основе шрифта текущего текстового стиля чертежа
                    Dim baseRecord As TextStyleTableRecord = tr.GetObject(db.Textstyle, OpenMode.ForRead)
                    Dim baseFont As Autodesk.AutoCAD.GraphicsInterface.FontDescriptor = baseRecord.Font
                    Dim nameFont As String = baseFont.TypeFace
                    If nameFont.Trim.Length = 0 Then nameFont = "Arial"
                    Dim newRecord As TextStyleTableRecord = New TextStyleTableRecord()
                    newRecord.Name = nameBoldStyle
                    newRecord.Font = New Autodesk.AutoCAD.GraphicsInterface.FontDescriptor(nameFont, True, baseFont.Italic, baseFont.CharacterSet, baseFont.PitchAndFamily)
                    Dim tstW As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForWrite)
                    idBoldStyle = tstW.Add(newRecord)
                    tr.AddNewlyCreatedDBObject(newRecord, True)
                End If
                If idBoldStyle.IsNull = False Then
                    tb.Cells(0, 0).TextStyleId = idBoldStyle
                End If
            Catch ex As System.Exception
            End Try
            'записываем координаты (аналог цикла For j1 эталона, строки 872-928)
            Dim count As Integer = 3
            Dim numPoint As Integer = 1
            For Each row As CadastralTools.CatalogCoordinatesRow In objCatalog.Rows
                tb.InsertRows(count, 1, 1)
                tb.Rows(count).TextHeight = 2
                If row.kindRow = -1 Then
                    'остров (дырка) — пустая объединенная строка-разделитель
                    Dim rangeIsland As CellRange = CellRange.Create(tb, count, 0, count, countColumns - 1)
                    tb.MergeCells(rangeIsland)
                    tb.Cells(count, 0).Value = ""
                ElseIf row.kindRow = 1 Then
                    'строка контура
                    Dim rangeHeader As CellRange = CellRange.Create(tb, count, 0, count, countColumns - 1)
                    tb.MergeCells(rangeHeader)
                    tb.Cells(count, 0).Value = row.Number
                Else
                    'обозначение точки (с авто-нумерацией, как в CSV/Excel)
                    Dim namePoint As String = row.Number
                    If IsNothing(namePoint) = True Then namePoint = ""
                    namePoint = namePoint.Trim
                    Dim testName As String = namePoint
                    If testName.Length > 1 And IsNumeric(testName) = False Then
                        testName = Mid(testName, 2)
                    End If
                    If testName.Length = 0 OrElse Val(testName) = 0 Then
                        namePoint = numPoint.ToString
                        numPoint += 1
                    End If
                    tb.Cells(count, 0).Value = namePoint
                    'X записывается в колонку 2, Y — в колонку 1 (порядок эталона сохранен)
                    tb.Cells(count, 2).Value = CadastralTools.MathFunction.formatNumber(Math.Round(row.X, objCatalog.PrecisionPosition), objCatalog.PrecisionPosition)
                    tb.Cells(count, 1).Value = CadastralTools.MathFunction.formatNumber(Math.Round(row.Y, objCatalog.PrecisionPosition), objCatalog.PrecisionPosition)
                    If objCatalog.ExtendedProperties = True Then
                        Dim metod As String = row.Method
                        If IsNothing(metod) = True Then metod = ""
                        If metod.Trim.Length = 0 Then metod = "Геодезический метод"
                        tb.Cells(count, 3).Value = metod
                        Dim di As Double = 0
                        If IsNothing(row.Delta) = False Then
                            di = Val(row.Delta.Replace(".", ","))
                        End If
                        If di = 0 Then di = 0.1
                        tb.Cells(count, 4).Value = CadastralTools.MathFunction.formatNumber(Math.Round(di, objCatalog.PrecisionPosition), objCatalog.PrecisionPosition)
                        tb.Cells(count, 5).Value = row.FastDesk
                    End If
                End If
                count += 1
            Next
            'итоговая строка: площадь или протяженность (аналог CheckBox7/flagArea эталона)
            If objCatalog.OutTotal = True Then
                tb.InsertRows(count, 1, 1)
                tb.Rows(count).TextHeight = 2
                If objCatalog.IsRegion = True Then
                    tb.Cells(count, 0).Value = "Площадь: "
                    tb.Cells(count, 2).Value = Math.Round(objCatalog.TotalArea, objCatalog.PrecisionArea) & " кв.м"
                Else
                    tb.Cells(count, 0).Value = "Протяженность: "
                    tb.Cells(count, 2).Value = Math.Round(objCatalog.TotalLength, objCatalog.PrecisionPosition) & " м"
                End If
            End If
            tr.Commit()
        End Using
    End Sub
End Class
