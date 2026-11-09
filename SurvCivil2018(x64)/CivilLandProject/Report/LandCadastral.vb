Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Gis.Map.MapBook
Imports CadastralTools
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Geometries
Imports OSGeo.MapGuide
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class LandCadastralReport
    '=========================================================================================================
    '1. таблица с координатами участков постоянного и временного отводов в MS Excel, Word, Autocad
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
                        classCatalogCoordinates.CatalogExportCoordinates(dictCadastralObject, listBlockReference, formCatalogCoordinates, AddressOf DrawCatalogTableCAD, AddressOf ExportParcelDwgCAD, AddressOf ExportExcelWGS84)
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
    Private Sub ExportExcelWGS84(ByVal objCatalog As CadastralTools.CatalogCoordinatesObject)
        If IsNothing(objCatalog) = True Then Return
        Dim CoordSystemCode As String = CivilApplication.ActiveDocument.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
        Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
        Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
        Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
        Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
        Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
        Dim DrawingSysCoord As MgCoordinateSystem = Nothing
        Try
            DrawingSysCoord = csDict.GetCoordinateSystem(CoordSystemCode)
        Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
            MsgBox("Чертежу не назначена система координат!!! Макро прервано!!!")
            Exit Sub
        End Try
        Dim coordSystem As MgCoordinateSystem = csDict.GetCoordinateSystem("LL84")
        Dim UserCoord_to_PseudoMarcator As MgCoordinateSystemTransform = coordSysFactory.GetTransform(DrawingSysCoord, coordSystem)
        'запускаем Excel
        Dim xlApp As Excel.Application = Nothing
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        xlApp = DirectCast(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        xlBook = xlApp.Workbooks.Add()
        xlSheet = DirectCast(xlBook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim boolExVisibl As Boolean = False
        Dim StartRows As Integer = 1
        StartRows += 1
        Dim cell1 As Excel.Range 'первая ячейка в строке
        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
        Dim CellRng As Excel.Range 'массив ячеек
        xlSheet = xlBook.ActiveSheet
        If StartRows = 2 Then
            xlSheet.Columns.Range("A:A").ColumnWidth = 15
            xlSheet.Columns.Range("B:B").ColumnWidth = 15
            xlSheet.Columns.Range("C:C").ColumnWidth = 15
            xlSheet.Columns.Range("D:D").ColumnWidth = 15
            xlSheet.Columns.Range("E:E").ColumnWidth = 15
        End If
        Dim flagArea As Boolean = False
        cell1 = xlSheet.Cells(StartRows, 1) 'ссылка на первую ячейку
        cell2 = xlSheet.Cells(StartRows, 5) 'ссылка на 6 ячейку в строке
        CellRng = xlSheet.Range(cell1, cell2) 'выбираем строку
        CellRng.Merge(Type.Missing)
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        'условный номер
        Dim conditionalParcel As String = objCatalog.ConditionNumber
        Dim numberParcel As Integer = 1
        If Integer.TryParse(conditionalParcel.ToString, numberParcel) = False Then
            numberParcel = DataFormatting.GetNumberBeforeParenthesis(conditionalParcel)
        End If
        cell1.Value = "Условный номер (обозначение): " & numberParcel
        '===========================================================================
        cell1 = xlSheet.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
        cell2 = xlSheet.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
        CellRng = xlSheet.Range(cell1, cell2) 'выбираем строку
        CellRng.Merge(Type.Missing)
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "Обозначение" & vbLf & "характерных" & vbLf & "точек границ"
        CellRng.WrapText = True
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
        cell2 = xlSheet.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
        CellRng = xlSheet.Range(cell1, cell2) 'выбираем строку
        CellRng.Merge(Type.Missing)
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "Координаты МСК, м"
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 1, 4) 'ссылка на первую ячейку
        cell2 = xlSheet.Cells(StartRows + 1, 5) 'ссылка на 6 ячейку в строке
        CellRng = xlSheet.Range(cell1, cell2) 'выбираем строку
        CellRng.Merge(Type.Missing)
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "Координаты WGS 84, град."
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "X"
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "Y"
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 2, 4) 'ссылка на первую ячейку
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "B"
        '==============================================================================
        cell1 = xlSheet.Cells(StartRows + 2, 5) 'ссылка на первую ячейку
        cell1.Font.Name = "Times New Roman"
        cell1.Font.Bold = 1
        cell1.Font.Size = 12
        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        cell1.Value = "L"
        '====================================================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'записываем координаты
        Dim currentRow As Integer = StartRows + 3
        Dim countColumnsExcel As Integer = 5
        Dim numPoint As Integer = 1
        Dim numberFormatZn As String = "0.00"
        If objCatalog.PrecisionPosition = 0 Then
            numberFormatZn = "0"
        ElseIf objCatalog.PrecisionPosition = 1 Then
            numberFormatZn = "0.0"
        ElseIf objCatalog.PrecisionPosition = 2 Then
            numberFormatZn = "0.00"
        ElseIf objCatalog.PrecisionPosition = 3 Then
            numberFormatZn = "0.000"
        ElseIf objCatalog.PrecisionPosition = 4 Then
            numberFormatZn = "0.0000"
        Else
            numberFormatZn = "0.000000"
        End If
        For Each row As CatalogCoordinatesRow In objCatalog.Rows
            If row.kindRow = 1 Then
                'строка контура
                cell1 = xlSheet.Cells(currentRow, 1)
                cell2 = xlSheet.Cells(currentRow, countColumnsExcel)
                CellRng = xlSheet.Range(cell1, cell2)
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                cell1.Value = row.Number
                currentRow += 1
            ElseIf row.kindRow = -1 Then
                'остров (дырка) — пустая объединенная строка-разделитель
                cell1 = xlSheet.Cells(currentRow, 1)
                cell2 = xlSheet.Cells(currentRow, countColumnsExcel)
                CellRng = xlSheet.Range(cell1, cell2)
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                cell1.Value = ""
                currentRow += 1
            Else
                'координаты
                cell1 = xlSheet.Cells(currentRow, 1)
                cell2 = xlSheet.Cells(currentRow, countColumnsExcel)
                CellRng = xlSheet.Range(cell1, cell2)
                CellRng.Font.Name = "Times New Roman"
                CellRng.Font.Size = 12
                CellRng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                CellRng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                'обозначение точки (с авто-нумерацией, как в CSV)
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
                xlSheet.Cells(currentRow, 1).Value = namePoint
                'X записывается в колонку 3, Y — в колонку 2 (порядок эталона сохранен)
                cell1 = xlSheet.Cells(currentRow, 3)
                cell1.NumberFormat = numberFormatZn
                cell1.Value = row.X
                cell1 = xlSheet.Cells(currentRow, 2)
                cell1.NumberFormat = numberFormatZn
                cell1.Value = row.Y
                Dim UserCoord_to_PseudoMercatorRez As MgCoordinate = UserCoord_to_PseudoMarcator.Transform(Val(row.X), Val(row.Y))
                Dim L As Double = UserCoord_to_PseudoMercatorRez.X
                Dim B As Double = UserCoord_to_PseudoMercatorRez.Y
                cell1 = xlSheet.Cells(currentRow, 4)
                cell1.NumberFormat = ".000000"
                cell1.Value = B
                cell1 = xlSheet.Cells(currentRow, 5)
                cell1.NumberFormat = ".000000"
                cell1.Value = L
                currentRow += 1
            End If
        Next
        xlApp.Visible = True
    End Sub
    <CommandMethod("ListCadastralLandObject")>
    Public Sub ListCadastralLandObject()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ШАГ 1. выбираем примитив (в том числе из внешней ссылки)
        Dim cancelled As Boolean = False
        Dim outerPolygon As NetTopologySuite.Geometries.Polygon = SelectCadastralObject.SelectOuterPolygon(ed, cancelled)

        If cancelled Then
            MsgBox("Прервано пользователем.")
            Exit Sub
        End If

        If outerPolygon Is Nothing Then
            ' Полигон не удалось построить — обработайте по своему усмотрению
            MsgBox("Не удалось определить внешний полигон.")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'шаг 2. запрашиваем тип отчета (Земельные участки, ОКС, ЗОУИТ)
        Dim pko As New PromptKeywordOptions(vbLf & "Выберите тип объекта: ")
        pko.Keywords.Add("ЗемельныеУчастки")
        pko.Keywords.Add("ОКС")
        pko.Keywords.Add("ЗОУИТ")
        pko.Keywords.Default = "ЗемельныеУчастки"
        pko.AllowNone = False

        Dim resTypeObject As PromptResult = ed.GetKeywords(pko)
        If resTypeObject.Status <> PromptStatus.OK Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'шаг 3. запрашиваем тип отчета (Земельные участки, ОКС, ЗОУИТ)
        pko = New PromptKeywordOptions(vbLf & "Выберите тип отчёта: ")
        pko.Keywords.Add("Excel")
        pko.Keywords.Add("Word")
        pko.Keywords.Add("Autocad")
        pko.Keywords.Default = "Excel"
        pko.AllowNone = False

        Dim resTypeReport As PromptResult = ed.GetKeywords(pko)
        If resTypeReport.Status <> PromptStatus.OK Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'шаг 4. выбираем внешнюю ссылку
        Dim peo As New PromptEntityOptions(vbLf & "Выберите внешнюю ссылку (XREF): ")
        peo.SetRejectMessage(vbLf & "Это не внешняя ссылка. Выберите XREF.")
        peo.AddAllowedClass(GetType(BlockReference), False)

        Dim per As PromptEntityResult = ed.GetEntity(peo)
        If per.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "Прервано пользователем.")
            Return
        End If
        Dim dbXRec As Database = per.ObjectId.Database
        Dim dictSelectAllObject As Dictionary(Of String, List(Of Geometry)) = SelectCadastralObject.selectAllCadastralObjectByDesk(db)
        If dictSelectAllObject.Count > 0 Then
            If resTypeObject.StringResult Like "земельныеУчастки" Then
                Dim dictParcels As Dictionary(Of String, List(Of Geometry)) = New Dictionary(Of String, List(Of Geometry))
                Dim nameTypeObject As String = "Кадастровый участок"
                dictParcels = CadastralTools.SelectCadastralObject.getCadastralObjectByDesk(dictSelectAllObject, nameTypeObject)
            ElseIf resTypeObject.StringResult Like "ОКС" Then

            ElseIf resTypeObject.StringResult Like "ЗОУИТ" Then

            End If



        End If



        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr

            End Using
        End Using

    End Sub

End Class
