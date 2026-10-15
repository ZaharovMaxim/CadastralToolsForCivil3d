Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
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
                        classCatalogCoordinates.CatalogExportCoordinates(dictCadastralObject, listBlockReference, formCatalogCoordinates, AddressOf DrawCatalogTableCAD)
                    End If
                End If
                tr.Commit()
            End Using
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод каталога координат в виде таблицы AutoCAD (делегат для CadastralTools.CatalogCoordinates)
    'таблицы размещаются последовательно с шагом 120 мм по оси X
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private pointInsertTableCAD As Autodesk.AutoCAD.Geometry.Point3d = Nothing
    Private boolAskPointInsert As Boolean = False

    Private Sub DrawCatalogTableCAD(ByVal objCatalog As CadastralTools.CatalogCoordinatesObject)
        If IsNothing(objCatalog) = True Then Return
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        'точка вставки таблицы
        If boolAskPointInsert = False Then
            Dim ppr As PromptPointResult = ed.GetPoint("Укажите точку вставки таблицы: ")
            If ppr.Status <> PromptStatus.OK Then Return
            pointInsertTableCAD = ppr.Value
            boolAskPointInsert = True
        Else
            pointInsertTableCAD = New Autodesk.AutoCAD.Geometry.Point3d(pointInsertTableCAD.X + 120, pointInsertTableCAD.Y, 0)
        End If
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim tb As New Autodesk.AutoCAD.DatabaseServices.Table()
            Dim ms As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            'количество колонок
            Dim countColumns As Integer = 3
            If objCatalog.ExtendedProperties = True Then countColumns = 6
            tb.TableStyle = db.Tablestyle
            tb.SetSize(3, countColumns)
            tb.SetRowHeight(8)
            tb.Rows(0).TextHeight = 2
            tb.Rows(1).TextHeight = 2
            tb.Rows(2).TextHeight = 2
            tb.Columns(0).Width = 25 'номер точки
            tb.Columns(1).Width = 20 'X
            tb.Columns(2).Width = 20 'Y
            If countColumns = 6 Then
                tb.Columns(3).Width = 45 'метод определения координат
                tb.Columns(4).Width = 45 'погрешность
                tb.Columns(5).Width = 45 'описание обозначения точки
            End If
            tb.Position = pointInsertTableCAD
            'шапка таблицы (по образцу: "Обозначение характерных точек границ" + "Координаты, м" / X, Y)
            Dim range As CellRange = CellRange.Create(tb, 1, 0, 2, 0)
            tb.MergeCells(range)
            tb.Cells(1, 0).Value = "Обозначение характерных точек границ"
            Dim range1 As CellRange = CellRange.Create(tb, 1, 1, 1, 2)
            tb.MergeCells(range1)
            tb.Cells(1, 1).Value = "Координаты, м"
            tb.Cells(2, 1).Value = "X, м"
            tb.Cells(2, 2).Value = "Y, м"
            If countColumns = 6 Then
                tb.Cells(1, 3).Value = "Метод определения координат характерной точки"
                Dim range2 As CellRange = CellRange.Create(tb, 1, 3, 2, 3)
                tb.MergeCells(range2)
                tb.Cells(1, 4).Value = "Средняя квадратическая погрешность положения характерной точки (Мt), м"
                Dim range3 As CellRange = CellRange.Create(tb, 1, 4, 2, 4)
                tb.MergeCells(range3)
                tb.Cells(1, 5).Value = "Описание обозначения точки на местности (при наличии)"
                Dim range4 As CellRange = CellRange.Create(tb, 1, 5, 2, 5)
                tb.MergeCells(range4)
            End If
            'заголовок таблицы
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
            'строки координат
            Dim count As Integer = 3
            For Each row As CadastralTools.CatalogCoordinatesRow In objCatalog.Rows
                tb.InsertRows(count, 1, 1)
                tb.Rows(count).TextHeight = 2
                If row.kindRow = 1 Then
                    Dim rangeHeader As CellRange = CellRange.Create(tb, count, 0, count, countColumns - 1)
                    tb.MergeCells(rangeHeader)
                    tb.Cells(count, 0).Value = row.Number
                ElseIf row.kindRow = -1 Then
                    'разделитель островка — пустая строка
                Else
                    tb.Cells(count, 0).Value = row.Number
                    tb.Cells(count, 1).Value = CadastralTools.MathFunction.formatNumber(Math.Round(row.X, objCatalog.PrecisionPosition), objCatalog.PrecisionPosition)
                    tb.Cells(count, 2).Value = CadastralTools.MathFunction.formatNumber(Math.Round(row.Y, objCatalog.PrecisionPosition), objCatalog.PrecisionPosition)
                    If countColumns = 6 Then
                        tb.Cells(count, 3).Value = row.Method
                        Dim strDelta As String = row.Delta
                        If strDelta.Trim.Length = 0 Then strDelta = "0.1"
                        tb.Cells(count, 4).Value = strDelta
                        tb.Cells(count, 5).Value = row.FastDesk
                    End If
                End If
                count += 1
            Next
            'итоговая строка (площадь или протяженность)
            tb.InsertRows(count, 1, 1)
            tb.Rows(count).TextHeight = 2
            If objCatalog.IsRegion = True Then
                tb.Cells(count, 0).Value = "Площадь: " & Math.Round(objCatalog.TotalArea, 0) & " кв.м"
            Else
                tb.Cells(count, 0).Value = "Протяженность: " & Math.Round(objCatalog.TotalLength, 2) & " м"
            End If
            ms.AppendEntity(tb)
            tr.AddNewlyCreatedDBObject(tb, True)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'присваиваем таблице набор характеристик CADASTRAL_OBJECT:
            'Имя таблицы (LABEL) = "Каталог координат", Описание (Number) = условный номер участка
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
            tr.Commit()
        End Using
    End Sub
End Class
