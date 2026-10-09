Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class ImportDataGeology
    '=============================================================================================
    'загрузка в базу данных геологических скважин
    <CommandMethod("ImportDataGeology")>
    Public Sub ImportDataGeology()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim DeskCogo As String = "GSKV"
        Dim PS_hole As String = "GEOLOGY_HOLE"
        Dim idTablePs As ObjectId = Nothing
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим путь к сборке
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTemplateDir As String = parentPatch2 & "\FileResources\Sample\"
        'путь к шаблону
        Dim PutchTemplateDwg As String = PutchTemplateDir & "Стандарт_ADSK.dwg"

        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                      OpenMode.ForWrite)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Выбираем файл
                Dim ArrayDataWrite As String(,) = Nothing
                Dim CountArrayDataWrite As Integer = 0
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                'opFile.InitialDirectory = "d:\Примеры NET\Примеры\Геология\"
                opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                Dim fFile1 As String = opFile.FileName
                If IsNothing(fFile1) = True Then Exit Sub
                If fFile1.Trim.Length = 0 Then Exit Sub
                Dim xlApp As Excel.Application = Nothing
                Dim xlBook As Excel.Workbook = Nothing
                Dim xlSheets As Excel.Worksheet = Nothing
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Open(fFile1)
                Dim xlsheet As Excel.Worksheet
                xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                Dim ArrayData As String(,) = Nothing
                Dim CountArrayData As Integer = 0
                Dim countErrData As Integer = 0
                'читаем данные
                '1-имя скважины,2-дата проходки, 3-Восточное положение, 4-Северное положение, 5-отметка, 6-Глубина скважины
                For indRows As Integer = 1 To xlsheet.Rows.Count - 1
                    Dim nameGWell As String = xlsheet.Cells(indRows, 1).value
                    Dim dataGWell As String = xlsheet.Cells(indRows, 2).value
                    Dim x1 As String = xlsheet.Cells(indRows, 3).value
                    Dim y1 As String = xlsheet.Cells(indRows, 4).value
                    Dim h1 As String = xlsheet.Cells(indRows, 5).value
                    Dim depth As String = xlsheet.Cells(indRows, 6).value
                    If IsNothing(nameGWell) = False Then
                        If IsNothing(x1) = False And IsNothing(y1) = False And IsNothing(h1) = False Then
                            If IsNumeric(x1) = True And IsNumeric(y1) = True And IsNumeric(h1) = True Then
                                ReDim Preserve ArrayData(5, CountArrayData)
                                ArrayData(0, CountArrayData) = nameGWell
                                ArrayData(1, CountArrayData) = dataGWell
                                ArrayData(2, CountArrayData) = x1
                                ArrayData(3, CountArrayData) = y1
                                ArrayData(4, CountArrayData) = h1
                                ArrayData(5, CountArrayData) = depth
                                CountArrayData += 1
                                countErrData = 0
                            Else
                                countErrData += 1
                            End If
                        Else
                            countErrData += 1
                        End If
                    Else
                        countErrData += 1
                    End If
                    If countErrData > 20 Then
                        Exit For
                    End If
                Next indRows
                xlBook.Close()
                xlsheet = Nothing
                xlBook = Nothing
                xlApp.Quit()
                xlApp = Nothing
                If IsArray(ArrayData) = False Then
                    MsgBox("Не удалось прочитать таблицу Excel. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'сортируем массив, и проверяем чтобы имена скважин были уникальными
                Dim rnduser As Random = New Random
                If ArrayData.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayData, 0)
                    Dim firstName As String = ArrayData(0, 0)
                    For i As Integer = 1 To ArrayData.GetUpperBound(1)
                        Dim name As String = ArrayData(0, i)
                        If name Like firstName Then
                            ArrayData(0, i) = name & "_" & rnduser.Next
                        Else
                            firstName = name
                        End If
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'копируем стили
                'копируем ключ-описатель
                Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(acCurDb).SearchOrder
                Dim boolKey As Boolean = False
                If pdKeySetIds.Count > 0 Then
                    For Each idKeySet As ObjectId In pdKeySetIds
                        Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = tr.GetObject(idKeySet, OpenMode.ForRead)
                        Dim idDesk As ObjectId = Nothing
                        Try
                            idDesk = DeskKeySet.Item(DeskCogo)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idDesk) = False Then
                            If idDesk.IsNull = False Then
                                boolKey = True
                                Exit For
                            Else
                                boolKey = False
                            End If
                        Else
                            boolKey = False
                        End If
                    Next
                End If
                If boolKey = False Then
                    Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplateDwg, DeskCogo)
                End If
                '===================================================================================================
                'копируем таблицу PS
                Dim dictPropSetDef As Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                If dictPropSetDef.Has(PS_hole, tr) = False Then
                    Dim boolInsPs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, PS_hole)
                    idTablePs = dictPropSetDef.GetAt(PS_hole)
                Else
                    idTablePs = dictPropSetDef.GetAt(PS_hole)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'создаем точки COGO
                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                    Try
                        Dim Name As String = ArrayData(0, i)
                        Dim data_hole As String = ArrayData(1, i)
                        Dim y As Double = Val(ArrayData(3, i))
                        Dim x As Double = Val(ArrayData(2, i))
                        Dim h As Double = Val(ArrayData(4, i))
                        Dim depth As String = ArrayData(5, i)
                        Dim pt As Point3d = New Point3d(x, y, h)
                        Dim IDcogoPoint As ObjectId = FuncSurvey.FuncCreateCoGoPoint(pt, DeskCogo,, h, Name)
                        '=================================================================================
                        'формируем массив для записи
                        Dim ArrayOBJ1 As String(,) = Nothing
                        ReDim Preserve ArrayOBJ1(2, 2)
                        ArrayOBJ1(0, 0) = "GEOLOGY_HOLE"
                        ArrayOBJ1(1, 0) = "DATA"
                        ArrayOBJ1(2, 0) = data_hole

                        ArrayOBJ1(0, 1) = "GEOLOGY_HOLE"
                        ArrayOBJ1(1, 1) = "DEPT"
                        ArrayOBJ1(2, 1) = depth

                        ArrayOBJ1(0, 2) = "GEOLOGY_HOLE"
                        ArrayOBJ1(1, 2) = "ZERO_LEVEL"
                        ArrayOBJ1(2, 2) = h

                        Dim cogoPt As Autodesk.Civil.DatabaseServices.CogoPoint = IDcogoPoint.GetObject(OpenMode.ForWrite)
                        If IsNothing(idTablePs) = False Then
                            If idTablePs.IsNull = False Then
                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(cogoPt, idTablePs)
                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(cogoPt, ArrayOBJ1)
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next i
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub

    '=============================================================================================
    'построение литологических колонок
    <CommandMethod("CreateLithologicalWells")>
    Public Sub CreateLithologicalWells()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim DeskCogo As String = "GSKV"
        Dim PS_hole As String = "GEOLOGY_LITHOLOGICAL_DESK"
        Dim idTablePs As ObjectId = Nothing
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим путь к сборке
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTemplateDir As String = parentPatch2 & "\FileResources\Sample\"
        'путь к шаблону
        Dim PutchTemplateDwg As String = PutchTemplateDir & "Стандарт_ADSK.dwg"
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                      OpenMode.ForWrite)
                'забираем массив с колонками геологическими
                Dim ArrayData As String(,) = Nothing
                Dim CountArrayData As Integer = 0

                '===================================================================================================
                'копируем таблицу PS
                Dim dictPropSetDef As Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                If dictPropSetDef.Has(PS_hole, tr) = False Then
                    Dim boolInsPs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, PS_hole)
                    idTablePs = dictPropSetDef.GetAt(PS_hole)
                Else
                    idTablePs = dictPropSetDef.GetAt(PS_hole)
                End If
                '==================================================================================================================
                'выбираем объекты
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If TypeOf acEnt Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                Dim ptCogo As Autodesk.Civil.DatabaseServices.CogoPoint = acEnt
                                If ptCogo.RawDescription Like DeskCogo Then
                                    ReDim Preserve ArrayData(3, CountArrayData)
                                    ArrayData(0, CountArrayData) = ptCogo.PointName
                                    ArrayData(1, CountArrayData) = ptCogo.Location.X
                                    ArrayData(2, CountArrayData) = ptCogo.Location.Y
                                    ArrayData(3, CountArrayData) = ptCogo.Elevation
                                    CountArrayData += 1
                                End If
                            End If
                        End If
                    Next
                End If

                If IsArray(ArrayData) = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Выбираем файл
                    Dim ArrayDataWrite As String(,) = Nothing
                    Dim CountArrayDataWrite As Integer = 0
                    Dim opFile As New System.Windows.Forms.OpenFileDialog()
                    'opFile.InitialDirectory = "d:\Примеры NET\Примеры\Геология\"
                    opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                    opFile.RestoreDirectory = True
                    opFile.ShowDialog()
                    Dim fFile1 As String = opFile.FileName
                    If IsNothing(fFile1) = True Then Exit Sub
                    If fFile1.Trim.Length = 0 Then Exit Sub
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Open(fFile1)
                    Dim xlsheet As Excel.Worksheet
                    xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
                    Dim countErrData As Integer = 0
                    Dim nameHoleGlob As String = "0"
                    Dim indColor As Integer = 1
                    'читаем данные
                    '1-имя скважины,2-индекс ИГИ, 3-геологический индекс,4-описание колонким,5-превышение по верху,6-превышение по низу
                    For indRows As Integer = 1 To xlsheet.Rows.Count - 1
                        Try
                            Dim nameGWell As String = xlsheet.Cells(indRows, 1).value
                            Dim GWell_num As String = xlsheet.Cells(indRows, 2).value
                            Dim GWell_ind As String = xlsheet.Cells(indRows, 3).value
                            Dim GWell_desk As String = xlsheet.Cells(indRows, 4).value
                            Dim GWell_top As String = xlsheet.Cells(indRows, 5).value
                            Dim GWell_down As String = xlsheet.Cells(indRows, 6).value
                            If IsNothing(nameGWell) = False Then
                                If IsNothing(GWell_top) = False And IsNothing(GWell_down) = False Then
                                    If IsNumeric(GWell_down) = True And IsNumeric(GWell_top) = True Then
                                        If Not (nameGWell Like nameHoleGlob) Then
                                            indColor = 1
                                        End If
                                        nameHoleGlob = nameGWell
                                        Dim top As Double = Val(GWell_top)
                                        Dim down As Double = Val(GWell_down)
                                        Dim heightSolid As Double = down - top
                                        If heightSolid > 0 Then
                                            'строим разрез
                                            '1. находим скважину в списке
                                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                If ArrayData(0, i).Trim Like nameGWell.Trim Then
                                                    Dim x As Double = Val(ArrayData(1, i))
                                                    Dim y As Double = Val(ArrayData(2, i))
                                                    Dim z As Double = Val(ArrayData(3, i))
                                                    Dim pt As Point3d = New Point3d(x, y, z - top - heightSolid / 2)
                                                    'строим разрез
                                                    Dim sol As Solid3d = AcadFunction.FuncCreateCylindr3d(heightSolid, 0.2, 0.2, pt)
                                                    '=================================================================================
                                                    'формируем массив для записи
                                                    Dim ArrayOBJ1 As String(,) = Nothing
                                                    ReDim Preserve ArrayOBJ1(2, 6)
                                                    ArrayOBJ1(0, 0) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 0) = "NAME_HOLE"
                                                    ArrayOBJ1(2, 0) = nameGWell

                                                    ArrayOBJ1(0, 1) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 1) = "NUMB"
                                                    ArrayOBJ1(2, 1) = GWell_num

                                                    ArrayOBJ1(0, 2) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 2) = "CODE"
                                                    ArrayOBJ1(2, 2) = GWell_ind

                                                    ArrayOBJ1(0, 3) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 3) = "Depth_Base"
                                                    ArrayOBJ1(2, 3) = z - down

                                                    ArrayOBJ1(0, 4) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 4) = "Depth_Top"
                                                    ArrayOBJ1(2, 4) = z - top

                                                    ArrayOBJ1(0, 5) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 5) = "Height_hole"
                                                    ArrayOBJ1(2, 5) = heightSolid

                                                    ArrayOBJ1(0, 6) = "GEOLOGY_LITHOLOGICAL_DESK"
                                                    ArrayOBJ1(1, 6) = "Description"
                                                    ArrayOBJ1(2, 6) = GWell_desk

                                                    If IsNothing(idTablePs) = False Then
                                                        If idTablePs.IsNull = False Then
                                                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(sol, idTablePs)
                                                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(sol, ArrayOBJ1)
                                                        End If
                                                    End If
                                                    sol.ColorIndex = indColor
                                                    indColor += 1
                                                    countErrData = 0
                                                End If
                                            Next i
                                        End If
                                    Else
                                        countErrData += 1
                                    End If
                                Else
                                    countErrData += 1
                                End If
                            Else
                                countErrData += 1
                            End If
                            If countErrData > 20 Then
                                Exit For
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next indRows
                    xlBook.Close()
                    xlsheet = Nothing
                    xlBook = Nothing
                    xlApp.Quit()
                    xlApp = Nothing
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub

    '=============================================================================================
    'построение поверхности грунтов
    <CommandMethod("CreateGeologySurface")>
    Public Sub CreateGeologySurface()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ArrayTBL As String(,) = Nothing
        ReDim Preserve ArrayTBL(1, 0)
        ArrayTBL(0, 0) = "GEOLOGY_LITHOLOGICAL_DESK"
        ArrayTBL(1, 0) = ""
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим путь к сборке
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTemplateDir As String = parentPatch2 & "\FileResources\Sample\"
        'путь к шаблону
        Dim PutchTemplateDwg As String = PutchTemplateDir & "Стандарт_ADSK.dwg"

        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Try
            Using tr
                Dim acBlkTbl As BlockTable
                acBlkTbl = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace),
                                      OpenMode.ForWrite)
                '==================================================================================================================
                'выбираем объекты
                Dim ArrayPointGroup As String() = Nothing
                Dim CountArrayPointGroup As Integer = 0

                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If TypeOf acEnt Is Solid3d Then
                                Dim solGeo As Solid3d = acEnt
                                Dim ArrayData As String(,) = Nothing
                                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(solGeo, ArrayData, ArrayTBL)
                                If IsArray(ArrayData) = True Then
                                    Dim indGeo As String = ArrayData(2, 3).Trim
                                    Dim Hdown As Double = Val(ArrayData(2, 5))
                                    Dim extSolid As Extents3d = solGeo.GeometricExtents
                                    Dim pt1 As Point3d = New Point3d((extSolid.MaxPoint.X + extSolid.MinPoint.X) / 2, (extSolid.MaxPoint.Y + extSolid.MinPoint.Y) / 2, Hdown)
                                    If IsArray(ArrayPointGroup) = False Then
                                        ReDim Preserve ArrayPointGroup(CountArrayPointGroup)
                                        ArrayPointGroup(CountArrayPointGroup) = indGeo
                                        CountArrayPointGroup += 1
                                    Else
                                        Dim boolFindName As Boolean = False
                                        For i As Integer = 0 To ArrayPointGroup.Length - 1
                                            If ArrayPointGroup(i) Like indGeo Then
                                                boolFindName = True
                                                Exit For
                                            End If
                                        Next i
                                        If boolFindName = False Then
                                            ReDim Preserve ArrayPointGroup(CountArrayPointGroup)
                                            ArrayPointGroup(CountArrayPointGroup) = indGeo
                                            CountArrayPointGroup += 1
                                        End If
                                    End If
                                    Dim IDcogoPoint As ObjectId = FuncSurvey.FuncCreateCoGoPoint(pt1, indGeo,, Hdown)
                                End If
                            End If
                        End If
                    Next
                End If

                If IsArray(ArrayPointGroup) = True Then
                    Dim ptGroupCol As PointGroupCollection = CivilDoc.PointGroups
                    Dim idTinSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
                    For i As Integer = 0 To ArrayPointGroup.Length - 1
                        Dim namePointGroup As String = ArrayPointGroup(i)
                        Dim idGroup As ObjectId = Nothing
                        If ptGroupCol.Contains(namePointGroup) = False Then
                            idGroup = ptGroupCol.Add(namePointGroup)
                            If IsNothing(idGroup) = False Then
                                If idGroup.IsNull = False Then
                                    Dim userGroup As PointGroup = idGroup.GetObject(OpenMode.ForWrite)

                                    Dim standardQuery As StandardPointGroupQuery = New StandardPointGroupQuery()
                                    standardQuery.IncludeRawDescriptions = namePointGroup
                                    userGroup.SetQuery(standardQuery)
                                    userGroup.Update()

                                End If
                            End If
                        Else
                            For Each id As ObjectId In ptGroupCol
                                Dim userGroup As PointGroup = id.GetObject(OpenMode.ForRead)
                                If userGroup.Name.Trim Like namePointGroup Then
                                    idGroup = id
                                    Exit For
                                End If
                            Next
                        End If
                        '==========================================================================================
                        'создаем поверхности
                        Dim boolTin As Boolean = False
                        For Each id As ObjectId In idTinSurf
                            Dim userTin As TinSurface = id.GetObject(OpenMode.ForRead)
                            If userTin.Name.Trim Like namePointGroup Then
                                boolTin = True
                                Exit For
                            End If
                        Next

                        If boolTin = False Then
                            Dim IDtinSurf1 As ObjectId = TinSurface.Create(namePointGroup, CivilDoc.Styles.SurfaceStyles.Item(0))
                            If IsNothing(IDtinSurf1) = False Then
                                If IDtinSurf1.IsNull = False Then
                                    Dim tinSufr1 As TinSurface = IDtinSurf1.GetObject(OpenMode.ForWrite)
                                    tinSufr1.PointGroupsDefinition.AddPointGroup(idGroup)
                                End If
                            End If
                        End If
                    Next i
                End If
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
    End Sub
End Class
