Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop

Public Class SurveyReport
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==================================================================================================
    'ОТЧЕТЫ ДЛЯ ГЕОДЕЗИИ
    '==================================================================================================
    'ведомость состояния пунктов государственной геодезической сети
    <CommandMethod("VedStateGeodeticPoints")>
    Public Sub VedStateGeodeticPoints()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
        Dim ODTable As String = "GEODETIC_network"
        Dim formSyrv As FormReportSurvey = New FormReportSurvey
        formSyrv.ShowDialog()

        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'ищем пути для таблиц
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Ведомость состояния пунктов государственной геодезической сети.docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            End If
            If File.Exists(PutchTbl) = False Then
                Dim opFile As New System.Windows.Forms.OpenFileDialog()
                opFile.InitialDirectory = "c:\"
                opFile.Filter = "Word files (*.docx)|*.docx"
                opFile.RestoreDirectory = True
                opFile.ShowDialog()
                PutchTbl = opFile.FileName
                If File.Exists(PutchTbl) = False Then
                    Exit Sub
                End If
            End If
            'открываем word
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
            End Try

            Try
                Dim PutchDoc As String = PutchTbl
                Dim rng As Random = New Random
                Dim Numbrng As Integer = rng.Next
                Dim DestDoc As String = TempDir & "Ведомость состояния пунктов ГГС_" & Numbrng & ".docx"
                Try
                    If File.Exists(PutchDoc) = True Then
                        File.Copy(PutchDoc, DestDoc)
                    End If
                Catch ex As UnauthorizedAccessException
                Catch ex As ArgumentException
                Catch ex As PathTooLongException
                Catch ex As DirectoryNotFoundException
                Catch ex As FileNotFoundException
                Catch ex As IOException
                Catch ex As NotSupportedException
                End Try
                If File.Exists(DestDoc) = True Then
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(DestDoc)
                Else
                    MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                    WApp.Quit()
                    WApp = Nothing
                    locker.Dispose()
                    Exit Sub
                End If
            Catch
                MsgBox("Не удалось найти выбранную таблицу!!!")
                locker.Dispose()
                WDoc = Nothing
                WApp = Nothing
                Exit Sub
            End Try

            'забираем в массив пункты
            Dim ArrayWord As String(,) = Nothing
            Dim CountArrayWord As Integer = 0
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            '==================================================================================================================
            'выбираем объекты
            Dim ObjectEntColl As ObjectIdCollection = New ObjectIdCollection
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            Dim boolFindTbl As Boolean = FuncPropertySet.FuncFindPSetTable(acEnt, ODTable)
                            If boolFindTbl = True Then
                                ObjectEntColl.Add(acEnt.Id)
                            End If
                        End If
                    End If
                Next
            End If
            If ObjectEntColl.Count = 0 Then
                MsgBox("Фильтр не выявил ни одного объекта с Набором характеристик " & ODTable & ". Макро прервано!!!")
                Exit Sub
            End If
            'перебираем объекты
            For Each id As ObjectId In ObjectEntColl
                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = id.GetObject(OpenMode.ForRead)
                'новая запись в массиве
                Dim namePoint As String = ""
                Dim klass As String = ""
                Dim typeCentre As String = ""
                Dim numbCatalog As String = "-"
                Dim StatCentre As String = "-"
                Dim StatDecor As String = "-"
                Dim DataEx As String = "-"
                ReDim Preserve ArrayWord(5, CountArrayWord)
                'считываем расширенные данные
                Dim ArrayData As String(,) = Nothing
                Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(acEnt, ArrayData, ODTable)
                If IsArray(ArrayData) = True Then
                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                        Dim NameField As String = ArrayData(1, i)
                        Dim valStr As String = ArrayData(2, i)
                        If IsNothing(NameField) = False Then
                            If IsNothing(valStr) = False Then
                                If NameField Like "NUMB_CATALOG" Then
                                    'номер по каталогу
                                    If valStr.Trim.Length > 0 Then
                                        numbCatalog = valStr
                                    End If
                                ElseIf NameField Like "NAME" Then
                                    'наименование пункта
                                    If valStr.Trim.Length > 0 Then
                                        namePoint = valStr
                                    End If
                                ElseIf NameField Like "CLASS_NETWORK" Then
                                    'класс сети
                                    If valStr.Trim.Length > 0 Then
                                        klass = valStr
                                    End If
                                ElseIf NameField Like "TYPE_CENTRE" Then
                                    'тип центра
                                    If valStr.Trim.Length > 0 Then
                                        typeCentre = valStr
                                    End If
                                ElseIf NameField Like "STATE_CENTRE" Then
                                    If valStr.Trim.Length > 0 Then
                                        StatCentre = valStr
                                    End If
                                ElseIf NameField Like "STATE_DECORATION" Then
                                    If valStr.Trim.Length > 0 Then
                                        StatDecor = valStr
                                    End If
                                ElseIf NameField Like "DATA_EXAMINATION" Then
                                    If valStr.Trim.Length > 0 Then
                                        DataEx = valStr
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
                '==================================================================================
                'смотрим объекты
                If TypeOf acEnt Is CogoPoint Then
                    Dim PtCoGo As CogoPoint = acEnt
                    If namePoint Like "" Or namePoint.Trim Like "-" Then
                        Dim NamePt As String = PtCoGo.PointName
                        If IsNothing(NamePt) = False Then
                            If NamePt.Trim.Length > 0 Then
                                namePoint = NamePt
                            End If
                        End If
                    End If
                    If numbCatalog.Trim.Length = 0 Then
                        numbCatalog = FuncSurvey.FuncReadUDPValue(PtCoGo, "NUMB_CATALOG", ODTable)
                    End If
                    If namePoint.Trim.Length = 0 Then
                        namePoint = FuncSurvey.FuncReadUDPValue(PtCoGo, "NAME", ODTable)
                    End If
                    If klass.Trim.Length = 0 Then
                        klass = FuncSurvey.FuncReadUDPValue(PtCoGo, "CLASS_NETWORK", ODTable)
                    End If

                    Dim NameRez As String = ""
                    If Not (namePoint Like "-") Then
                        NameRez = namePoint
                    End If
                    If Not (klass Like "") Then
                        If NameRez.Trim.Length > 0 Then
                            NameRez = NameRez & vbLf & klass
                        Else
                            NameRez = klass
                        End If
                    End If
                    If Not (typeCentre Like "") Then
                        If NameRez.Trim.Length > 0 Then
                            NameRez = NameRez & vbLf & typeCentre
                        Else
                            NameRez = typeCentre
                        End If
                    End If
                    ArrayWord(0, CountArrayWord) = numbCatalog
                    ArrayWord(1, CountArrayWord) = NameRez
                    ArrayWord(2, CountArrayWord) = StatCentre
                    ArrayWord(3, CountArrayWord) = StatDecor
                    ArrayWord(4, CountArrayWord) = DataEx
                    If StatCentre Like "Сохранился" Then
                        ArrayWord(5, CountArrayWord) = "Пригоден для производства ИГДИ"
                    Else
                        ArrayWord(5, CountArrayWord) = "Не пригоден для производства ИГДИ"
                    End If
                ElseIf TypeOf acEnt Is BlockReference Then
                    Dim blk As BlockReference = acEnt
                    If namePoint.Trim.Length = 0 Then
                        namePoint = AcadFunction.FuncReadValueAttr(blk, "NAME")
                    End If
                End If
                '=====================================================================================
                'заполняем массив
                ArrayWord(0, CountArrayWord) = numbCatalog
                ArrayWord(1, CountArrayWord) = namePoint
                ArrayWord(2, CountArrayWord) = StatCentre
                ArrayWord(3, CountArrayWord) = StatDecor
                ArrayWord(4, CountArrayWord) = DataEx
                If StatCentre Like "Сохранился" Then
                    ArrayWord(5, CountArrayWord) = "Пригоден для производства ИГДИ"
                Else
                    ArrayWord(5, CountArrayWord) = "Не пригоден для производства ИГДИ"
                End If
                CountArrayWord += 1
            Next
            '===========================================================================================
            'заполняем таблицу
            If IsArray(ArrayWord) = True Then
                    For i As Integer = 0 To ArrayWord.GetUpperBound(1)
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            tabl.Cell(countTableword, 1).Range.Text = ArrayWord(0, i) 'номер по каталогу
                            tabl.Cell(countTableword, 2).Range.Text = ArrayWord(1, i) 'наименование
                            tabl.Cell(countTableword, 3).Range.Text = ArrayWord(2, i) 'состояние ценра
                            tabl.Cell(countTableword, 4).Range.Text = ArrayWord(3, i) 'наружнее оформление
                            tabl.Cell(countTableword, 5).Range.Text = ArrayWord(4, i) 'дата обследования
                            tabl.Cell(countTableword, 6).Range.Text = ArrayWord(5, i) 'Заключение
                            Exit For
                        Next
                    Next i
                End If
                WApp.Visible = True
            End Using
        locker.Dispose()
        WDoc = Nothing
        WApp = Nothing
    End Sub
    '==================================================================================================
    'Каталог координат и высот пунктов государственной опорной геодезической сети
    <CommandMethod("CatalogCoordNetworkPoints")>
    Public Sub CatalogCoordNetworkPoints()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'ищем пути для таблиц
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Каталог координат и высот пунктов опорной геодезической сети.docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            Else
                If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Каталог координат и высот пунктов опорной геодезической сети.docx") = False Then
                    MsgBox("Шаблон не найден! Макро прервано.")
                    Exit Sub
                End If
            End If
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения")
            End Try
            'открываем word
            Try
                Dim PutchDoc As String = PutchTbl & "Каталог координат и высот пунктов опорной геодезической сети.docx"
                Dim rng As Random = New Random
                Dim Numbrng As Integer = rng.Next
                Dim DestDoc As String = TempDir & "Каталог координат пунктов ГГС_" & Numbrng & ".docx"
                Try
                    If File.Exists(PutchDoc) = True Then
                        File.Copy(PutchDoc, DestDoc)
                    End If
                Catch ex As UnauthorizedAccessException
                Catch ex As ArgumentException
                Catch ex As PathTooLongException
                Catch ex As DirectoryNotFoundException
                Catch ex As FileNotFoundException
                Catch ex As IOException
                Catch ex As NotSupportedException
                End Try
                If File.Exists(DestDoc) = True Then
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(DestDoc)
                Else
                    MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                    WApp.Quit()
                    WApp = Nothing
                    locker.Dispose()
                    Exit Sub
                End If
            Catch
                MsgBox("Не удалось найти выбранную таблицу!!!")
                locker.Dispose()
                WDoc = Nothing
                WApp = Nothing
                Exit Sub
            End Try

            'забираем в массив пункты
            Dim ArrayWord As String(,) = Nothing
            Dim CountArrayWord As Integer = 0
            For Each Point1 As ObjectId In CivilDoc.CogoPoints
                Dim PointCOGO1 As CogoPoint = Point1.GetObject(OpenMode.ForRead)
                Dim boolFindTbl As Boolean = FuncPropertySet.FuncFindPSetTable(PointCOGO1, "GEODETIC_network")
                If boolFindTbl = True Then
                    Dim NumbCOGOPoint As Integer = PointCOGO1.PointNumber
                    ReDim Preserve ArrayWord(5, CountArrayWord)
                    'считываем расширенные данные
                    Dim ArrayData As String(,) = Nothing
                    Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(PointCOGO1, ArrayData, "GEODETIC_network")
                    If IsArray(ArrayData) = True Then
                        Dim namePoint As String = ""
                        Dim klass As String = ""
                        Dim typeCentre As String = ""
                        Dim numbCatalog As String = "-"
                        Dim StatCentre As String = "-"
                        Dim LocationDesk As String = "-"
                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                            Dim NameField As String = ArrayData(1, i)
                            Dim val As String = ArrayData(3, i)
                            If NameField Like "NUMB_CATALOG" Then
                                'номер по каталогу
                                If val.Trim.Length > 0 Then
                                    numbCatalog = val
                                End If
                            ElseIf NameField Like "NAME" Then
                                'наименование пункта
                                If val.Trim.Length > 0 Then
                                    namePoint = val
                                End If
                                If namePoint.Trim.Length = 0 Then
                                    namePoint = PointCOGO1.PointName
                                End If
                                If IsNothing(namePoint) = True Or namePoint.Trim.Length = 0 Then
                                    namePoint = PointCOGO1.PointNumber
                                End If
                            ElseIf NameField Like "CLASS_NETWORK" Then
                                'класс сети
                                klass = val
                                If IsNumeric(klass) Then
                                    klass = klass & " класс"
                                End If
                            ElseIf NameField Like "TYPE_CENTRE" Then
                                'тип центра
                                If val.Trim.Length > 0 Then
                                    typeCentre = val
                                End If
                            ElseIf NameField Like "LOCATION_DESK" Then
                                'описание местоположения
                                If val.Trim.Length > 0 Then
                                    LocationDesk = val
                                End If
                            End If
                        Next i
                        Dim NameRez As String = ""
                        If Not (namePoint Like "-") Then
                            NameRez = namePoint
                        End If
                        If Not (klass Like "") Then
                            If NameRez.Trim.Length > 0 Then
                                NameRez = NameRez & vbLf & klass
                            Else
                                NameRez = klass
                            End If
                        End If
                        If Not (typeCentre Like "") Then
                            If NameRez.Trim.Length > 0 Then
                                NameRez = NameRez & vbLf & typeCentre
                            Else
                                NameRez = typeCentre
                            End If
                        End If
                        ArrayWord(0, CountArrayWord) = numbCatalog
                        ArrayWord(1, CountArrayWord) = NameRez
                        ArrayWord(2, CountArrayWord) = PointCOGO1.Northing
                        ArrayWord(3, CountArrayWord) = PointCOGO1.Easting
                        ArrayWord(4, CountArrayWord) = PointCOGO1.Elevation
                        ArrayWord(5, CountArrayWord) = LocationDesk
                    End If
                    CountArrayWord += 1
                End If
            Next
            'перебираем блоки
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'находим габаритный контейнер исходной полилиниии
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                If TypeOf ent Is BlockReference Then 'если объект имеет тип атрибут
                    Dim blk As BlockReference = ent
                    Dim boolFindTbl As Boolean = FuncPropertySet.FuncFindPSetTable(blk, "GEODETIC_network")
                    If boolFindTbl = True Then
                        ReDim Preserve ArrayWord(5, CountArrayWord)
                        'считываем расширенные данные
                        Dim ArrayData As String(,) = Nothing
                        Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(blk, ArrayData, "GEODETIC_network")
                        If IsArray(ArrayData) = True Then
                            Dim namePoint As String = ""
                            Dim klass As String = ""
                            Dim typeCentre As String = ""
                            Dim numbCatalog As String = "-"
                            Dim StatCentre As String = "-"
                            Dim LocationDesk As String = "-"
                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                Dim NameField As String = ArrayData(1, i)
                                Dim val As String = ArrayData(2, i)
                                If NameField Like "NUMB_CATALOG" Then
                                    'номер по каталогу
                                    If val.Trim.Length > 0 Then
                                        numbCatalog = val
                                    End If
                                ElseIf NameField Like "NAME" Then
                                    'наименование пункта
                                    If val.Trim.Length > 0 Then
                                        namePoint = val
                                    End If
                                    If namePoint.Trim.Length = 0 Then
                                        namePoint = AcadFunction.FuncReadValueAttr(blk, "NAME")
                                    End If
                                    If IsNothing(namePoint) = True Or namePoint.Trim.Length = 0 Then
                                        namePoint = blk.Name
                                    End If
                                ElseIf NameField Like "CLASS_NETWORK" Then
                                    'класс сети
                                    klass = val
                                    If IsNumeric(klass) Then
                                        klass = klass & " класс"
                                    End If
                                ElseIf NameField Like "TYPE_CENTRE" Then
                                    'тип центра
                                    If val.Trim.Length > 0 Then
                                        typeCentre = val
                                    End If
                                ElseIf NameField Like "LOCATION_DESK" Then
                                    'описание местоположения
                                    If val.Trim.Length > 0 Then
                                        LocationDesk = val
                                    End If
                                End If
                            Next i
                            Dim NameRez As String = ""
                            If Not (namePoint Like "-") Then
                                NameRez = namePoint
                            End If
                            If Not (klass Like "") Then
                                If NameRez.Trim.Length > 0 Then
                                    NameRez = NameRez & vbLf & klass
                                Else
                                    NameRez = klass
                                End If
                            End If
                            If Not (typeCentre Like "") Then
                                If NameRez.Trim.Length > 0 Then
                                    NameRez = NameRez & vbLf & typeCentre
                                Else
                                    NameRez = typeCentre
                                End If
                            End If
                            ArrayWord(0, CountArrayWord) = numbCatalog
                            ArrayWord(1, CountArrayWord) = NameRez
                            ArrayWord(2, CountArrayWord) = blk.Position.Y
                            ArrayWord(3, CountArrayWord) = blk.Position.X
                            ArrayWord(4, CountArrayWord) = blk.Position.Z
                            ArrayWord(5, CountArrayWord) = LocationDesk
                        End If
                        CountArrayWord += 1
                    End If
                End If
            Next
            '===========================================================================================
            'заполняем таблицу
            If IsArray(ArrayWord) = True Then
                For i As Integer = 0 To ArrayWord.GetUpperBound(1)
                    For Each tabl As Word.Table In WDoc.Tables
                        Dim countTableword As Integer = tabl.Rows.Count
                        tabl.Rows.Add()
                        tabl.Cell(countTableword, 1).Range.Text = ArrayWord(0, i) 'номер по каталогу
                        tabl.Cell(countTableword, 2).Range.Text = ArrayWord(1, i) 'наименование
                        tabl.Cell(countTableword, 3).Range.Text = ArrayWord(2, i) 'X
                        tabl.Cell(countTableword, 4).Range.Text = ArrayWord(3, i) 'Y
                        tabl.Cell(countTableword, 5).Range.Text = ArrayWord(4, i) 'Z
                        tabl.Cell(countTableword, 6).Range.Text = ArrayWord(5, i) 'Описание
                        Exit For
                    Next
                Next i
            End If
            WApp.Visible = True
        End Using
        locker.Dispose()
        WDoc = Nothing
        WApp = Nothing
    End Sub
    '==================================================================================================
    'Каталог координат и высот пунктов съемочной геодезической сети
    <CommandMethod("CatalogCoordSurveyPoints")>
    Public Sub CatalogCoordSurveyPoints()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'ищем пути для таблиц
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Каталог координат и высот точек геодезической съемочной сети.docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            Else
                If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Каталог координат и высот точек геодезической съемочной сети.docx") = False Then
                    MsgBox("Шаблон не найден! Макро прервано.")
                    Exit Sub
                End If
            End If
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения")
            End Try
            'открываем word
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем документ Word
            Try
                Dim PutchDoc As String = PutchTbl & "Каталог координат и высот точек геодезической съемочной сети.docx"
                Dim rng As Random = New Random
                Dim Numbrng As Integer = rng.Next
                Dim DestDoc As String = TempDir & "Каталог координат пунктов съемочной сети_" & Numbrng & ".docx"
                Try
                    If File.Exists(PutchDoc) = True Then
                        File.Copy(PutchDoc, DestDoc)
                    End If
                Catch ex As UnauthorizedAccessException
                Catch ex As ArgumentException
                Catch ex As PathTooLongException
                Catch ex As DirectoryNotFoundException
                Catch ex As FileNotFoundException
                Catch ex As IOException
                Catch ex As NotSupportedException
                End Try
                If File.Exists(DestDoc) = True Then
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(DestDoc)
                Else
                    MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                    WApp.Quit()
                    WApp = Nothing
                    locker.Dispose()
                    Exit Sub
                End If
            Catch
                MsgBox("Не удалось найти выбранную таблицу!!!")
                locker.Dispose()
                WDoc = Nothing
                WApp = Nothing
                Exit Sub
            End Try
            'забираем в массив точки
            Dim ArrayWord As String(,) = Nothing
            Dim CountArrayWord As Integer = 0
            For Each Point1 As ObjectId In CivilDoc.CogoPoints
                Dim PointCOGO1 As CogoPoint = Point1.GetObject(OpenMode.ForRead)
                'считываем расширенные данные
                Dim boolFindTbl As Boolean = FuncPropertySet.FuncFindPSetTable(PointCOGO1, "GEODETIC_temporary")
                If boolFindTbl = True Then
                    Dim NumbCOGOPoint As Integer = PointCOGO1.PointNumber
                    ReDim Preserve ArrayWord(5, CountArrayWord)
                    Dim ArrayData As String(,) = Nothing
                    Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(PointCOGO1, ArrayData, "GEODETIC_temporary")
                    If IsArray(ArrayData) = True Then
                        Dim namePoint As String = ""
                        Dim klass As String = "-"
                        Dim typeCentre As String = "-"
                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                            Dim NameField As String = ArrayData(1, i)
                            Dim val As String = ArrayData(3, i)
                            If NameField Like "NAME" Then
                                If val.Trim.Length = 0 Then
                                    namePoint = PointCOGO1.PointName
                                End If
                                If IsNothing(namePoint) = True Or namePoint.Trim.Length = 0 Then
                                    namePoint = PointCOGO1.PointNumber
                                End If
                            ElseIf NameField Like "CLASS_NETWORK" Then
                                If val.Trim.Length > 0 Then
                                    klass = val 'класс сети
                                End If
                            ElseIf NameField Like "TYPE_CENTRE" Then
                                If val.Trim.Length > 0 Then
                                    typeCentre = val 'тип центра
                                End If
                            End If
                        Next i
                        ArrayWord(0, CountArrayWord) = namePoint
                        ArrayWord(1, CountArrayWord) = klass
                        ArrayWord(5, CountArrayWord) = typeCentre
                    End If
                    ArrayWord(2, CountArrayWord) = Math.Round(PointCOGO1.Northing, 3)
                    ArrayWord(3, CountArrayWord) = Math.Round(PointCOGO1.Easting, 3)
                    ArrayWord(4, CountArrayWord) = Math.Round(PointCOGO1.Elevation, 3)
                End If
            Next
            'перебираем блоки
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'находим габаритный контейнер исходной полилиниии
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                If TypeOf ent Is BlockReference Then 'если объект имеет тип атрибут
                    Dim blk As BlockReference = ent
                    Dim boolFindTbl As Boolean = FuncPropertySet.FuncFindPSetTable(blk, "GEODETIC_temporary")
                    If boolFindTbl = True Then
                        ReDim Preserve ArrayWord(5, CountArrayWord)
                        Dim ArrayData As String(,) = Nothing
                        Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(blk, ArrayData, "GEODETIC_temporary")
                        If IsArray(ArrayData) = True Then
                            Dim namePoint As String = ""
                            Dim klass As String = "-"
                            Dim typeCentre As String = "-"
                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                Dim NameField As String = ArrayData(1, i)
                                Dim val As String = ArrayData(2, i)
                                If NameField Like "NAME" Then
                                    'наименование пункта
                                    If val.Trim.Length > 0 Then
                                        namePoint = val
                                    End If
                                    If namePoint.Trim.Length = 0 Then
                                        namePoint = AcadFunction.FuncReadValueAttr(blk, "NAME")
                                    End If
                                    If namePoint.Trim.Length = 0 Then
                                        namePoint = AcadFunction.FuncReadValueAttr(blk, "NUMBER")
                                    End If
                                ElseIf NameField Like "CLASS_NETWORK" Then
                                    If val.Trim.Length > 0 Then
                                        klass = val 'класс сети
                                    End If
                                ElseIf NameField Like "TYPE_CENTRE" Then
                                    If val.Trim.Length > 0 Then
                                        typeCentre = val 'тип центра
                                    End If
                                End If
                            Next i
                            ArrayWord(0, CountArrayWord) = namePoint
                            ArrayWord(1, CountArrayWord) = klass
                            ArrayWord(5, CountArrayWord) = typeCentre
                        End If
                        ArrayWord(2, CountArrayWord) = Math.Round(blk.Position.Y, 3)
                        ArrayWord(3, CountArrayWord) = Math.Round(blk.Position.X, 3)
                        ArrayWord(4, CountArrayWord) = Math.Round(blk.Position.Z, 3)
                    End If
                End If
            Next
            '===========================================================================================
            'заполняем таблицу
            If IsArray(ArrayWord) = True Then
                For i As Integer = 0 To ArrayWord.GetUpperBound(1)
                    For Each tabl As Word.Table In WDoc.Tables
                        Dim countTableword As Integer = tabl.Rows.Count
                        tabl.Rows.Add()
                        tabl.Cell(countTableword, 1).Range.Text = ArrayWord(0, i) 'номер по каталогу
                        tabl.Cell(countTableword, 2).Range.Text = ArrayWord(1, i) 'X
                        tabl.Cell(countTableword, 3).Range.Text = ArrayWord(2, i) 'Y
                        tabl.Cell(countTableword, 4).Range.Text = ArrayWord(3, i) 'Z
                        tabl.Cell(countTableword, 5).Range.Text = ArrayWord(4, i) 'Описание
                        Exit For
                    Next
                Next i
            End If
            WApp.Visible = True
        End Using
        locker.Dispose()
        WDoc = Nothing
        WApp = Nothing
    End Sub
    '==================================================================================================
    'Каталог координат и высот пунктов ГРО
    <CommandMethod("CatalogCoordGROPoints")>
    Public Sub CatalogCoordGROPoints()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'ищем пути для таблиц
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Каталог координат и высот пунктов ГРО.docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            Else
                If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Каталог координат и высот пунктов ГРО.docx") = False Then
                    MsgBox("Шаблон не найден! Макро прервано.")
                    Exit Sub
                End If
            End If
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения")
            End Try
            'находим ось существующей трассы
            Dim AlignProject As Alignment = Nothing
            Dim AlignEX As Alignment = Nothing
            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
                    AlignEX = TempAlign
                End If
            Next
            If IsNothing(AlignEX) = True Then
                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги: ")
                If ppr1.Status = PromptStatus.OK Then
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                        AlignEX = ent1
                    End If
                End If
            End If
            'выбираем ось проектной трассы
            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги: ")
            If ppr2.Status = PromptStatus.OK Then
                Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
                If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
                    AlignProject = ent2
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем документ Word
            'выбираем документ Word
            Try
                Dim PutchDoc As String = PutchTbl & "Каталог координат и высот пунктов ГРО.docx"
                Dim rng As Random = New Random
                Dim Numbrng As Integer = rng.Next
                Dim DestDoc As String = TempDir & "Каталог координат и высот пунктов ГРО_" & Numbrng & ".docx"
                Try
                    If File.Exists(PutchDoc) = True Then
                        File.Copy(PutchDoc, DestDoc)
                    End If
                Catch ex As UnauthorizedAccessException
                Catch ex As ArgumentException
                Catch ex As PathTooLongException
                Catch ex As DirectoryNotFoundException
                Catch ex As FileNotFoundException
                Catch ex As IOException
                Catch ex As NotSupportedException
                End Try
                If File.Exists(DestDoc) = True Then
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(DestDoc)
                Else
                    MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                    WApp.Quit()
                    WApp = Nothing
                    locker.Dispose()
                    Exit Sub
                End If
            Catch
                MsgBox("Не удалось найти выбранную таблицу!!!")
                locker.Dispose()
                WDoc = Nothing
                WApp = Nothing
                Exit Sub
            End Try

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'забираем точки в массив
            Dim ArrayPointWord As String(,) = Nothing
            Dim CountArrayPointWord As Integer = 0
            For Each Point1 As ObjectId In CivilDoc.CogoPoints
                Dim NamePoint As String = "" 'название пункта
                Dim PointCOGO1 As CogoPoint = Point1.GetObject(OpenMode.ForRead)
                If Not (PointCOGO1.RawDescription Like "PIKET") Then
                    Dim NumbCOGOPoint As Integer = PointCOGO1.PointNumber
                    'считываем расширенные данные
                    Dim ArrayData As String(,) = Nothing
                    Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(PointCOGO1, ArrayData, "GEODETIC_point")
                    If IsArray(ArrayData) = True Then
                        '0-км для сортировки, 1-КМ+, смещение, 2-ПК+, смещение, 3-имя пункта, 4-X, 5-Y, 6-H, 7-описание местоположения, 8-Класс плановой сети, 9-Класс высотной сети, 10 - Примечание
                        ReDim Preserve ArrayPointWord(10, CountArrayPointWord)
                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                            Dim NameField As String = ArrayData(1, i)
                            Dim val As String = ArrayData(2, i)
                            If IsNothing(val) = True Then val = ""
                            If NameField Like "NAME" Then
                                If val.Trim.Length > 0 Then
                                    NamePoint = val 'обозначение пункта
                                End If
                                If NamePoint.Trim.Length = 0 Then
                                    NamePoint = PointCOGO1.PointName
                                End If
                                If IsNothing(NamePoint) = True Or NamePoint.Trim.Length = 0 Then
                                    NamePoint = PointCOGO1.PointNumber
                                End If
                                ArrayPointWord(3, CountArrayPointWord) = NamePoint
                            ElseIf NameField Like "CLASS_NETWORK_XY" Then
                                Dim KlassXY As String = "-"
                                If val.Trim.Length > 0 Then
                                    KlassXY = val
                                End If
                                ArrayPointWord(8, CountArrayPointWord) = KlassXY
                            ElseIf NameField Like "CLASS_NETWORK_H" Then
                                Dim KlassH As String = "-"
                                If val.Trim.Length > 0 Then
                                    KlassH = val
                                End If
                                ArrayPointWord(9, CountArrayPointWord) = KlassH
                            ElseIf NameField Like "LOCATION_DESK" Then
                                Dim LocationDesk As String = "-"
                                If val.Trim.Length > 0 Then
                                    LocationDesk = val
                                End If
                                ArrayPointWord(7, CountArrayPointWord) = LocationDesk
                            ElseIf NameField Like "NOTE" Then
                                Dim Note As String = "-"
                                If val.Trim.Length > 0 Then
                                    Note = val
                                End If
                                ArrayPointWord(10, CountArrayPointWord) = Note
                            End If
                        Next i
                        '=================================================================================
                        'записываем координаты
                        ArrayPointWord(4, CountArrayPointWord) = Math.Round(PointCOGO1.Northing, 3)
                        ArrayPointWord(5, CountArrayPointWord) = Math.Round(PointCOGO1.Easting, 3)
                        ArrayPointWord(6, CountArrayPointWord) = Math.Round(PointCOGO1.Elevation, 3)
                        '=================================================================================
                        'определяем километр
                        'записываем существующий километраж
                        If IsNothing(AlignEX) = False Then
                            Dim PK As Double = -1
                            Dim off As Double = -1
                            Dim pt3d As Point3d = PointCOGO1.Location
                            Dim boolPkOpred As Boolean = False
                            'определяем пикетаж
                            Try
                                If IsNothing(AlignEX) = False Then
                                    If AlignEX.Id.IsNull = False Then
                                        AlignEX.StationOffset(pt3d.X, pt3d.Y, PK, off)
                                    End If
                                End If
                            Catch ex As ArgumentOutOfRangeException
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                            'если точка в пределах трассы определяем ее пикет
                            If off <> -1 Then
                                Dim km As String = FuncFormatZN.FuncFormatKM(PK)
                                Dim offStr As String = ""
                                If off < 0 Then
                                    offStr = "Лево " & Math.Round(off, 2)
                                ElseIf off > 0 Then
                                    offStr = "Право " & Math.Round(off, 2)
                                End If
                                ArrayPointWord(1, CountArrayPointWord) = km & vbLf & offStr
                                ArrayPointWord(0, CountArrayPointWord) = PK
                            Else
                                ArrayPointWord(0, CountArrayPointWord) = PK
                                ArrayPointWord(1, CountArrayPointWord) = "-"
                            End If
                        End If
                        '=====================================================================================================================
                        'записываем пикетаж проектный
                        If IsNothing(AlignProject) = False Then
                            Dim PK As Double = -1
                            Dim off As Double = -1
                            Dim pt3d As Point3d = PointCOGO1.Location
                            Dim boolPkOpred As Boolean = False
                            'если точка в пределах трассы определяем ее пикет
                            Try
                                If IsNothing(AlignProject) = False Then
                                    If AlignProject.Id.IsNull = False Then
                                        AlignProject.StationOffset(pt3d.X, pt3d.Y, PK, off)
                                    End If
                                End If
                            Catch ex As ArgumentOutOfRangeException
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                            If off <> -1 Then
                                Try
                                    Dim pkStrt As String = AlignProject.GetStationStringWithEquations(PK)
                                    Dim offStr As String = ""
                                    If off < 0 Then
                                        offStr = "Лево " & Math.Round(off, 2)
                                    ElseIf off > 0 Then
                                        offStr = "Право " & Math.Round(off, 2)
                                    End If
                                    ArrayPointWord(2, CountArrayPointWord) = pkStrt & vbLf & offStr
                                    If IsNothing(AlignEX) = True Then
                                        ArrayPointWord(0, CountArrayPointWord) = PK
                                    End If
                                Catch ex As System.ArgumentException
                                    Dim offStr As String = ""
                                    If off < 0 Then
                                        offStr = "Лево " & Math.Round(off, 2)
                                    ElseIf off > 0 Then
                                        offStr = "Право " & Math.Round(off, 2)
                                    End If
                                    ArrayPointWord(2, CountArrayPointWord) = PK & vbLf & offStr
                                    If IsNothing(AlignEX) = True Then
                                        ArrayPointWord(0, CountArrayPointWord) = PK
                                    End If
                                End Try
                            Else
                                ArrayPointWord(2, CountArrayPointWord) = "-"
                                If IsNothing(AlignEX) = True Then
                                    ArrayPointWord(0, CountArrayPointWord) = PK
                                End If
                            End If
                        End If
                        CountArrayPointWord += 1
                    End If
                End If
            Next
            '================================================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем блоки автокад
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'находим габаритный контейнер исходной полилиниии
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                If TypeOf ent Is BlockReference Then 'если объект имеет тип атрибут
                    Dim blk As BlockReference = ent
                    Dim NamePoint As String = "" 'название пункта
                    Dim ArrayTBLPs As String() = Nothing
                    Dim ArrayTBLXML As String(,) = Nothing
                    ReDim Preserve ArrayTBLXML(1, 0)
                    ArrayTBLXML(0, 0) = "GEODETIC_point"
                    ArrayTBLXML(1, 0) = ""
                    Dim boolFindTblXML As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, ArrayTBLPs, ArrayTBLXML)
                    If IsArray(ArrayTBLPs) = True Then
                        'считываем расширенные данные
                        Dim ArrayData As String(,) = Nothing
                        Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(ent, ArrayData, "GEODETIC_point")
                        If IsArray(ArrayData) = True Then
                            '0-км для сортировки, 1-КМ+, смещение, 2-ПК+, смещение, 3-имя пункта, 4-X, 5-Y, 6-H, 7-описание местоположения, 8-Класс плановой сети, 9-Класс высотной сети, 10 - Примечание
                            ReDim Preserve ArrayPointWord(10, CountArrayPointWord)
                            NamePoint = AcadFunction.FuncReadValueAttr(blk, "NAME")
                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                Dim NameField As String = ArrayData(1, i)
                                Dim val As String = ArrayData(2, i)
                                If IsNothing(val) = True Then val = ""
                                If NameField Like "NAME" Then
                                    'считываем атрибут
                                    If NamePoint.Trim.Length = 0 Or NamePoint Like "б/н" Then
                                        If val.Trim.Length > 0 Then
                                            NamePoint = val 'обозначение пункта
                                        End If
                                    End If
                                    If NamePoint.Trim.Length = 0 Then
                                        NamePoint = AcadFunction.FuncReadValueAttr(blk, "NUMBER")
                                    End If
                                    ArrayPointWord(3, CountArrayPointWord) = NamePoint
                                ElseIf NameField Like "CLASS_NETWORK_XY" Then
                                    Dim KlassXY As String = "-"
                                    If val.Trim.Length > 0 Then
                                        KlassXY = val
                                    End If
                                    ArrayPointWord(8, CountArrayPointWord) = KlassXY
                                ElseIf NameField Like "CLASS_NETWORK_H" Then
                                    Dim KlassH As String = "-"
                                    If val.Trim.Length > 0 Then
                                        KlassH = val
                                    End If
                                    ArrayPointWord(9, CountArrayPointWord) = KlassH
                                ElseIf NameField Like "LOCATION_DESK" Then
                                    Dim LocationDesk As String = "-"
                                    If val.Trim.Length > 0 Then
                                        LocationDesk = val
                                    End If
                                    ArrayPointWord(7, CountArrayPointWord) = LocationDesk
                                ElseIf NameField Like "NOTE" Then
                                    Dim Note As String = "-"
                                    If val.Trim.Length > 0 Then
                                        Note = val
                                    End If
                                    ArrayPointWord(10, CountArrayPointWord) = Note
                                End If
                            Next i
                            '=================================================================================
                            'записываем координаты
                            ArrayPointWord(4, CountArrayPointWord) = Math.Round(blk.Position.Y, 3)
                            ArrayPointWord(5, CountArrayPointWord) = Math.Round(blk.Position.X, 3)
                            ArrayPointWord(6, CountArrayPointWord) = Math.Round(blk.Position.Z, 3)
                            '=================================================================================
                            'определяем километр
                            'записываем существующий километраж
                            If IsNothing(AlignEX) = False Then
                                Dim PK As Double = -1
                                Dim off As Double = -1
                                Dim pt3d As Point3d = blk.Position
                                Try
                                    If IsNothing(AlignEX) = False Then
                                        If AlignEX.Id.IsNull = False Then
                                            AlignEX.StationOffset(pt3d.X, pt3d.Y, PK, off)
                                        End If
                                    End If
                                Catch ex As ArgumentOutOfRangeException
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                If off <> -1 Then
                                    Dim km As String = FuncFormatZN.FuncFormatKM(PK)
                                    Dim offStr As String = ""
                                    If off < 0 Then
                                        offStr = "Лево " & Math.Round(off, 2)
                                    ElseIf off > 0 Then
                                        offStr = "Право " & Math.Round(off, 2)
                                    End If
                                    ArrayPointWord(1, CountArrayPointWord) = km & vbLf & offStr
                                    ArrayPointWord(0, CountArrayPointWord) = PK
                                Else
                                    'точка за пределами трассы
                                    ArrayPointWord(0, CountArrayPointWord) = PK
                                    ArrayPointWord(1, CountArrayPointWord) = "-"
                                End If
                            End If
                            'записываем пикетаж проектный
                            If IsNothing(AlignProject) = False Then
                                Dim PK As Double = -1
                                Dim off As Double = -1
                                Dim pt3d As Point3d = blk.Position
                                Try
                                    If IsNothing(AlignProject) = False Then
                                        If AlignProject.Id.IsNull = False Then
                                            AlignProject.StationOffset(pt3d.X, pt3d.Y, PK, off)
                                        End If
                                    End If
                                Catch ex As ArgumentOutOfRangeException
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                If off <> -1 Then
                                    Try
                                        Dim pkStrt As String = AlignProject.GetStationStringWithEquations(PK)
                                        Dim offStr As String = ""
                                        If off < 0 Then
                                            offStr = "Лево " & Math.Round(off, 2)
                                        ElseIf off > 0 Then
                                            offStr = "Право " & Math.Round(off, 2)
                                        End If
                                        ArrayPointWord(2, CountArrayPointWord) = pkStrt & vbLf & offStr
                                        If IsNothing(AlignEX) = True Then
                                            ArrayPointWord(0, CountArrayPointWord) = PK
                                        End If
                                    Catch ex As System.ArgumentException
                                    End Try
                                Else
                                    'точка за пределами трассы
                                    ArrayPointWord(2, CountArrayPointWord) = "-"
                                    If IsNothing(AlignEX) = True Then
                                        ArrayPointWord(0, CountArrayPointWord) = PK
                                    End If
                                End If
                            End If
                            CountArrayPointWord += 1
                        End If
                    End If
                End If
            Next
            If IsArray(ArrayPointWord) = True Then
                If ArrayPointWord.GetUpperBound(1) > 0 Then
                    Dim boolsort As Boolean = MathFunction.FuncSortStrArray(ArrayPointWord, 0)
                End If
                '=========================================================================================
                'заполняем таблицу
                For i As Integer = 0 To ArrayPointWord.GetUpperBound(1)
                    For Each tabl As Word.Table In WDoc.Tables
                        Dim countTableword As Integer = tabl.Rows.Count
                        tabl.Rows.Add()
                        tabl.Cell(countTableword, 1).Range.Text = ArrayPointWord(3, i)
                        tabl.Cell(countTableword, 2).Range.Text = ArrayPointWord(1, i)
                        tabl.Cell(countTableword, 3).Range.Text = ArrayPointWord(2, i)
                        tabl.Cell(countTableword, 4).Range.Text = String.Format("{0:0.000}", Val(ArrayPointWord(4, i)))
                        tabl.Cell(countTableword, 5).Range.Text = String.Format("{0:0.000}", Val(ArrayPointWord(5, i)))
                        tabl.Cell(countTableword, 6).Range.Text = String.Format("{0:0.000}", Val(ArrayPointWord(6, i)))
                        tabl.Cell(countTableword, 7).Range.Text = ArrayPointWord(8, i)
                        tabl.Cell(countTableword, 8).Range.Text = ArrayPointWord(9, i)
                        tabl.Cell(countTableword, 9).Range.Text = ArrayPointWord(7, i)
                        tabl.Cell(countTableword, 10).Range.Text = ArrayPointWord(10, i)
                        Exit For
                    Next
                Next i
            End If
            WApp.Visible = True
            acTrans.Commit()
        End Using
        locker.Dispose()
        WDoc = Nothing
        WApp = Nothing
    End Sub
    '==================================================================================================
    'Ведомость искусственных сооружений (водопропускных труб)
    <CommandMethod("VedCulvertPipe")>
    Public Sub VedCulvertPipe()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'ищем пути для таблиц
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Дефектная ведомость существующих искусственных сооружений (водопропускные трубы).docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            Else
                If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Дефектная ведомость существующих искусственных сооружений (водопропускные трубы).docx") = False Then
                    MsgBox("Шаблон не найден! Макро прервано.")
                    Exit Sub
                End If
            End If
            'находим ось существующей трассы
            Dim AlignProject As Alignment = Nothing
            Dim AlignEX As Alignment = Nothing
            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
                    AlignEX = TempAlign
                End If
            Next
            If IsNothing(AlignEX) = True Then
                Dim ppr1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите трассу существующей автодороги: ")
                If ppr1.Status = PromptStatus.OK Then
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                        AlignEX = ent1
                    End If
                End If
            End If
            'выбираем ось проектной трассы
            Dim ppr2 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите трассу проектируемой автодороги: ")
            If ppr2.Status = PromptStatus.OK Then
                Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
                If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
                    AlignProject = ent2
                End If
            End If
            '========================================================================================
            'находим поверхность рабочую
            Dim ColSurv As ObjectIdCollection = CivilDoc.GetSurfaceIds
            Dim Surv As Autodesk.Civil.DatabaseServices.Surface = Nothing
            If ColSurv.Count > 1 Then
                Dim ppr3 As PromptEntityResult = ed.GetEntity("Выберите рабочую поверхность: ")
                If ppr3.Status = PromptStatus.OK Then
                    Dim ent3 As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(ppr3.ObjectId, OpenMode.ForRead)
                    If TypeOf ent3 Is Autodesk.Civil.DatabaseServices.Surface Then
                        Surv = ent3
                    End If
                End If
            ElseIf ColSurv.Count = 1 Then
                Dim ent3 As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(ColSurv.Item(0), OpenMode.ForRead)
                Surv = ent3
            End If
            '========================================================================================
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения")
            End Try
            'забираем в массив все Характерные линии со стилем ИИ-Водопропускные трубы
            Dim ArrayDataWord As String(,) = Nothing
            Dim CountArrayDataWord As Integer = 0
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'выбираем объекты
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead, False) 'открываем объект на чтение
                        Dim ArrayData As String(,) = Nothing 'массив с данными
                        Dim LPipe As Double = 0 'длина трубы
                        Dim PTStart As Point3d = Nothing 'точка начала xyz
                        Dim PTEnd As Point3d = Nothing 'точка конца xyz
                        Dim km As Double = -1 'пикет и смещение точки пересечения
                        Dim kmoff As Double = -1
                        Dim pk As Double = -1
                        Dim pkoff As Double = -1
                        Dim NameObj As String = "Водопропускная труба под дорогой"
                        Dim sectionPipe As String = ""
                        Dim widthPipe As Double = 0
                        Dim heightPipe As Double = 0
                        Dim holeCount As Integer = 1
                        Dim Mat As String = "ж.б."
                        Dim SlopeWal As String = ""
                        Dim Top1 As Double = -1
                        Dim Top2 As Double = -1
                        Dim Down1 As Double = -1
                        Dim Down2 As Double = -1
                        Dim diameter As Double = 0
                        Dim boolFindObj As Boolean = False 'флаг того что труба найдена
                        Dim AnglePipe As Double = 0
                        Dim State As String = "Удовлетворительное"
                        Dim boolReadPS As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(ent, ArrayData, "CULVERT_pipe")
                        If IsArray(ArrayData) = True Then
                            'структурная линия
                            If TypeOf ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = ent
                                boolFindObj = True
                                'получаем длину трубы
                                LPipe = FLine.Length2D
                                PTStart = FLine.StartPoint
                                PTEnd = FLine.EndPoint
                            ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
                                Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = ent
                                'считываем PSSET
                                boolFindObj = True
                                'получаем длину трубы
                                LPipe = FLine.Length
                                PTStart = FLine.StartPoint
                                PTEnd = FLine.EndPoint
                            ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                Dim FLine As Autodesk.AutoCAD.DatabaseServices.BlockReference = ent
                                'считываем атрибуты блока
                                Dim ArrayAtribBlk As String(,) = Nothing
                                Dim boolreadattrib As Boolean = AcadFunction.FuncReadAllBlockAttr(FLine, ArrayAtribBlk)
                                If IsArray(ArrayAtribBlk) = True Then
                                    For i As Integer = 0 To ArrayAtribBlk.GetUpperBound(1)
                                        If ArrayAtribBlk(0, i) Like "OD" Then
                                            If ArrayAtribBlk(1, i) Like "CULVERT" Or ArrayAtribBlk(1, i) Like "CULVERT_pipe" Then
                                                boolFindObj = True
                                                'находим доступ в сложному блоку (динамическому)
                                                Dim blockDyn As BlockTableRecord = Nothing
                                                If FLine.IsDynamicBlock = True Then
                                                    blockDyn = acTrans.GetObject(FLine.DynamicBlockTableRecord, OpenMode.ForRead)
                                                    Dim NameTemp As String = blockDyn.Name
                                                    If NameTemp.IndexOf("A-") = 0 Then
                                                        NameObj = Mid(NameTemp, 3)
                                                    Else
                                                        NameObj = NameTemp
                                                    End If
                                                    'начало
                                                    PTStart = FLine.Position
                                                    Dim blkColl As DynamicBlockReferencePropertyCollection = FLine.DynamicBlockReferencePropertyCollection
                                                    If blkColl.Count > 0 Then
                                                        For Each blkref As DynamicBlockReferenceProperty In blkColl
                                                            Dim NameF As String = blkref.PropertyName
                                                            If NameF Like "Длина трубы" Then
                                                                LPipe = Val(blkref.Value)
                                                            End If
                                                            If NameF Like "Ширина трубы (Диаметр)" Then
                                                                diameter = Val(blkref.Value)
                                                            End If
                                                            If NameF Like "Угол поворота трубы" Then
                                                                AnglePipe = Val(blkref.Value)
                                                            End If
                                                        Next
                                                        If LPipe > 0 Then
                                                            Dim x As Double = PTStart.X + LPipe * Math.Cos(AnglePipe)
                                                            Dim y As Double = PTStart.Y + LPipe * Math.Sin(AnglePipe)
                                                            PTEnd = New Point3d(x, y, 0)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                            If TypeOf ent Is Autodesk.Civil.DatabaseServices.Pipe Then
                                Dim FLine As Autodesk.Civil.DatabaseServices.Pipe = ent
                                boolFindObj = True
                                PTStart = FLine.StartPoint
                                PTEnd = FLine.EndPoint
                                diameter = FLine.OuterDiameterOrWidth
                            End If
                        End If
                        '=============================================================================================================================================================
                        '0-наименование, 1-КМ, 2-ПК, 3-Материал, 4-Размер, 5-Длина, 6-Уклон, 7-Отметка входного оголовка, 8-Отметка выходного оголовка, 9-отметка верха входа трубы, 10-Отметка низа входа трубы, 11-Отметка верха выхода трубы, 12-Отметка низа выхода трубы, 13-Оголовок, 14-Откосные стенки, 15-Состояние, 16-Примечание, 17-Наименование
                        If boolFindObj = True Then
                            ReDim Preserve ArrayDataWord(17, CountArrayDataWord)
                            ArrayDataWord(17, CountArrayDataWord) = "Водопропускная труба под дорогой"
                            'ищем КМ
                            Dim acLine As Line = New Line(PTStart, PTEnd)
                            Dim pointIDColl As Point3dCollection = New Point3dCollection
                            Dim planeUsr As Plane = New Plane
                            Dim pointCogoIns As Point3d = New Point3d(-1, -1, -1)
                            If IsNothing(AlignEX) = False Then
                                If AlignEX.Id.IsNull = False Then
                                    Try
                                        AlignEX.IntersectWith(acLine, Intersect.OnBothOperands, New Plane(), pointIDColl, IntPtr.Zero, IntPtr.Zero)
                                    Catch ex As System.Exception
                                    End Try
                                    'объект пересекает трасу
                                    If pointIDColl.Count > 0 Then
                                        'пересечение найдено
                                        For i1 As Integer = 0 To pointIDColl.Count - 1
                                            Dim tempPoint As Point3d = pointIDColl.Item(i1)
                                            Try
                                                AlignEX.StationOffset(tempPoint.X, tempPoint.Y, km, kmoff)
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        Next i1
                                    Else
                                        'труба ось не пересекает, находим ее среднюю точку и вычисляем пикетаж
                                        Dim SRPoint As Point3d = New Point3d((PTStart.X + PTEnd.X) / 2, (PTStart.Y + PTEnd.Y) / 2, 0)
                                        Try
                                            AlignEX.StationOffset(SRPoint.X, SRPoint.Y, km, kmoff)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    End If
                                End If
                                'находим километр
                                If kmoff <> -1 And km <> -1 Then
                                    Dim kmStr As String = FuncFormatZN.FuncFormatKM(km)
                                    Dim offStr As String = ""
                                    If kmoff < 0 Then
                                        offStr = "Лево " & Math.Round(kmoff, 2)
                                    ElseIf kmoff > 0 Then
                                        offStr = "Право " & Math.Round(kmoff, 2)
                                    End If
                                    ArrayDataWord(1, CountArrayDataWord) = kmStr & vbLf & offStr
                                    ArrayDataWord(0, CountArrayDataWord) = km
                                End If
                            End If
                            'ищем проектный пикетаж
                            Dim boolIntersectPipe As Boolean = True
                            pointIDColl = New Point3dCollection
                            If IsNothing(AlignProject) = False Then
                                If AlignProject.Id.IsNull = False Then
                                    Try
                                        AlignProject.IntersectWith(acLine, Intersect.OnBothOperands, New Plane(), pointIDColl, IntPtr.Zero, IntPtr.Zero)
                                    Catch ex As system.Exception
                                    End Try
                                    'объект пересекает трасу
                                    If pointIDColl.Count > 0 Then
                                        'пересечение найдено
                                        For i1 As Integer = 0 To pointIDColl.Count - 1
                                            Dim tempPoint As Point3d = pointIDColl.Item(i1)
                                            pointCogoIns = tempPoint
                                            Try
                                                AlignProject.StationOffset(tempPoint.X, tempPoint.Y, pk, pkoff)
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        Next i1
                                    Else
                                        'труба ось не пересекает, находим ее среднюю точку и вычисляем пикетаж
                                        Dim SRPoint As Point3d = New Point3d((PTStart.X + PTEnd.X) / 2, (PTStart.Y + PTEnd.Y) / 2, 0)
                                        boolIntersectPipe = False
                                        Try
                                            AlignProject.StationOffset(SRPoint.X, SRPoint.Y, pk, pkoff)
                                            pointCogoIns = SRPoint
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    End If
                                    'находим форматированный проектный пикет
                                    If pkoff <> -1 Then
                                        Try
                                            Dim pkStrt As String = AlignProject.GetStationStringWithEquations(pk)
                                            Dim offStr As String = ""
                                            If pkoff < 0 Then
                                                offStr = "Лево " & Math.Round(pkoff, 2)
                                            ElseIf pkoff > 0 Then
                                                offStr = "Право " & Math.Round(pkoff, 2)
                                            End If
                                            ArrayDataWord(2, CountArrayDataWord) = pkStrt & vbLf & offStr
                                            If IsNothing(AlignEX) = True Then
                                                ArrayDataWord(0, CountArrayDataWord) = pk
                                            End If
                                        Catch ex As System.ArgumentException
                                            ArrayDataWord(2, CountArrayDataWord) = pk & vbLf & pkoff
                                            If IsNothing(AlignEX) = True Then
                                                ArrayDataWord(0, CountArrayDataWord) = pk
                                            End If
                                        End Try
                                    End If
                                End If
                            End If
                            '===========================================================================
                            'записываем OD Характеристики
                            If IsArray(ArrayData) = True Then
                                For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                    Dim NameF As String = ArrayData(1, j)
                                    Dim val As String = ArrayData(2, j)
                                    If NameF Like "STATE" Then
                                        State = val
                                    ElseIf NameF Like "NOTE" Then
                                        ArrayDataWord(16, CountArrayDataWord) = val
                                    ElseIf NameF Like "DEPTH_WATER" Then
                                        ArrayDataWord(15, CountArrayDataWord) = val
                                    ElseIf NameF Like "MATERIAL" Then
                                        Mat = val
                                        If Mat Like "Железобетон" Then
                                            Mat = "ж.б. "
                                        End If
                                        ArrayDataWord(3, CountArrayDataWord) = val
                                    ElseIf NameF Like "DIAMETER" Then
                                        If IsNumeric(val) = False Then
                                            diameter = 0
                                        Else
                                            diameter = val
                                        End If
                                    ElseIf NameF Like "SECTION" Then
                                        sectionPipe = val
                                    ElseIf NameF Like "WIDTH" Then
                                        If IsNumeric(val) = False Then
                                            widthPipe = 0
                                        Else
                                            widthPipe = val
                                        End If
                                    ElseIf NameF Like "HEIGHT" Then
                                        If IsNumeric(val) = False Then
                                            heightPipe = 0
                                        Else
                                            heightPipe = val
                                        End If
                                    ElseIf NameF Like "HOLECOUNT" Then
                                        If IsNumeric(val) = False Then
                                            holeCount = 1
                                        Else
                                            holeCount = CInt(val)
                                        End If
                                    ElseIf NameF Like "SLOPWALL" Then
                                        SlopeWal = val
                                    ElseIf NameF Like "LEVEL1_TOP" Then
                                        Dim TempVal As Double = 0
                                        If IsNumeric(val) = False Then
                                            TempVal = 0
                                        Else
                                            TempVal = val
                                        End If
                                        If TempVal = 0 Then
                                            Top1 = -1
                                        Else
                                            Top1 = TempVal
                                        End If
                                    ElseIf NameF Like "LEVEL2_TOP" Then
                                        Dim TempVal As Double = 0
                                        If IsNumeric(val) = False Then
                                            TempVal = 0
                                        Else
                                            TempVal = val
                                        End If
                                        If TempVal = 0 Then
                                            Top2 = -1
                                        Else
                                            Top2 = TempVal
                                        End If
                                    ElseIf NameF Like "LEVEL1_DOWN" Then
                                        Dim TempVal As Double = 0
                                        If IsNumeric(val) = False Then
                                            TempVal = 0
                                        Else
                                            TempVal = val
                                        End If
                                        If TempVal = 0 Then
                                            Down1 = -1
                                        Else
                                            Down1 = TempVal
                                        End If
                                    ElseIf NameF Like "LEVEL2_DOWN" Then
                                        Dim TempVal As Double = 0
                                        If IsNumeric(val) = False Then
                                            TempVal = 0
                                        Else
                                            TempVal = val
                                        End If
                                        If TempVal = 0 Then
                                            Down2 = -1
                                        Else
                                            Down2 = TempVal
                                        End If
                                    End If
                                Next j
                            End If
                            'начинаем разбор найденных свойств
                            If Top1 = -1 And Down1 = -1 Then
                                Top1 = Math.Round(PTStart.Z, 3)
                                If diameter > 0 Then
                                    Down1 = Top1 - diameter
                                Else
                                    Down1 = Top1 - heightPipe
                                End If
                            ElseIf Top1 = -1 And Down1 <> -1 Then
                                If diameter > 0 Then
                                    Top1 = Down1 + diameter
                                Else
                                    Top1 = Down1 + heightPipe
                                End If
                            ElseIf Top1 <> -1 And Down1 = -1 Then
                                If diameter > 0 Then
                                    Down1 = Top1 - diameter
                                Else
                                    Down1 = Top1 - heightPipe
                                End If
                            End If

                            If Top2 = -1 And Down2 = -1 Then
                                Top2 = Math.Round(PTEnd.Z, 3)
                                If diameter > 0 Then
                                    Down2 = Top2 - diameter
                                Else
                                    Down2 = Top2 - heightPipe
                                End If
                            ElseIf Top2 = -1 And Down2 <> -1 Then
                                If diameter > 0 Then
                                    Top2 = Down2 + diameter
                                Else
                                    Top2 = Down2 + heightPipe
                                End If
                            ElseIf Top2 <> -1 And Down2 = -1 Then
                                If diameter > 0 Then
                                    Down2 = Top2 - diameter
                                Else
                                    Down2 = Top2 - heightPipe
                                End If
                            End If
                            'записываем материал и диаметр трубы
                            Dim RezMatDiam As String = ""
                            If Mat.Trim.Length > 0 Then
                                RezMatDiam = Mat
                            End If
                            If RezMatDiam.Trim.Length > 0 Then
                                RezMatDiam = RezMatDiam & vbLf & String.Format("{0:0.0}", diameter)
                            Else
                                RezMatDiam = String.Format("{0:0.0}", diameter)
                            End If
                            ArrayDataWord(4, CountArrayDataWord) = RezMatDiam
                            'записываем отметки в массив
                            ArrayDataWord(9, CountArrayDataWord) = Math.Round(Top1, 2)
                            ArrayDataWord(10, CountArrayDataWord) = Math.Round(Down1, 2)
                            ArrayDataWord(11, CountArrayDataWord) = Math.Round(Top2, 2)
                            ArrayDataWord(12, CountArrayDataWord) = Math.Round(Down2, 2)
                            'записываем отметки оголовка трубы
                            If IsNothing(Surv) = False Then
                                Try
                                    Dim OGVx As Double = Surv.FindElevationAtXY(PTStart.X, PTStart.Y)
                                    ArrayDataWord(7, CountArrayDataWord) = Math.Round(OGVx, 2)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    ArrayDataWord(7, CountArrayDataWord) = "-"
                                End Try
                                Try
                                    Dim OGVx As Double = Surv.FindElevationAtXY(PTEnd.X, PTEnd.Y)
                                    ArrayDataWord(8, CountArrayDataWord) = Math.Round(OGVx, 2)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    ArrayDataWord(8, CountArrayDataWord) = "-"
                                End Try
                            End If
                            'записываем длину трубы
                            ArrayDataWord(5, CountArrayDataWord) = Math.Round(LPipe, 2)
                            'записываем уклон трубы
                            Dim i As Double = (Math.Abs(Top1 - Top2) / LPipe) * 1000
                            ArrayDataWord(6, CountArrayDataWord) = Math.Round(i, 0)
                            'записываем форму трубы
                            Dim Note As String = ""
                            If sectionPipe.Trim.Length > 0 Then
                                Note = "Сечение трубы - " & sectionPipe
                            End If
                            If Note.Trim.Length > 0 Then
                                Note = Note & vbLf & "Число очков - " & holeCount
                            Else
                                Note = "Число очков - " & holeCount
                            End If
                            If State.Trim.Length > 0 Then
                                If Note.Trim.Length > 0 Then
                                    Note = Note & vbLf & "Состояние - " & State
                                Else
                                    Note = "Состояние - " & State
                                End If
                            End If
                            If IsNothing(ArrayDataWord(16, CountArrayDataWord)) = False Then
                                If ArrayDataWord(16, CountArrayDataWord).Trim.Length > 0 Then
                                    If Note.Trim.Length > 0 Then
                                        Note = Note & vbLf & ArrayDataWord(16, CountArrayDataWord)
                                    Else
                                        Note = ArrayDataWord(16, CountArrayDataWord)
                                    End If
                                End If
                            End If
                            ArrayDataWord(16, CountArrayDataWord) = Note
                            CountArrayDataWord += 1
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'если массив существует
            If IsArray(ArrayDataWord) = True Then
                If ArrayDataWord.GetUpperBound(1) > 0 Then
                    Dim boolsort As Boolean = MathFunction.FuncSortStrArray(ArrayDataWord, 0)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем документ Word
                Try
                    Dim PutchDoc As String = PutchTbl & "Дефектная ведомость существующих искусственных сооружений (водопропускные трубы).docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "Дефектная ведомость водопропускных труб_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        locker.Dispose()
                        Exit Sub
                    End If
                Catch
                    MsgBox("Не удалось найти выбранную таблицу!!!")
                    locker.Dispose()
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try
                '=======================================================================================================================
                'заполняем таблицу
                For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
                    For Each tabl As Word.Table In WDoc.Tables
                        Dim countTableword As Integer = tabl.Rows.Count
                        If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
                            Exit For
                        End If
                        tabl.Rows.Add()
                        tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(17, i) 'название объекта
                        tabl.Cell(countTableword, 2).Range.Text = ArrayDataWord(1, i) 'существующий/проектный пикетаж
                        tabl.Cell(countTableword, 3).Range.Text = ArrayDataWord(2, i)
                        tabl.Cell(countTableword, 4).Range.Text = "-" 'наименование водотока
                        tabl.Cell(countTableword, 5).Range.Text = String.Format("{0:0.0}", Val(ArrayDataWord(15, i))) 'глубина подпора воды
                        tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(4, i) 'материал и размеры
                        tabl.Cell(countTableword, 7).Range.Text = String.Format("{0:0.00}", Val(ArrayDataWord(5, i))) 'длина
                        tabl.Cell(countTableword, 8).Range.Text = String.Format("{0:0.0}", Val(ArrayDataWord(6, i))) 'уклон
                        'отметки оголовков
                        Dim og1 As Double = ArrayDataWord(7, i)
                        Dim og2 As Double = ArrayDataWord(8, i)
                        If og1 = 0 And og2 = 0 Then
                            tabl.Cell(countTableword, 9).Range.Text = "-"
                        ElseIf og1 = 0 And og2 <> 0 Then
                            tabl.Cell(countTableword, 9).Range.Text = "нет" & vbLf & String.Format("{0:0.00}", og2)
                        ElseIf og1 <> 0 And og2 = 0 Then
                            tabl.Cell(countTableword, 9).Range.Text = String.Format("{0:0.00}", og1) & vbLf & "нет"
                        Else
                            tabl.Cell(countTableword, 9).Range.Text = String.Format("{0:0.00}", og1) & vbLf & String.Format("{0:0.00}", og2)
                        End If
                        'отметки труб
                        'отметки верха/низа входа трубы
                        tabl.Cell(countTableword, 10).Range.Text = String.Format("{0:0.00}", Val(ArrayDataWord(9, i))) & vbLf & String.Format("{0:0.00}", Val(ArrayDataWord(10, i)))
                        'отметки верха/низа выхода трубы
                        tabl.Cell(countTableword, 11).Range.Text = String.Format("{0:0.00}", Val(ArrayDataWord(11, i))) & vbLf & String.Format("{0:0.00}", Val(ArrayDataWord(12, i)))
                        'примечание
                        tabl.Cell(countTableword, 12).Range.Text = ArrayDataWord(16, i)
                        Exit For
                    Next
                Next i
                WApp.Visible = True
            End If
        End Using
        locker.Dispose()
        WDoc = Nothing
        WApp = Nothing
    End Sub
    '==================================================================================================
    'ведомость инженерных коммуникаций
    <CommandMethod("VedEnginiringСommunications")>
    Public Sub VedEnginiringСommunications()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        'документ word
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        'массив с элевентами
        Dim ArrayDataWord As String(,) = Nothing
        Dim CountArrayDataWord As Integer = 0
        'вспомогательные переменные
        Dim boolRez As Boolean = False
        Dim AngleEX As Double(,) = Nothing
        Dim countAngleEx As Integer = 0
        Dim AngleProject As Double(,) = Nothing
        Dim countAngleProject As Integer = 0
        'ищем пути для таблиц
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
        If File.Exists(parentPatch2 & "\TemplateTables\Survey\Ведомость инженерных коммуникаций_Росавтодор.docx") = True Then
            PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
        Else
            If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Ведомость инженерных коммуникаций_Росавтодор.docx") = False Then
                MsgBox("Шаблон не найден! Макро прервано.")
                Exit Sub
            End If
        End If
        'временный каталог для автосохранения таблицы
        Dim TempDir As String = ""
        Try
            TempDir = System.IO.Path.GetTempPath()
        Catch ex As System.Security.SecurityException
            TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
        End Try
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'находим ось существующей трассы
            Dim AlignProject As Alignment = Nothing
            Dim AlignEX As Alignment = Nothing
            Dim PlineAlignEx As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
            Dim PlineAlignProject As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
            Dim ArrayData As String(,) = Nothing
            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
                    AlignEX = TempAlign
                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
                End If
            Next
            If IsNothing(AlignEX) = True Then
                Dim ppr1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите трассу существующей автодороги:")
                If ppr1.Status = PromptStatus.OK Then
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                        AlignEX = ent1
                        PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
                    End If
                End If
            End If
            'выбираем ось проектной трассы
            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
            If ppr2.Status = PromptStatus.OK Then
                Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
                If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
                    AlignProject = ent2
                    PlineAlignProject = acTrans.GetObject(AlignProject.GetPolyline, OpenMode.ForRead)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'тело программы
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'выбираем объекты
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Try
                            Dim colPointEX As Point3dCollection = New Point3dCollection 'точки пересечения с существующей трассой
                            Dim colPointProject As Point3dCollection = New Point3dCollection 'точки пересечения с проектной трассой
                            'начало и конец километрового положения трубопровода
                            Dim km1 As Double = -1
                            Dim kmoff1 As Double = -1
                            Dim km2 As Double = -1
                            Dim kmoff2 As Double = -1
                            'начало и конец пикетажного положения трубопровода
                            Dim pk1 As Double = -1
                            Dim pkoff1 As Double = -1
                            Dim pk2 As Double = -1
                            Dim pkoff2 As Double = -1
                            'чистим массивы
                            Erase AngleEX 'массив с углами пересечек по сущестыующей трассе
                            Erase AngleProject 'массив с углами пересечек по проектной трассе
                            Erase ArrayData 'масив с расширенными данными
                            countAngleEx = 0
                            countAngleProject = 0
                            boolRez = False
                            Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            'читаем расширенные данные с примитива
                            Dim boolReadPS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(Ent, ArrayData)
                            If IsArray(ArrayData) = True Then
                                boolRez = True
                            End If
                            If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
                                Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
                                Dim PTStart As Point3d = FLine.StartPoint
                                Dim PTEnd As Point3d = FLine.EndPoint
                                'превращаем характерныю линию в полилинию
                                Dim colPt As Point3dCollection = New Point3dCollection
                                colPt = FLine.GetPoints(Autodesk.Civil.FeatureLinePointType.PIPoint)
                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                                If colPt.Count > 1 Then
                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                                    For i As Integer = 0 To colPt.Count - 1
                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
                                    Next
                                End If
                                '1. Ищем пересечение с осью трассы
                                If IsNothing(AlignEX) = False Then
                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
                                    'ещем касательную к полилинии
                                    If colPointEX.Count > 0 Then
                                        For Each pt As Point3d In colPointEX
                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                            Dim dAngle As Double = Math.Abs(a2 - a1)
                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
                                            ReDim Preserve AngleEX(1, countAngleEx)
                                            AngleEX(0, countAngleEx) = dist
                                            AngleEX(1, countAngleEx) = dAngle
                                            countAngleEx += 1
                                        Next
                                    End If
                                    'ищем пикетаж (существующий километр) начало участка
                                    Try
                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(acPoly) = False Then
                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
                                            Try
                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                    'ищем пикетаж (существующий километр) конец участка
                                    Try
                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(acPoly) = False Then
                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.EndPoint, False)
                                            Try
                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'проектный пикетаж (начало участка)
                                '1. Ищем пересечение с осью трассы
                                If IsNothing(AlignProject) = False Then
                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
                                    If colPointProject.Count > 0 Then
                                        For Each pt As Point3d In colPointProject
                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                            Dim dAngle As Double = Math.Abs(a2 - a1)
                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
                                            ReDim Preserve AngleProject(1, countAngleProject)
                                            AngleProject(0, countAngleProject) = dist
                                            AngleProject(1, countAngleProject) = dAngle
                                            countAngleProject += 1
                                        Next
                                    End If
                                    Try
                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(acPoly) = False Then
                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
                                            Try
                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                    'проектный пикетаж (конец участка)
                                    Try
                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(acPoly) = False Then
                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.EndPoint, False)
                                            Try
                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                End If
                                '2d полилинии
                            ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                                Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
                                Dim NameTypeLine As String = FLine.Linetype
                                If IsArray(ArrayData) = True Then
                                    Dim PTStart As Point3d = FLine.StartPoint
                                    Dim PTEnd As Point3d = FLine.EndPoint
                                    boolRez = True
                                    'ищем пикетаж (существующий километр) начало участка
                                    If IsNothing(AlignEX) = False Then
                                        AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
                                        If colPointEX.Count > 0 Then
                                            For Each pt As Point3d In colPointEX
                                                Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
                                                Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                                Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
                                                Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                                Dim dAngle As Double = Math.Abs(a2 - a1)
                                                Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
                                                ReDim Preserve AngleEX(1, countAngleEx)
                                                AngleEX(0, countAngleEx) = dist
                                                AngleEX(1, countAngleEx) = dAngle
                                                countAngleEx += 1
                                            Next
                                        End If
                                        Try
                                            AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            'конец линии за пределами трассы
                                            'ищем пересечение с началом трассы
                                            If IsNothing(FLine) = False Then
                                                Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
                                                Try
                                                    AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            End If
                                        End Try
                                        'ищем пикетаж (существующий километр) конец участка
                                        Try
                                            AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            'конец линии за пределами трассы
                                            'ищем пересечение с началом трассы
                                            If IsNothing(FLine) = False Then
                                                Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
                                                Try
                                                    AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            End If
                                        End Try
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'проектный пикетаж (начало участка)
                                    If IsNothing(AlignProject) = False Then
                                        AlignProject.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
                                        If colPointProject.Count > 0 Then
                                            For Each pt As Point3d In colPointProject
                                                Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
                                                Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                                Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
                                                Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                                Dim dAngle As Double = Math.Abs(a2 - a1)
                                                Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
                                                ReDim Preserve AngleProject(1, countAngleProject)
                                                AngleProject(0, countAngleProject) = dist
                                                AngleProject(1, countAngleProject) = dAngle
                                                countAngleProject += 1
                                            Next
                                        End If
                                        Try
                                            AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            'конец линии за пределами трассы
                                            'ищем пересечение с началом трассы
                                            If IsNothing(FLine) = False Then
                                                Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
                                                Try
                                                    AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            End If
                                        End Try
                                        'проектный пикетаж (конец участка)
                                        Try
                                            AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            'конец линии за пределами трассы
                                            'ищем пересечение с началом трассы
                                            If IsNothing(FLine) = False Then
                                                Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
                                                Try
                                                    AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
                                                Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            End If
                                        End Try
                                    End If
                                End If
                                'полилиния
                            ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
                                Dim NameTypeLine As String = FLine.Linetype
                                Dim PTStart As Point3d = FLine.StartPoint
                                Dim PTEnd As Point3d = FLine.EndPoint
                                boolRez = True
                                'ищем пикетаж (существующий километр) начало участка
                                If IsNothing(AlignEX) = False Then
                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
                                    If colPointEX.Count > 0 Then
                                        For Each pt As Point3d In colPointEX
                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                            Dim dAngle As Double = Math.Abs(a2 - a1)
                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
                                            ReDim Preserve AngleEX(1, countAngleEx)
                                            AngleEX(0, countAngleEx) = dist
                                            AngleEX(1, countAngleEx) = dAngle
                                            countAngleEx += 1
                                        Next
                                    End If
                                    Try
                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(FLine) = False Then
                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
                                            Try
                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                    'ищем пикетаж (существующий километр) конец участка
                                    Try
                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(FLine) = False Then
                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
                                            Try
                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'проектный пикетаж (начало участка)
                                If IsNothing(AlignProject) = False Then
                                    AlignProject.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
                                    If colPointProject.Count > 0 Then
                                        For Each pt As Point3d In colPointProject
                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
                                            Dim dAngle As Double = Math.Abs(a2 - a1)
                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
                                            ReDim Preserve AngleProject(1, countAngleProject)
                                            AngleProject(0, countAngleProject) = dist
                                            AngleProject(1, countAngleProject) = dAngle
                                            countAngleProject += 1
                                        Next
                                    End If
                                    Try
                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(FLine) = False Then
                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
                                            Try
                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                    'проектный пикетаж (конец участка)
                                    Try
                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        'конец линии за пределами трассы
                                        'ищем пересечение с началом трассы
                                        If IsNothing(FLine) = False Then
                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
                                            Try
                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        End If
                                    End Try
                                End If
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-КМ начало, 6,7-КМ конец, 8,9-ПК начало, 10,11-ПК конец, 12-Глубина закладки, 13,14-Угол пересечения, 15-материал, 16-Диаметр, 17-Владелец, 18-Адрес, 19-Примечание
                            '===========================================================================
                            'записываем OD Характеристики
                            Dim GlobalName As String = "" 'наименование сети в общем виде например Водопровод
                            Dim NameNet As String = "" 'название линии
                            Dim SizeNet As String = "" 'размер
                            Dim MatNet As String = "" 'марка провода
                            Dim HeightNet As Double = 0 'средняя глубина заложения (высота над поверхностью земли) до верхней образующей
                            Dim CountPipes As String = "1" 'число каналов, кабелей
                            Dim Owner As String = "" 'владелец
                            Dim Adress As String = "" 'адрес
                            Dim Power As String = ""
                            If boolRez = True Then
                                If IsArray(ArrayData) = True Then
                                    For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Dim NameF As String = ArrayData(1, j)
                                        Dim val As String = ArrayData(2, j)
                                        If IsNothing(val) = True Then
                                            val = ""
                                        End If
                                        'общее наименование
                                        If NameF Like "LABEL" Then
                                            If val.Trim.Length > 0 Then
                                                GlobalName = val
                                            End If
                                            'имя 1 сети
                                        ElseIf NameF Like "NAME1" Then
                                            If val.Trim.Length > 0 Then
                                                NameNet = NameNet & " " & val
                                            End If
                                            'имя 2 сети
                                        ElseIf NameF Like "NAME2" Then
                                            If val.Trim.Length > 0 Then
                                                NameNet = NameNet & " " & val
                                            End If
                                            'марка провода
                                        ElseIf NameF Like "MARK_WIRES" Then
                                            If val.Trim.Length > 0 Then
                                                MatNet = val
                                            End If
                                            'материал трубы
                                        ElseIf NameF Like "MATERIAL" Then
                                            If val.Trim.Length > 0 Then
                                                MatNet = val
                                            End If
                                        ElseIf NameF Like "DIAMETER" Then
                                            If val.Trim.Length > 0 Then
                                                SizeNet = val
                                            End If
                                        ElseIf NameF Like "HEIGHT" Then
                                            Try
                                                If IsNumeric(val) = True Then
                                                    HeightNet = CDbl(val)
                                                End If
                                            Catch ex As System.InvalidCastException
                                            End Try
                                        ElseIf NameF Like "DEPTH" Then
                                            Try
                                                If IsNumeric(val) = True Then
                                                    HeightNet = CDbl(val)
                                                End If
                                            Catch ex As System.InvalidCastException
                                            End Try
                                        ElseIf NameF Like "SLAK_WIRES" Then
                                            Try
                                                If IsNumeric(val) = True Then
                                                    HeightNet = CDbl(val)
                                                End If
                                            Catch ex As System.InvalidCastException
                                            End Try
                                        ElseIf NameF Like "WORK_PRESS" Then
                                            If val.Trim.Length > 0 Then
                                                Power = val
                                            End If
                                        ElseIf NameF Like "VOLTAGE" Then
                                            If val.Trim.Length > 0 Then
                                                Power = val
                                            End If
                                        ElseIf NameF Like "COUNT_PIPES" Then
                                            If val.Trim.Length > 0 Then
                                                CountPipes = val
                                            End If
                                        ElseIf NameF Like "NUMB_WIRES" Then
                                            If val.Trim.Length > 0 Then
                                                CountPipes = val
                                            End If
                                        ElseIf NameF Like "OWNER" Then
                                            If val.Trim.Length > 0 Then
                                                Owner = val
                                            End If
                                        ElseIf NameF Like "ADRESS" Then
                                            If val.Trim.Length > 0 Then
                                                Adress = val
                                            End If
                                        End If
                                    Next j
                                Else
                                    GlobalName = Ent.Layer
                                End If
                                '///////////////////////////////////////////////////////////////////////////////////////////////
                                '=========================================================================================
                                'начинаем разбор найденных свойств
                                ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
                                '1. записываем наименование объекта
                                GlobalName = GlobalName.Replace("_", " ")
                                If NameNet.Trim.Length > 0 Then
                                    GlobalName = GlobalName & vbLf & NameNet.Trim
                                End If
                                ArrayDataWord(1, CountArrayDataWord) = GlobalName
                                '2. записываем пересечение КМ
                                If colPointEX.Count > 0 Then
                                    Dim str As String = ""
                                    Dim km As Double = 99999999
                                    For Each pt As Point3d In colPointEX
                                        Dim kmper As Double = -1
                                        Dim offper As Double = -1
                                        Try
                                            AlignEX.StationOffset(pt.X, pt.Y, kmper, offper)
                                            ArrayDataWord(2, CountArrayDataWord) = ArrayDataWord(2, CountArrayDataWord) & " " & FuncFormatZN.FuncFormatKM(kmper)
                                            If kmper < km Then
                                                km = kmper
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    Next
                                    ArrayDataWord(0, CountArrayDataWord) = km
                                Else
                                    Dim ErrDist As Integer = 9999999 + CountArrayDataWord
                                    ArrayDataWord(0, CountArrayDataWord) = ErrDist
                                End If
                                '3. записывваем пересечение с проектной трассой
                                If colPointProject.Count > 0 Then
                                    Dim str As String = ""
                                    Dim pk As Double = 99999999
                                    For Each pt As Point3d In colPointProject
                                        Dim pkper As Double = -1
                                        Dim offper As Double = -1
                                        Try
                                            AlignProject.StationOffset(pt.X, pt.Y, pk, offper)
                                            ArrayDataWord(3, CountArrayDataWord) = ArrayDataWord(3, CountArrayDataWord) & " " & FuncFormatZN.FuncFormatPK(pk)
                                            If pkper < pk Then
                                                pk = pkper
                                            End If
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                    Next
                                    If IsNothing(AlignEX) = True Then
                                        ArrayDataWord(0, CountArrayDataWord) = pk
                                    End If
                                Else
                                    If IsNothing(AlignEX) = True Then
                                        Dim ErrDist As Integer = 9999999 + CountArrayDataWord
                                        ArrayDataWord(0, CountArrayDataWord) = ErrDist
                                    End If
                                End If
                                '3. записываем км c,kb;tybz
                                If km1 <> -1 And km2 <> -1 Then
                                    If km1 > km2 Then
                                        ArrayDataWord(4, CountArrayDataWord) = km2
                                        ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff2, 2)
                                        ArrayDataWord(6, CountArrayDataWord) = km1
                                        ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff1, 2)
                                    Else
                                        ArrayDataWord(4, CountArrayDataWord) = km1
                                        ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff1, 2)
                                        ArrayDataWord(6, CountArrayDataWord) = km2
                                        ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff2, 2)
                                    End If
                                End If
                                If pk1 <> -1 And pk2 <> -1 Then
                                    If pk1 > pk2 Then
                                        ArrayDataWord(8, CountArrayDataWord) = pk2
                                        ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff2, 2)
                                        ArrayDataWord(10, CountArrayDataWord) = pk1
                                        ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff1, 2)
                                    Else
                                        ArrayDataWord(8, CountArrayDataWord) = pk1
                                        ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff1, 2)
                                        ArrayDataWord(10, CountArrayDataWord) = pk2
                                        ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff2, 2)
                                    End If
                                End If
                                'глубина закладки
                                ArrayDataWord(12, CountArrayDataWord) = HeightNet
                                'вычисляем угол
                                If IsArray(AngleEX) = True Then
                                    If AngleEX.GetUpperBound(1) > 0 Then
                                        Dim boolSort As Boolean = MathFunction.FuncSortDblArray(AngleEX, 0)
                                    End If
                                    For i As Integer = 0 To AngleEX.GetUpperBound(1)
                                        Dim angl As Double = Math.Round((AngleEX(1, i) * 180) / Math.PI, 0)
                                        If angl > 180 Then angl = 360 - angl
                                        ArrayDataWord(13, CountArrayDataWord) = ArrayDataWord(13, CountArrayDataWord) & vbLf & CStr(angl)
                                    Next i
                                End If
                                If IsArray(AngleProject) = True Then
                                    If AngleProject.GetUpperBound(1) > 0 Then
                                        Dim boolSort As Boolean = MathFunction.FuncSortDblArray(AngleProject, 0)
                                    End If
                                    For i As Integer = 0 To AngleProject.GetUpperBound(1)
                                        Dim angl As Double = Math.Round((AngleProject(1, i) * 180) / Math.PI, 0)
                                        If angl > 180 Then angl = 360 - angl
                                        ArrayDataWord(14, CountArrayDataWord) = ArrayDataWord(14, CountArrayDataWord) & vbLf & CStr(angl)
                                    Next i
                                End If
                                ArrayDataWord(15, CountArrayDataWord) = Owner  'владелец и адрес
                                ArrayDataWord(16, CountArrayDataWord) = " " & Adress 'адрес
                                ArrayDataWord(17, CountArrayDataWord) = HeightNet 'глубина заложения/провис провода
                                ArrayDataWord(18, CountArrayDataWord) = CountPipes 'ЧИСЛО труб, проводов
                                ArrayDataWord(19, CountArrayDataWord) = Power 'напряжение и рабочее давление
                                Dim StrMatS As String = ""
                                If MatNet.Trim.Length > 0 Then
                                    StrMatS = MatNet
                                End If
                                If SizeNet.Trim.Length > 0 Then
                                    StrMatS = StrMatS & " " & SizeNet
                                End If
                                ArrayDataWord(20, CountArrayDataWord) = StrMatS  'материал, сечение провода, диаметр
                                CountArrayDataWord += 1
                            End If
                        Catch ex As System.Exception
                        End Try
                    End If
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'записываем все в массив
            If IsArray(ArrayDataWord) = True Then
                If ArrayDataWord.GetUpperBound(1) > 0 Then
                    'сортируем массив
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayDataWord, 0)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем документ Word
                Try
                    Dim PutchDoc As String = PutchTbl & "Ведомость инженерных коммуникаций_Росавтодор.docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "Ведомость инженерных коммуникаций_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        locker.Dispose()
                        Exit Sub
                    End If
                Catch
                    MsgBox("Не удалось найти выбранную таблицу!!!")
                    locker.Dispose()
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try

                For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
                    For Each tabl As Word.Table In WDoc.Tables
                        Try
                            Dim countTableword As Integer = tabl.Rows.Count
                            If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
                                Exit For
                            End If
                            tabl.Rows.Add()
                            '===================================================================================
                            'название объекта
                            If IsNothing(ArrayDataWord(1, i)) = False Then
                                tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(1, i)
                            End If
                            '=================================================================================================
                            'владелец
                            Dim StrOwner As String = ""
                            If IsNothing(ArrayDataWord(15, i)) = False Then
                                If ArrayDataWord(15, i).Trim.Length > 0 Then
                                    StrOwner = ArrayDataWord(15, i).Trim
                                End If
                            End If
                            If IsNothing(ArrayDataWord(16, i)) = False Then
                                If ArrayDataWord(16, i).Trim.Length > 0 Then
                                    StrOwner = StrOwner & " " & ArrayDataWord(16, i).Trim
                                End If
                            End If
                            tabl.Cell(countTableword, 2).Range.Text = StrOwner.Trim   'владелец и адрес
                            '===================================================================================
                            'километр пересечения
                            Dim boolPer As Boolean = False
                            If IsNothing(ArrayDataWord(2, i)) = False Then
                                If ArrayDataWord(2, i).Trim.Length > 1 Then
                                    tabl.Cell(countTableword, 3).Range.Text = ArrayDataWord(2, i)
                                    boolPer = True
                                End If
                            End If
                            '===================================================================================
                            'записываем смещение относительно существующей трассы
                            If boolPer = False Then
                                Dim StrKMOFF As String = ""
                                If IsNothing(ArrayDataWord(4, i)) = False And IsNothing(ArrayDataWord(5, i)) = False Then
                                    Dim kmStart As String = FuncFormatZN.FuncFormatKM(Val(ArrayDataWord(4, i)))
                                    Dim kmoff1 As Double = Val(ArrayDataWord(5, i))
                                    If kmoff1 < 0 Then
                                        StrKMOFF = kmStart & " Лево " & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'влево
                                    ElseIf kmoff1 > 0 Then
                                        StrKMOFF = kmStart & " Право " & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'вправо
                                    End If
                                End If
                                If IsNothing(ArrayDataWord(6, i)) = False And IsNothing(ArrayDataWord(7, i)) = False Then
                                    Dim kmEnd As String = FuncFormatZN.FuncFormatKM(Val(ArrayDataWord(6, i)))
                                    Dim kmoff2 As Double = Val(ArrayDataWord(7, i))
                                    If kmoff2 < 0 Then
                                        StrKMOFF = StrKMOFF & vbLf & kmEnd & " Лево " & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'влево
                                    ElseIf kmoff2 > 0 Then
                                        StrKMOFF = StrKMOFF & vbLf & kmEnd & " Право " & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'вправо
                                    End If
                                End If
                                tabl.Cell(countTableword, 3).Range.Text = StrKMOFF
                            End If
                            '=================================================================================================
                            'записываем угол пересечения
                            If IsNothing(ArrayDataWord(13, i)) = False Then
                                tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(13, i)
                            End If
                            '=================================================================================================
                            'записываем пикет пересечения
                            Dim boolPer1 As Boolean = False
                            If IsNothing(ArrayDataWord(3, i)) = False Then
                                If ArrayDataWord(3, i).Trim.Length > 1 Then
                                    tabl.Cell(countTableword, 4).Range.Text = ArrayDataWord(3, i)
                                    boolPer1 = True
                                End If
                            End If
                            '=================================================================================================
                            'записываем смещение относительно проектной оси трассы
                            If boolPer = False Then
                                Dim StrPKOFF As String = ""
                                If IsNothing(ArrayDataWord(8, i)) = False And IsNothing(ArrayDataWord(9, i)) = False Then
                                    Dim pkStart As String = FuncFormatZN.FuncFormatPK(Val(ArrayDataWord(8, i)))
                                    Dim pkoff1 As Double = Val(ArrayDataWord(9, i))
                                    If pkoff1 < 0 Then
                                        StrPKOFF = pkStart & " Лево " & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'влево
                                    ElseIf pkoff1 > 0 Then
                                        StrPKOFF = pkStart & " Право " & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'вправо
                                    End If
                                End If
                                If IsNothing(ArrayDataWord(10, i)) = False And IsNothing(ArrayDataWord(11, i)) = False Then
                                    Dim pkEnd As String = FuncFormatZN.FuncFormatPK(Val(ArrayDataWord(10, i)))
                                    Dim pkoff2 As Double = Val(ArrayDataWord(11, i))
                                    If pkoff2 < 0 Then
                                        StrPKOFF = StrPKOFF & pkEnd & " Лево " & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'влево
                                    ElseIf pkoff2 > 0 Then
                                        StrPKOFF = StrPKOFF & pkEnd & " Право " & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'вправо
                                    End If
                                End If
                                tabl.Cell(countTableword, 4).Range.Text = StrPKOFF
                            End If
                            '=================================================================================================
                            'записываем угол пересечения
                            If IsNothing(ArrayDataWord(14, i)) = False Then
                                If ArrayDataWord(14, i).Trim.Length > 0 Then
                                    tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(14, i)
                                End If
                            End If
                            '===================================================================================================
                            'характеристика высота нижнего провода или глубина залегания
                            If IsNothing(ArrayDataWord(17, i)) = False Then
                                tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(17, i)
                            End If
                            '===================================================================================================
                            'характеристика число проводов или труб
                            If IsNothing(ArrayDataWord(18, i)) = False Then
                                tabl.Cell(countTableword, 7).Range.Text = ArrayDataWord(18, i)
                            End If
                            '===================================================================================================
                            'напряжение и рабочее давление
                            If IsNothing(ArrayDataWord(19, i)) = False Then
                                tabl.Cell(countTableword, 8).Range.Text = ArrayDataWord(19, i)
                            End If
                            '===================================================================================================
                            'материал и сечение провода
                            If IsNothing(ArrayDataWord(20, i)) = False Then
                                tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(20, i) 'описание
                            End If
                            Exit For
                        Catch ex As System.Exception
                        End Try
                    Next
                Next i
                WApp.Visible = True
            End If
            acTrans.Commit()
        End Using
    End Sub

    '==================================================================================================
    'формирование дефектной ведомости дорожных знаков
    <CommandMethod("DefVedDorZN")>
    Public Sub DefVedDorZN()
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        'открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim NumbPoint As String() = Nothing
        Dim CountPoint As Integer = 0
        'для дорожных знаков
        Const UDPname1 As String = "GOST_NUMBER1"
        Const UDPname2 As String = "GOST_NUMBER2"
        Const UDPname3 As String = "GOST_NUMBER3"
        Const UDPname4 As String = "GOST_NUMBER4"
        Const UDPname5 As String = "NUMBER_MILE"
        Const UDPname6 As String = "TEXT_SIGN"
        Const UDPname7 As String = "HEIGHT"
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            'смотрим пути доступа к файлам
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'БЛОК С НАЧАЛЬНЫМИ УСТАНОВКАМИ
            'находим все директории с файлами
            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
            'отделяем имя файла от папки
            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
            'определяем имя родительской папки для папки где находится сборка
            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
            'ищем пути для изображений знака
            Dim PutchDorZn As String = "d:\Civil3DToolsUtility\FileResources\RoadZnaks\"
            If IO.Directory.Exists(parentPatch2 & "\FileResources\RoadZnaks\") = True Then
                PutchDorZn = parentPatch2 & "\FileResources\RoadZnaks\"
            End If
            'ищем пути для таблиц
            Dim PutchTbl As String = "d:\Civil3DToolsUtility\TemplateTables\Survey\"
            If File.Exists(parentPatch2 & "\TemplateTables\Survey\Ведомость существующих дорожных знаков.docx") = True Then
                PutchTbl = parentPatch2 & "\TemplateTables\Survey\"
            Else
                If File.Exists("d:\Civil3DToolsUtility\TemplateTables\Survey\Ведомость существующих дорожных знаков.docx") = False Then
                    MsgBox("Шаблон не найден! Макро прервано.")
                    Exit Sub
                End If
            End If
            '=========================================================================================
            'максимальное расстояние от оси трассы назначаем
            Dim LenghtDorZn As Double = 1000000000000000
            Dim PLineObl As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing 'область
            Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
            pKeyOpts.Message = vbLf & "Укажите максимальное расстояние от оси дороги до дорожного знака "
            pKeyOpts.Keywords.Add("Указать")
            pKeyOpts.Keywords.Add("Расстояние")
            pKeyOpts.Keywords.Add("Область")
            pKeyOpts.Keywords.Add("Нет")
            pKeyOpts.Keywords.Default = "Нет"
            pKeyOpts.AllowNone = False
            Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
            If Len(pKeyRes.StringResult.Trim) = 0 Then
                Exit Sub
            ElseIf pKeyRes.StringResult Like "Указать" Then 'указываются 2 точки и вычисляется расстояние
                Dim resPoint As PromptPointResult = Nothing
                Dim optPoint As PromptPointOptions = New PromptPointOptions(vbLf & "Первая точка: ")
                resPoint = ed.GetPoint(optPoint)
                Dim Point1 As Point3d = resPoint.Value
                Dim Pt12d As Point2d = New Point2d(Point1.X, Point1.Y)
                Dim optPoint1 As PromptPointOptions = New PromptPointOptions(vbLf & "Вторая точка: ")
                resPoint = ed.GetPoint(optPoint1)
                Dim Point2 As Point3d = resPoint.Value
                Dim Pt22d As Point2d = New Point2d(Point2.X, Point2.Y)
                LenghtDorZn = Pt12d.GetVectorTo(Pt22d).Length
            ElseIf pKeyRes.StringResult Like "Расстояние" Then
                Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Положительное число: ")
                pStrOpts.AllowSpaces = True
                Dim pStrRes As PromptResult = ed.GetString(pStrOpts)
                LenghtDorZn = Val(pStrRes.StringResult)
                LenghtDorZn = Math.Abs(LenghtDorZn)
            ElseIf pKeyRes.StringResult Like "Область" Then
                Dim SelectPline As PromptEntityResult = ed.GetEntity(vbLf & "Выберите полилинию:  ")
                Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectPline.ObjectId, OpenMode.ForRead)
                If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    PLineObl = acEnt1
                End If
            Else
                LenghtDorZn = 1000000000000000
            End If
            '=============================================================================================
            'выбираем ось дороги (фактическую и пректную)
            Dim align1 As Alignment = Nothing 'существующая дорога
            Dim align2 As Alignment = Nothing 'проектная ось

            Dim selectAlignResult As PromptEntityResult = ed.GetEntity(vbCrLf + "Выберите ось существующей дороги")
            If selectAlignResult.Status = PromptStatus.OK Then
                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(selectAlignResult.ObjectId, OpenMode.ForRead)
                align1 = TryCast(dbObj, Alignment)
            End If

            Dim selectAlignResult2 As PromptEntityResult = ed.GetEntity(vbCrLf + "Выберите проектную ось")
            If selectAlignResult2.Status = PromptStatus.OK Then
                Dim dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(selectAlignResult2.ObjectId, OpenMode.ForRead)
                align2 = TryCast(dbObj, Alignment)
            End If
            '====================================================================================================
            '====================================================================================================
            'массив для отчета
            Dim ArrayZn As String(,) = Nothing
            Dim CountArrayZn As Integer = 0
            '0-номер п\п
            '1-координата x
            '2-координата y
            '3-километр существующий
            '4-смещение
            '5-пикет проектный
            '6-смещение
            '7-индекс знака1
            '8-индекс знака2
            '9-индекс знака3
            '====================================================================================================
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            'выбираем объекты
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead, False) 'открываем объект на чтение
                        If TypeOf ent Is CogoPoint Then
                            Dim PTCogo As CogoPoint = ent

                        End If
                    End If
                Next
            End If







            '            Dim UDP1 As UDPString = FindUDPStr(UDPname1, "ROAD_SIGN")
            '            Dim UDP2 As UDPString = FindUDPStr(UDPname2, "ROAD_SIGN")
            '            Dim UDP3 As UDPString = FindUDPStr(UDPname3, "ROAD_SIGN")
            '            Dim UDP4 As UDPString = FindUDPStr(UDPname4, "ROAD_SIGN")

            '            'для километровых знаков
            '            Dim UDP5 As UDPString = FindUDPStr(UDPname5, "MILE_STONE")
            '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '            'начинаем перебор точек
            '            For i As Integer = 0 To cogoPoints.Length - 1
            '                Dim PK As Double = -1 'пикет проектный
            '                Dim OfsetPK As Double = -1 'смещение проектное
            '                Dim idCogo As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(cogoPoints(i)) 'находим id точки
            '                Dim PointCOGO1 As CogoPoint = idCogo.GetObject(OpenMode.ForRead) 'ссылаемся на объект
            '                'проверяем точку на наличе в области
            '                If IsNothing(PLineObl) = False Then
            '                    'проверяем точку на принадлежность области
            '                    Dim Pt2d As Point2d = New Point2d(PointCOGO1.Easting, PointCOGO1.Northing)
            '                    Dim boolCrossPt As Boolean = FuncCrossPoint(PLineObl, Pt2d)
            '                    If boolCrossPt = False Then
            '                        GoTo Line5
            '                    End If
            '                Else
            '                    If IsNothing(align2) = False Then
            '                        Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(align2, PointCOGO1.Location, PK, OfsetPK) 'определям пикет и смещение
            '                        If OfsetPK > LenghtDorZn Then
            '                            GoTo Line5
            '                        End If
            '                    End If
            '                End If
            '                '=======================================================================================
            '                'записываем характеристики в массив
            '                ReDim Preserve ArrayZn(18, CountArrayZn)
            '                ArrayZn(0, CountArrayZn) = CountArrayZn + 1 'номер
            '                ArrayZn(1, CountArrayZn) = Math.Round(PointCOGO1.Northing, 2) 'координата x
            '                ArrayZn(2, CountArrayZn) = Math.Round(PointCOGO1.Easting, 2) 'координата y
            '                'определяем пикет и смещение для проектной оси
            '                If IsNothing(align2) = False Then
            '                    Dim RawPK As Double = -1
            '                    Dim RawOffset As Double = -1
            '                    Try
            '                        align2.StationOffset(PointCOGO1.Easting, PointCOGO1.Northing, RawPK, RawOffset)
            '                        Dim PiketNPKStrFormat As String = FuncFormatPiketSTEq(align2, RawPK)
            '                        ArrayZn(5, CountArrayZn) = PiketNPKStrFormat 'проектный пикет
            '                        ArrayZn(6, CountArrayZn) = RawOffset 'смещение
            '                    Catch ex As Autodesk.Civil.PointNotOnEntityException
            '                    End Try
            '                End If
            '                'для существующей оси
            '                If IsNothing(align1) = True Then
            '                    ArrayZn(3, CountArrayZn) = 0 'существующий километр
            '                    ArrayZn(4, CountArrayZn) = 0 'смещение
            '                Else
            '                    Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(align1, PointCOGO1.Location, PK, OfsetPK)
            '                    If PkAlign = False Then
            '                        ArrayZn(3, CountArrayZn) = 0 'существующий километр
            '                        ArrayZn(4, CountArrayZn) = 0 'смещение
            '                    Else
            '                        ArrayZn(3, CountArrayZn) = PK 'существующий километр
            '                        ArrayZn(4, CountArrayZn) = OfsetPK 'смещение
            '                    End If
            '                End If
            '                '===========================================================================================
            '                'ищем номера и названия знаков
            '                Dim ArrayStrRAwDesk As String() = PointCOGO1.RawDescription.Split()
            '                If ArrayStrRAwDesk(0) Like "KMZN" Then
            '                    Dim indZn5 As String = "" 'номер знака по ГОСТ
            '                    Try
            '                        indZn5 = PointCOGO1.GetUDPValue(UDP5)
            '                        ArrayZn(7, CountArrayZn) = "6.13"
            '                        ArrayZn(8, CountArrayZn) = "Километровый знак. КМ. " & indZn5
            '                        ArrayZn(9, CountArrayZn) = PutchDorZn & "6 Информационно-указательные знаки\Километровый знак (6.13).png"
            '                    Catch ex As System.InvalidOperationException
            '                    End Try
            '                Else
            '                    Dim indZn1 As String = "" 'номер знака по ГОСТ
            '                    Dim Putsh As String = "" 'путь к файлу
            '                    Dim NameZn As String = "" 'название знака
            '                    Try
            '                        indZn1 = PointCOGO1.GetUDPValue(UDP1)
            '                        If indZn1.Trim.Length > 0 Then
            '                            Dim NameZn1 As String = FuncFindFileDorZn(indZn1, Putsh, NameZn, PutchDorZn)
            '                            ArrayZn(7, CountArrayZn) = indZn1
            '                            ArrayZn(8, CountArrayZn) = NameZn
            '                            ArrayZn(9, CountArrayZn) = Putsh
            '                        End If
            '                    Catch ex As System.InvalidOperationException
            '                    End Try
            '                    'номер знака 2
            '                    Dim indZn2 As String = "" 'номер знака по ГОСТ
            '                    Try
            '                        indZn2 = PointCOGO1.GetUDPValue(UDP2)
            '                        If indZn2.Trim.Length > 0 Then
            '                            Dim NameZn2 As String = FuncFindFileDorZn(indZn2, Putsh, NameZn, PutchDorZn)
            '                            ArrayZn(10, CountArrayZn) = indZn2
            '                            ArrayZn(11, CountArrayZn) = NameZn
            '                            ArrayZn(12, CountArrayZn) = Putsh
            '                        End If
            '                    Catch ex As System.InvalidOperationException
            '                    End Try
            '                    'номер знака 3
            '                    Dim indZn3 As String = "" 'номер знака по ГОСТ
            '                    Try
            '                        indZn3 = PointCOGO1.GetUDPValue(UDP3)
            '                        If indZn3.Trim.Length > 0 Then
            '                            Dim NameZn3 As String = FuncFindFileDorZn(indZn3, Putsh, NameZn, PutchDorZn)
            '                            ArrayZn(13, CountArrayZn) = indZn3
            '                            ArrayZn(14, CountArrayZn) = NameZn
            '                            ArrayZn(15, CountArrayZn) = Putsh
            '                        End If
            '                    Catch ex As System.InvalidOperationException
            '                    End Try
            '                    'номер знака 4
            '                    Dim indZn4 As String = "" 'номер знака по ГОСТ
            '                    Try
            '                        indZn4 = PointCOGO1.GetUDPValue(UDP4)
            '                        If indZn4.Trim.Length > 0 Then
            '                            Dim NameZn3 As String = FuncFindFileDorZn(indZn4, Putsh, NameZn, PutchDorZn)
            '                            ArrayZn(16, CountArrayZn) = indZn4
            '                            ArrayZn(17, CountArrayZn) = NameZn
            '                            ArrayZn(18, CountArrayZn) = Putsh
            '                        End If
            '                    Catch ex As System.InvalidOperationException
            '                    End Try
            '                End If
            '                CountArrayZn += 1
            'Line5:
            '            Next i
            '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '            'фильтруем блоки автокад
            '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Line1:
            '            Dim CountRow As Integer = 0
            '            If IsArray(ArrayZn) = True Then
            '                CountRow = ArrayZn.GetUpperBound(1) + 1
            '            End If
            '            'имена блоков которые подлежат фильтрации
            '            Dim NameBlk1 As String = "A-Знак дорожный на двух стойках"
            '            Dim NameBlk2 As String = "A-Знак дорожный километровый"
            '            Dim NameBlk3 As String = "A-Знак дорожный на одной стойке"
            '            Dim NameBlk4 As String = "A-Знак дорожный подвесной"
            '            Dim NameBlk5 As String = "A-Знак дорожный расположенный на опоре"
            '            Dim NameBlk6 As String = "A-Знак-указатель населенных пунктов на двух стойках"
            '            Dim NameBlk7 As String = "A-Знак-указатель населенных пунктов на одной стойке"
            '            '==============================================================================
            '            'формируем фильтр
            '            Dim BlkSelect As PromptSelectionResult
            '            Dim ssetBlk As SelectionSet = Nothing
            '            Dim ValType1(2) As TypedValue
            '            ValType1.SetValue(New TypedValue(DxfCode.Start, "INSERT"), 0)
            '            ValType1.SetValue(New TypedValue(DxfCode.Visibility, 0), 1)
            '            ValType1.SetValue(New TypedValue(67, 0), 2)
            '            Dim FilterParcel As SelectionFilter = New SelectionFilter(ValType1)
            '            BlkSelect = ed.SelectAll(FilterParcel)
            '            'смотрим что попалось
            '            If BlkSelect.Status = PromptStatus.OK Then
            '                ssetBlk = BlkSelect.Value
            '                Dim countObj As Integer = ssetBlk.Count
            '                If countObj = 0 Then GoTo line4
            '                For Each idobj As ObjectId In ssetBlk.GetObjectIds
            '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(idobj, OpenMode.ForRead)
            '                    If Not IsDBNull(acEnt) Then
            '                        If TypeOf acEnt Is BlockReference Then
            '                            Dim boolBlk As Boolean = False
            '                            Dim obj As BlockReference = acEnt
            '                            Dim NameBLk As String = ""
            '                            Dim blockDyn As BlockTableRecord = Nothing
            '                            If obj.IsDynamicBlock = True Then
            '                                Dim blockRef As BlockReference = acTrans.GetObject(idobj, OpenMode.ForRead)
            '                                blockDyn = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
            '                                If blockDyn.Name Like NameBlk1 Or blockDyn.Name Like NameBlk2 Or blockDyn.Name Like NameBlk3 Or blockDyn.Name Like NameBlk4 Or blockDyn.Name Like NameBlk5 Or blockDyn.Name Like NameBlk6 Or blockDyn.Name Like NameBlk7 Then
            '                                    NameBLk = blockDyn.Name
            '                                    boolBlk = True
            '                                End If
            '                            Else
            '                                If obj.Name Like NameBlk1 Or obj.Name Like NameBlk2 Or obj.Name Like NameBlk3 Or obj.Name Like NameBlk4 Or obj.Name Like NameBlk5 Or obj.Name Like NameBlk6 Or obj.Name Like NameBlk7 Then
            '                                    NameBLk = obj.Name
            '                                    boolBlk = True
            '                                End If
            '                            End If
            '                            If boolBlk = True Then
            '                                Dim PtXY As Point3d = obj.Position
            '                                Dim PK As Double = 0
            '                                Dim OfsetPK As Double = 0
            '                                'делаем проверку
            '                                If IsNothing(PLineObl) = False Then
            '                                    'проверяем точку на принадлежность области
            '                                    Dim Pt2d As Point2d = New Point2d(obj.Position.X, obj.Position.Y)
            '                                    Dim boolCrossPt As Boolean = FuncCrossPoint(PLineObl, Pt2d)
            '                                    If boolCrossPt = False Then
            '                                        GoTo Line6
            '                                    End If
            '                                Else
            '                                    If IsNothing(align2) = False Then
            '                                        Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(align2, PtXY, PK, OfsetPK) 'определям пикет и смещение
            '                                        If OfsetPK > LenghtDorZn Then
            '                                            GoTo Line6
            '                                        End If
            '                                    End If
            '                                End If
            '                                'записываем характеристики знака
            '                                ReDim Preserve ArrayZn(18, CountArrayZn)
            '                                ArrayZn(0, CountArrayZn) = CountRow 'номер
            '                                ArrayZn(1, CountArrayZn) = Math.Round(obj.Position.Y, 2) 'координата x
            '                                ArrayZn(2, CountArrayZn) = Math.Round(obj.Position.X, 2) 'координата y
            '                                'записываем положение относительно проектной оси
            '                                If IsNothing(align2) = False Then
            '                                    Dim RawPK As Double = -1
            '                                    Dim RawOffset As Double = -1
            '                                    Try
            '                                        align2.StationOffset(PtXY.X, PtXY.Y, RawPK, RawOffset)
            '                                        Dim PiketNPKStrFormat As String = FuncFormatPiketSTEq(align2, RawPK)
            '                                        ArrayZn(5, CountArrayZn) = PiketNPKStrFormat 'проектный пикет
            '                                        ArrayZn(6, CountArrayZn) = RawOffset 'смещение
            '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
            '                                    End Try
            '                                End If
            '                                'для существующей оси
            '                                If IsNothing(align1) = True Then
            '                                    ArrayZn(3, CountArrayZn) = 0 'существующий километр
            '                                    ArrayZn(4, CountArrayZn) = 0 'смещение
            '                                Else
            '                                    Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(align1, PtXY, PK, OfsetPK)
            '                                    If PkAlign = False Then
            '                                        ArrayZn(3, CountArrayZn) = 0 'существующий километр
            '                                        ArrayZn(4, CountArrayZn) = 0 'смещение
            '                                    Else
            '                                        ArrayZn(3, CountArrayZn) = FuncFormatKM(PK) 'существующий километр
            '                                        ArrayZn(4, CountArrayZn) = OfsetPK 'смещение
            '                                    End If
            '                                End If
            '                                'считываем атрибуты
            '                                Dim DorZnAttr As String(,) = Nothing

            '                                Dim BoolReadAttr As Boolean = FuncReadAllBlockAttr(obj, DorZnAttr)
            '                                If IsArray(DorZnAttr) = True Then
            '                                    If NameBLk Like "A-Знак дорожный километровый" Then
            '                                        ArrayZn(7, CountArrayZn) = "6.13"
            '                                        ArrayZn(8, CountArrayZn) = "Километровый знак. КМ. " & DorZnAttr(1, 0)
            '                                        ArrayZn(9, CountArrayZn) = PutchDorZn & "6 Информационно-указательные знаки\Километровый знак (6.13).png"
            '                                    Else
            '                                        For i2 As Integer = 0 To DorZnAttr.GetUpperBound(1)
            '                                            Dim ind As String = DorZnAttr(1, i2) 'находим индекс
            '                                            Dim Putsh As String = ""
            '                                            Dim NameZn As String = ""
            '                                            Dim NameZn1 As String = FuncFindFileDorZn(ind, Putsh, NameZn, PutchDorZn) 'находим описание дорожного знака
            '                                            If NameZn.Length > 0 Then
            '                                                If i2 = 0 Then
            '                                                    ArrayZn(7, CountArrayZn) = ind
            '                                                    ArrayZn(8, CountArrayZn) = NameZn
            '                                                    ArrayZn(9, CountArrayZn) = Putsh
            '                                                ElseIf i2 = 1 Then
            '                                                    ArrayZn(10, CountArrayZn) = ind
            '                                                    ArrayZn(11, CountArrayZn) = NameZn
            '                                                    ArrayZn(12, CountArrayZn) = Putsh
            '                                                ElseIf i2 = 2 Then
            '                                                    ArrayZn(13, CountArrayZn) = ind
            '                                                    ArrayZn(14, CountArrayZn) = NameZn
            '                                                    ArrayZn(15, CountArrayZn) = Putsh
            '                                                ElseIf i2 = 3 Then
            '                                                    ArrayZn(16, CountArrayZn) = ind
            '                                                    ArrayZn(17, CountArrayZn) = NameZn
            '                                                    ArrayZn(18, CountArrayZn) = Putsh
            '                                                End If
            '                                            End If

            '                                        Next i2
            '                                    End If
            '                                End If
            '                                CountArrayZn += 1
            '                            End If
            '                        End If
            '                    End If
            'Line6:
            '                Next
            '            End If
            'line4:

            '            If IsNothing(ArrayZn) = True Then
            '                MsgBox("Программе не удалось выявить ни одного дорожного знака!!! Макро прервано.")
            '                Exit Sub
            '            End If
            '            'делаем сортировку массива по пикетажу
            '            If IsNothing(align1) = False Then
            '                Dim sortArray As Boolean = FuncSortArray3(ArrayZn, 3)
            '            Else
            '                Dim sortArray As Boolean = FuncSortArray3(ArrayZn, 5)
            '            End If
            '            'проверяем есть ли такой шаблон (проверяем наличие файла)
            '            Dim FileTable As FileInfo = FileIO.FileSystem.GetFileInfo(PutchTbl & "Ведомость существующих дорожных знаков.docx")
            '            If FileTable.Exists = True Then
            '                Try
            '                    WApp = CreateObject("Word.Application")
            '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость существующих дорожных знаков.docx")
            '                Catch
            '                    MsgBox("Не удалось найти выбранную таблицу!!!")
            '                    GoTo Line1
            '                End Try
            '                WApp.Visible = True
            '                WApp.Dialogs(Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFileSaveAs).Show()
            '                Dim fName As String = WDoc.Name
            '                For i As Integer = 0 To ArrayZn.GetUpperBound(1)
            '                    For Each tabl As Word.Table In WDoc.Tables
            '                        Dim countTableword As Integer = tabl.Rows.Count
            '                        tabl.Rows.Add()
            '                        tabl.Cell(countTableword, 1).Range.Text = i + 1 'номер
            '                        tabl.Cell(countTableword, 2).Range.Text = ArrayZn(1, i) 'координата x
            '                        tabl.Cell(countTableword, 3).Range.Text = ArrayZn(2, i) 'координата y
            '                        'существующий километр
            '                        If ArrayZn(3, i).Trim.Length = 0 Then
            '                            tabl.Cell(countTableword, 4).Range.Text = "-"
            '                        Else
            '                            tabl.Cell(countTableword, 4).Range.Text = ArrayZn(3, i) 'существующий километр
            '                        End If
            '                        Dim offsetPK1 As Double = Val(ArrayZn(4, i)) 'смещение
            '                        offsetPK1 = Math.Round(offsetPK1, 2)
            '                        If offsetPK1 < 0 Then
            '                            tabl.Cell(countTableword, 5).Range.Text = Math.Abs(offsetPK1) 'смещение влево
            '                        ElseIf offsetPK1 > 0 Then
            '                            tabl.Cell(countTableword, 6).Range.Text = Math.Abs(offsetPK1) 'смещение вправо
            '                        Else
            '                            tabl.Cell(countTableword, 6).Range.Text = "-"
            '                        End If
            '                        '=============================================================
            '                        'проектный километр
            '                        tabl.Cell(countTableword, 7).Range.Text = ArrayZn(5, i)  'проектный километр
            '                        Dim offsetPK2 As Double = Val(ArrayZn(6, i)) 'смещение
            '                        offsetPK2 = Math.Round(offsetPK2, 2)
            '                        If offsetPK2 < 0 Then
            '                            tabl.Cell(countTableword, 8).Range.Text = Math.Abs(offsetPK2) 'смещение влево
            '                        ElseIf offsetPK2 > 0 Then
            '                            tabl.Cell(countTableword, 9).Range.Text = Math.Abs(offsetPK2) 'смещение вправо
            '                        Else
            '                            tabl.Cell(countTableword, 9).Range.Text = "-"
            '                        End If

            '                        Dim strInd As String = ""
            '                        'заполняем индексы согласно ГОСТ
            '                        If Not (IsNothing(ArrayZn(7, i))) Then
            '                            strInd = ArrayZn(7, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(10, i))) Then
            '                            strInd = strInd & vbLf & ArrayZn(10, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(13, i))) Then
            '                            strInd = strInd & vbLf & ArrayZn(13, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(16, i))) Then
            '                            strInd = strInd & vbLf & ArrayZn(16, i)
            '                        End If
            '                        If strInd.Length = 0 Then strInd = "-"
            '                        tabl.Cell(countTableword, 10).Range.Text = strInd
            '                        'заполняем названия знаков согласно ГОСТ
            '                        Dim strName As String = ""
            '                        If Not (IsNothing(ArrayZn(8, i))) Then
            '                            strName = ArrayZn(8, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(11, i))) Then
            '                            strName = strName & vbLf & ArrayZn(11, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(14, i))) Then
            '                            strName = strName & vbLf & ArrayZn(14, i)
            '                        End If
            '                        If Not (IsNothing(ArrayZn(17, i))) Then
            '                            strName = strName & vbLf & ArrayZn(17, i)
            '                        End If
            '                        If strName.Length = 0 Then strName = "-"
            '                        tabl.Cell(countTableword, 11).Range.Text = strName
            '                        'заполняем графу с изображениями
            '                        Dim pic1 As String = ArrayZn(9, i)
            '                        If IsNothing(pic1) = False Then
            '                            If File.Exists(pic1) = True Then
            '                                tabl.Cell(countTableword, 12).Range.InlineShapes.AddPicture(pic1)
            '                            End If
            '                        End If
            '                        Dim pic2 As String = ArrayZn(12, i)
            '                        If IsNothing(pic2) = False Then
            '                            If File.Exists(pic2) = True Then
            '                                tabl.Cell(countTableword, 12).Range.InlineShapes.AddPicture(pic2)
            '                            End If
            '                        End If
            '                        Dim pic3 As String = ArrayZn(15, i)
            '                        If IsNothing(pic3) = False Then
            '                            If File.Exists(pic3) = True Then
            '                                tabl.Cell(countTableword, 12).Range.InlineShapes.AddPicture(pic3)
            '                            End If
            '                        End If
            '                        Dim pic4 As String = ArrayZn(18, i)
            '                        If IsNothing(pic4) = False Then
            '                            If File.Exists(pic4) = True Then
            '                                tabl.Cell(countTableword, 12).Range.InlineShapes.AddPicture(pic4)
            '                            End If
            '                        End If
            '                        tabl.Cell(countTableword, 13).Range.Text = "Состояние удовлетворительное"
            '                        Exit For
            '                    Next
            '                Next i
            '            End If
            '            acTrans.Commit()
        End Using
        locker.Dispose()
        WDoc.Close()
        WDoc = Nothing
        WApp.Quit()
        WApp = Nothing
    End Sub
    '    '==================================================================================================
    '    'Ведомость барьерного ограждения на автодорогах
    '    <CommandMethod("STPR-VedRoadBarrierd")>
    '    Public Sub VedRoadBarrier()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayDataWord As String(,) = Nothing
    '        Dim CountArrayDataWord As Integer = 0
    '        Dim boolRez As Boolean = False
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость размещения барьерного ограждения.docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость размещения барьерного ограждения.docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'забираем в массив все Характерные линии с расширенными данными Road_Barrier
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'выбираем объекты
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim km1 As Double = -1
    '                        Dim kmoff1 As Double = -1
    '                        Dim km2 As Double = -1
    '                        Dim kmoff2 As Double = -1
    '                        Dim pk1 As Double = -1
    '                        Dim pkoff1 As Double = -1
    '                        Dim pk2 As Double = -1
    '                        Dim pkoff2 As Double = -1
    '                        Dim LLine As Double = -1
    '                        Dim Name As String = ""
    '                        Dim LenObj As Double = 0
    '                        Dim TypeObj As String = "-"
    '                        Dim Mat As String = ""
    '                        Dim Height As Double = 0
    '                        Dim State As String = ""
    '                        Dim Note As String = ""
    '                        Dim StyleType As String = ""
    '                        Erase ArrayData
    '                        boolRez = False
    '                        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
    '                            Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
    '                            If FLine.StyleName Like "ИИ-17_Барьерное ограждение на автодороге более 1 метра" Or FLine.StyleName Like "ИИ-16_Барьерное ограждение на автодороге менее 1 метра" Or FLine.StyleName Like "ИИ-18_Перильное ограждение на автодороге" Then
    '                                StyleType = FLine.StyleName
    '                                boolRez = True
    '                                Dim pos As Integer = StyleType.IndexOf("_")
    '                                If pos > -1 Then
    '                                    Name = Mid(StyleType, pos)
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "ROAD_BARRIERS")
    '                                LLine = FLine.Length2D
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                'превращаем характерныю линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                colPt = FLine.GetPoints(FeatureLinePointType.PIPoint)
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                Try
    '                                    AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(acPoly) = False Then
    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'ищем пикетаж (существующий километр) конец участка
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                Try
    '                                    AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(acPoly) = False Then
    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(acPoly) = False Then
    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'проектный пикетаж (конец участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(acPoly) = False Then
    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                            End If
    '                            '3d полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "L16" Or NameLayer Like "L17" Or NameLayer = "L18" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "ROAD_BARRIERS")
    '                                If NameLayer Like "L17" Then
    '                                    Name = "Барьерное ограждение на автодороге более 1 метра"
    '                                ElseIf NameLayer Like "L16" Then
    '                                    Name = "Барьерное ограждение на автодороге менее 1 метра"
    '                                ElseIf NameLayer Like "L18" Then
    '                                    Name = "Перильное ограждение на автодороге"
    '                                End If
    '                                Dim LPipe As Double = FLine.Length
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                Try
    '                                    AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'ищем пикетаж (существующий километр) конец участка
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                Try
    '                                    AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'проектный пикетаж (конец участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                            End If
    '                            '2d полилинии
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Or NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Or NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                    Name = "Перильное ограждение на автодороге"
    '                                    Mat = "Металлическое"
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "ROAD_BARRIERS")
    '                                Dim LPipe As Double = FLine.Length
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                Try
    '                                    AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'ищем пикетаж (существующий километр) конец участка
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                Try
    '                                    AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'проектный пикетаж (конец участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                            End If
    '                            'полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Or NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Or NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                    Name = "Перильное ограждение на автодороге"
    '                                    Mat = "Металлическое"
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "ROAD_BARRIERS")
    '                                Dim LPipe As Double = FLine.Length
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                Try
    '                                    AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'ищем пикетаж (существующий километр) конец участка
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                Try
    '                                    AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'проектный пикетаж (конец участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                            End If
    '                            'линия
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Line = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Or NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Or NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Or NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                If NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение бетонное на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Бетонное"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой менее 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой менее 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Барьерное ограждение металлическое на автодорогах высотой более 1м" Then
    '                                    Name = "Барьерное ограждение на автодорогах высотой более 1м"
    '                                    Mat = "Металлическое"
    '                                ElseIf NameTypeLine Like "A-Перильное ограждение на автодороге" Then
    '                                    Name = "Перильное ограждение на автодороге"
    '                                    Mat = "Металлическое"
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "ROAD_BARRIERS")
    '                                Dim LPipe As Double = FLine.Length
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                Try
    '                                    AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'ищем пикетаж (существующий километр) конец участка
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                Try
    '                                    AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                        Try
    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                                'проектный пикетаж (конец участка)
    '                                Try
    '                                    AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    'конец линии за пределами трассы
    '                                    'ищем пересечение с началом трассы
    '                                    If IsNothing(FLine) = False Then
    '                                        Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                        Try
    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                    End If
    '                                End Try
    '                            End If
    '                        End If
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        '0-наименование, 1-КМ начало, 2-КМ конец, 3-ПК начало, 4-ПК конец, 5-Протяженность, 6-Тип ограждения, 7-Высота, 8-материал
    '                        '===========================================================================
    '                        'записываем OD Характеристики
    '                        If boolRez = True Then
    '                            If IsArray(ArrayData) = True Then
    '                                For j As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                    Dim NameF As String = ArrayData(1, j)
    '                                    Dim val As String = ArrayData(2, j)
    '                                    If NameF Like "MATERIAL" Then
    '                                        Mat = val
    '                                    ElseIf NameF Like "LABEL" Then
    '                                        Name = val
    '                                    ElseIf NameF Like "TYPE" Then
    '                                        TypeObj = val
    '                                    ElseIf NameF Like "HEIGHT" Then
    '                                        Try
    '                                            If IsNumeric(val) Then
    '                                                Height = CDbl(val)
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "STATE" Then
    '                                        State = val
    '                                    ElseIf NameF Like "NOTE" Then
    '                                        Note = val
    '                                    End If
    '                                Next j
    '                            End If
    '                            'начинаем разбор найденных свойств
    '                            ReDim Preserve ArrayDataWord(13, CountArrayDataWord)
    '                            ArrayDataWord(0, CountArrayDataWord) = Name
    '                            'записываем км
    '                            If km1 <> -1 And km2 <> -1 Then
    '                                If km1 > km2 Then
    '                                    ArrayDataWord(1, CountArrayDataWord) = km2
    '                                    ArrayDataWord(2, CountArrayDataWord) = kmoff2
    '                                    ArrayDataWord(3, CountArrayDataWord) = km1
    '                                    ArrayDataWord(4, CountArrayDataWord) = kmoff1
    '                                Else
    '                                    ArrayDataWord(1, CountArrayDataWord) = km1
    '                                    ArrayDataWord(2, CountArrayDataWord) = kmoff1
    '                                    ArrayDataWord(3, CountArrayDataWord) = km2
    '                                    ArrayDataWord(4, CountArrayDataWord) = kmoff2
    '                                End If
    '                            End If
    '                            If pk1 <> -1 And pk2 <> -1 Then
    '                                If pk1 > pk2 Then
    '                                    ArrayDataWord(5, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(6, CountArrayDataWord) = pkoff2
    '                                    ArrayDataWord(7, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(8, CountArrayDataWord) = pkoff1
    '                                Else
    '                                    ArrayDataWord(5, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(6, CountArrayDataWord) = pkoff1
    '                                    ArrayDataWord(7, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(8, CountArrayDataWord) = pkoff2
    '                                End If
    '                            End If
    '                            ArrayDataWord(9, CountArrayDataWord) = Math.Round(LLine, 2) 'протяженность
    '                            ArrayDataWord(10, CountArrayDataWord) = TypeObj  'тип ограждения
    '                            ArrayDataWord(11, CountArrayDataWord) = Height 'высота ограждения
    '                            ArrayDataWord(12, CountArrayDataWord) = Note '
    '                            ArrayDataWord(13, CountArrayDataWord) = Mat '
    '                            CountArrayDataWord += 1
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'записываем все в массив
    '            If IsArray(ArrayDataWord) = True Then
    '                'сортируем массив
    '                If IsNothing(AlignEX) = False Then
    '                    Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 1)
    '                ElseIf IsNothing(AlignProject) = False Then
    '                    Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 5)
    '                End If
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость размещения барьерного ограждения.docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость размещения барьерного ограждения_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                End Try
    '                For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
    '                            Exit For
    '                        End If
    '                        tabl.Rows.Add()
    '                        'название объекта
    '                        tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(0, i)
    '                        '==================================================================================
    '                        'форматируем километр
    '                        Dim km As String = "-"
    '                        If IsNumeric(ArrayDataWord(1, i)) = True Then
    '                            km = FuncFormatKM(Val(ArrayDataWord(1, i)))
    '                        Else
    '                            If ArrayDataWord(1, i).Trim.Length > 0 Then
    '                                km = ArrayDataWord(1, i)
    '                            End If
    '                        End If
    '                        'находим смещение
    '                        'КМ начала
    '                        Dim offStr As String = ""
    '                        Dim kmoff1 As Double = 0
    '                        If IsNumeric(ArrayDataWord(2, i)) = True Then
    '                            kmoff1 = Val(ArrayDataWord(2, i))
    '                            If kmoff1 < 0 Then
    '                                offStr = "Лево " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            ElseIf kmoff1 > 0 Then
    '                                offStr = "Право " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            End If
    '                        End If
    '                        tabl.Cell(countTableword, 2).Range.Text = km & vbLf & offStr

    '                        'КМ конца
    '                        km = "-"
    '                        If IsNumeric(ArrayDataWord(3, i)) = True Then
    '                            km = FuncFormatKM(Val(ArrayDataWord(3, i)))
    '                        Else
    '                            If ArrayDataWord(3, i).Trim.Length > 0 Then
    '                                km = ArrayDataWord(3, i)
    '                            End If
    '                        End If
    '                        offStr = ""
    '                        kmoff1 = 0
    '                        If IsNumeric(ArrayDataWord(4, i)) = True Then
    '                            kmoff1 = Val(ArrayDataWord(4, i))
    '                            If kmoff1 < 0 Then
    '                                offStr = "Лево " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            ElseIf kmoff1 > 0 Then
    '                                offStr = "Право " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            End If
    '                        End If
    '                        tabl.Cell(countTableword, 3).Range.Text = km & vbLf & offStr
    '                        '=============================================================================================
    '                        'ПК начала
    '                        km = "-"
    '                        If IsNumeric(ArrayDataWord(5, i)) = True Then
    '                            km = FuncFormatPK2(Val(ArrayDataWord(5, i)))
    '                        Else
    '                            If ArrayDataWord(5, i).Trim.Length > 0 Then
    '                                km = ArrayDataWord(5, i)
    '                            End If
    '                        End If
    '                        offStr = ""
    '                        kmoff1 = 0
    '                        If IsNumeric(ArrayDataWord(6, i)) = True Then
    '                            kmoff1 = Val(ArrayDataWord(6, i))
    '                            If kmoff1 < 0 Then
    '                                offStr = "Лево " & String.Format("{0:0.00}", -1 * Math.Round(kmoff1, 2))
    '                            ElseIf kmoff1 > 0 Then
    '                                offStr = "Право " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            End If
    '                        End If
    '                        tabl.Cell(countTableword, 4).Range.Text = km & vbLf & offStr
    '                        'ПК конца
    '                        km = "-"
    '                        If IsNumeric(ArrayDataWord(7, i)) = True Then
    '                            km = FuncFormatKM(Val(ArrayDataWord(7, i)))
    '                        Else
    '                            If ArrayDataWord(5, i).Trim.Length > 0 Then
    '                                km = ArrayDataWord(7, i)
    '                            End If
    '                        End If
    '                        offStr = ""
    '                        kmoff1 = 0
    '                        If IsNumeric(ArrayDataWord(8, i)) = True Then
    '                            If kmoff1 < 0 Then
    '                                offStr = "Лево " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            ElseIf kmoff1 > 0 Then
    '                                offStr = "Право " & String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                            End If
    '                        End If
    '                        tabl.Cell(countTableword, 5).Range.Text = km & vbLf & offStr

    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(9, i) 'длина
    '                        tabl.Cell(countTableword, 7).Range.Text = ArrayDataWord(10, i) 'расположение относительно дороги
    '                        tabl.Cell(countTableword, 8).Range.Text = ArrayDataWord(13, i) 'материал
    '                        If IsNumeric(ArrayDataWord(11, i)) = True Then
    '                            tabl.Cell(countTableword, 9).Range.Text = String.Format("{0:0.0}", ArrayDataWord(11, i))
    '                        Else
    '                            tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(11, i) 'высота
    '                        End If
    '                        tabl.Cell(countTableword, 10).Range.Text = ArrayDataWord(12, i) 'описание
    '                        Exit For
    '                    Next
    '                Next i
    '                WApp.Visible = True
    '            End If
    '        End Using
    '        locker.Dispose()
    '        WDoc = Nothing
    '        WApp = Nothing
    '    End Sub
    '    '==================================================================================================
    '    'Ведомости автобусных остановок
    '    <CommandMethod("STPR-DefVedOst")>
    '    Public Sub DefVedOst()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayPointWord As String(,) = Nothing
    '        Dim CountArrayPointWord As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость остановок общественного транспорта.docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость остановок общественного транспорта.docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'забираем в массив все Характерные линии с расширенными данными Road_Barrier
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'выбираем объекты
    '            Dim NameBlk1 As String = "A-Остановки маршрутных транспортных средств необорудованные"
    '            Dim NameBlk2 As String = "A-Остановки маршрутных транспортных средств оборудованные"
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        'фильтруем блоки автокад
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf acEnt Is BlockReference Then
    '                            Dim blockDyn As BlockTableRecord = Nothing
    '                            Dim obj As BlockReference = acEnt
    '                            Dim TempName As String = obj.Name
    '                            If obj.IsDynamicBlock = True Then
    '                                Dim blockRef As BlockReference = acTrans.GetObject(acEnt.ObjectId, OpenMode.ForRead)
    '                                blockDyn = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
    '                                TempName = blockDyn.Name
    '                            End If
    '                            If TempName Like NameBlk1 Or TempName Like NameBlk2 Then
    '                                Dim PtXY As Point3d = obj.Position
    '                                'удаляем предыдущие созданные точки
    '                                Dim NumbPoint As Integer() = Nothing
    '                                Dim boolIns As Boolean = FuncFindCogoPoint(PtXY.Convert2d(New Plane), NumbPoint, 0.02, "Авт. остановка")
    '                                If IsArray(NumbPoint) = True Then
    '                                    For k As Integer = 0 To NumbPoint.Length - 1
    '                                        Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '                                        Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '                                        If IsDBNull(PointCogoDel) = False Then
    '                                            PointCogoDel.Erase()
    '                                        End If
    '                                    Next k
    '                                End If
    '                                '===================================================================================
    '                                'создаем COGO - точку
    '                                'генерируем номер для cogo точки
    '                                Dim numbCogo As Long = 0
    '                                For i As Integer = 0 To 100
    '                                    Dim rng As Random = New Random
    '                                    numbCogo = rng.Next(1000500, 1000999)
    '                                    If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '                                        Exit For
    '                                    End If
    '                                Next
    '                                'проециреем нужную высоту
    '                                '========================================================================================
    '                                'находим поверхность рабочую
    '                                Dim elev As Double = PtXY.Z
    '                                Dim ColSurv As ObjectIdCollection = CivilDoc.GetSurfaceIds
    '                                For Each idSurf As ObjectId In ColSurv
    '                                    Dim surf As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(idSurf, OpenMode.ForRead)
    '                                    Try
    '                                        elev = surf.FindElevationAtXY(PtXY.X, PtXY.Y)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                Next
    '                                Dim idCogo As ObjectId = FuncCreateCOGOPoint(PtXY, "Авт. остановка", numbCogo, elev)
    '                                Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '                                'присваиваем набор данных
    '                                Dim ArrayODRec As String(,) = Nothing
    '                                ReDim Preserve ArrayODRec(2, 1)
    '                                ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '                                ArrayODRec(1, 0) = "LABEL"
    '                                ArrayODRec(2, 0) = "Остановка маршрутных транспортных средств"
    '                                ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '                                ArrayODRec(1, 1) = "DESK"
    '                                ArrayODRec(2, 1) = "Авт. остановка"
    '                                If IsNothing(CreateTBLPSID) = False Then
    '                                    'присваиваем таблицу объекту
    '                                    Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '                                    'записываем данные в таблицу
    '                                    Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '                                End If
    '                                'считываем расширенные данные
    '                                'Dim ArrayData As String(,) = Nothing
    '                                Erase ArrayData
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(obj, ArrayData, "ROAD_SITES")
    '                                If IsArray(ArrayData) = True Then
    '                                    '0-км для сортировки, 1-КМ+,2-3-смещение, 4-ПК+, 5-6-смещение, 7-функциональное назначене, 8-Посадочная площадка, 9-Павильон, 10-Карман, 11-ПСП, 12-Покрытие, 13-Примечание
    '                                    ReDim Preserve ArrayPointWord(13, CountArrayPointWord)
    '                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                        Dim NameField As String = ArrayData(1, i)
    '                                        Dim val As String = ArrayData(2, i)
    '                                        If IsNothing(val) = True Then val = ""
    '                                        If NameField Like "FUNCTION" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(7, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(7, CountArrayPointWord) = "Остановка маршрутных транспортных средств"
    '                                            End If
    '                                        ElseIf NameField Like "LAND_SITES" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(8, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(8, CountArrayPointWord) = "Нет"
    '                                            End If
    '                                        ElseIf NameField Like "PAVILION" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(9, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(9, CountArrayPointWord) = "Нет"
    '                                            End If
    '                                        ElseIf NameField Like "ZD_KARMAN" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(10, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(10, CountArrayPointWord) = "Нет"
    '                                            End If
    '                                        ElseIf NameField Like "POLOSA" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(11, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(11, CountArrayPointWord) = "Нет"
    '                                            End If
    '                                        ElseIf NameField Like "PAVING" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(12, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(12, CountArrayPointWord) = "-"
    '                                            End If
    '                                        ElseIf NameField Like "NOTE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(13, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(13, CountArrayPointWord) = "-"
    '                                            End If
    '                                        End If
    '                                    Next i
    '                                End If
    '                                '=================================================================================
    '                                'определяем километр
    '                                'записываем существующий километраж
    '                                If IsNothing(AlignEX) = False Then
    '                                    Dim KMPK As Double = -1
    '                                    Dim KMoff As Double = -1
    '                                    Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(AlignEX, PtXY, KMPK, KMoff, 2)
    '                                    If KMoff <> -1 Then
    '                                        Dim km As String = FuncFormatKM(KMPK)
    '                                        ArrayPointWord(1, CountArrayPointWord) = km
    '                                        Dim offStr As String = ""
    '                                        If KMoff < 0 Then
    '                                            offStr = Math.Round(Math.Abs(KMoff), 2)
    '                                            ArrayPointWord(2, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                            ArrayPointWord(3, CountArrayPointWord) = "-"
    '                                        ElseIf KMoff > 0 Then
    '                                            offStr = Math.Round(KMoff, 2)
    '                                            ArrayPointWord(3, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                            ArrayPointWord(2, CountArrayPointWord) = "-"
    '                                        End If
    '                                        ArrayPointWord(0, CountArrayPointWord) = KMPK
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        ArrayPointWord(0, CountArrayPointWord) = KMPK
    '                                        ArrayPointWord(1, CountArrayPointWord) = "-"
    '                                        ArrayPointWord(2, CountArrayPointWord) = "-"
    '                                        ArrayPointWord(3, CountArrayPointWord) = "-"
    '                                    End If
    '                                End If
    '                                'записываем пикетаж проектный
    '                                If IsNothing(AlignProject) = False Then
    '                                    Dim PK As Double = -1
    '                                    Dim off As Double = -1
    '                                    Dim boolPK As Boolean = FuncFindPKOffsetToAlign(AlignProject, PtXY, PK, off, 2)
    '                                    If off <> -1 Then
    '                                        Try
    '                                            Dim pkStrt As String = AlignProject.GetStationStringWithEquations(PK)
    '                                            ArrayPointWord(4, CountArrayPointWord) = pkStrt
    '                                            Dim offStr As String = ""
    '                                            If off < 0 Then
    '                                                offStr = Math.Round(Math.Abs(off), 2)
    '                                                ArrayPointWord(5, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                                ArrayPointWord(6, CountArrayPointWord) = "-"
    '                                            ElseIf off > 0 Then
    '                                                offStr = Math.Round(off, 2)
    '                                                ArrayPointWord(6, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                                ArrayPointWord(5, CountArrayPointWord) = "-"
    '                                            End If
    '                                            If IsNothing(AlignEX) = True Then
    '                                                ArrayPointWord(0, CountArrayPointWord) = PK
    '                                            End If
    '                                        Catch ex As System.ArgumentException
    '                                        End Try
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        If IsNothing(AlignEX) = True Then
    '                                            ArrayPointWord(0, CountArrayPointWord) = PK
    '                                        End If
    '                                        ArrayPointWord(4, CountArrayPointWord) = "-"
    '                                        ArrayPointWord(5, CountArrayPointWord) = "-"
    '                                        ArrayPointWord(6, CountArrayPointWord) = "-"
    '                                    End If
    '                                    CountArrayPointWord += 1
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            'делаем сортировку массива по пикетажу
    '            If IsNothing(ArrayPointWord) = False Then
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость остановок общественного транспорта.docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость остановок общественного транспорта_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                    GoTo Line1
    '                End Try
    '                If ArrayPointWord.GetUpperBound(1) > 0 Then
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                Else
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                End If
    '                '=========================================================================================
    '                'заполняем таблицу
    '                '0-км для сортировки, 1-КМ+,2-3-смещение, 4-ПК+, 5-6-смещение, 7-функциональное назначене, 8-Посадочная площадка, 9-Павильон, 10-Карман, 11-ПСП, 12-Покрытие, 13-Примечание
    '                For i As Integer = 0 To ArrayPointWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        tabl.Rows.Add()
    '                        If IsNothing(ArrayPointWord(7, i)) = False Then
    '                            If ArrayPointWord(7, i).Trim.Length > 0 Then
    '                                tabl.Cell(countTableword, 1).Range.Text = ArrayPointWord(7, i) 'функциональное назначение
    '                            Else
    '                                tabl.Cell(countTableword, 1).Range.Text = "Остановка маршрутных транспортных средств"
    '                            End If
    '                        Else
    '                            tabl.Cell(countTableword, 1).Range.Text = "Остановка маршрутных транспортных средств"
    '                        End If
    '                        'километр
    '                        tabl.Cell(countTableword, 2).Range.Text = ArrayPointWord(1, i)
    '                        tabl.Cell(countTableword, 3).Range.Text = ArrayPointWord(2, i)
    '                        tabl.Cell(countTableword, 4).Range.Text = ArrayPointWord(3, i)
    '                        'пикет
    '                        tabl.Cell(countTableword, 5).Range.Text = ArrayPointWord(4, i)
    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayPointWord(5, i)
    '                        tabl.Cell(countTableword, 7).Range.Text = ArrayPointWord(6, i)
    '                        'посадочная площадка
    '                        tabl.Cell(countTableword, 8).Range.Text = ArrayPointWord(8, i)
    '                        'павильон
    '                        tabl.Cell(countTableword, 9).Range.Text = ArrayPointWord(9, i)
    '                        'карман
    '                        tabl.Cell(countTableword, 10).Range.Text = ArrayPointWord(10, i)
    '                        'ПСП
    '                        tabl.Cell(countTableword, 11).Range.Text = ArrayPointWord(11, i)
    '                        'Покрытие
    '                        tabl.Cell(countTableword, 12).Range.Text = ArrayPointWord(12, i)
    '                        'Примечани
    '                        tabl.Cell(countTableword, 13).Range.Text = ArrayPointWord(13, i)
    '                        Exit For
    '                    Next
    '                Next i
    '            End If
    '            acTrans.Commit()
    '        End Using
    '        WApp.Visible = True
    'Line1:
    '        locker.Dispose()

    '        WDoc = Nothing
    '        WApp = Nothing
    '    End Sub
    '    '==================================================================================================
    '    'формирование ведомости пешеходных переходов
    '    <CommandMethod("STPR-VedCross")>
    '    Public Sub VedCross()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayPointWord As String(,) = Nothing
    '        Dim CountArrayPointWord As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость пешеходных переходов.docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость пешеходных переходов.docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\Topografiya\MapClassification\"
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'забираем в массив все Характерные линии с расширенными данными Road_Barrier
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'выбираем объекты
    '            Dim NameBlk1 As String = "A-Разметка пешеходного перехода"
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        'фильтруем блоки автокад
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf acEnt Is BlockReference Then
    '                            Dim blockDyn As BlockTableRecord = Nothing
    '                            Dim obj As BlockReference = acEnt
    '                            Dim TempName As String = obj.Name
    '                            Dim LCross As Double = 0
    '                            If obj.IsDynamicBlock = True Then
    '                                Dim blockRef As BlockReference = acTrans.GetObject(acEnt.ObjectId, OpenMode.ForRead)
    '                                blockDyn = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
    '                                TempName = blockDyn.Name
    '                                'считаем длину, вычленаем два параметра динамического блока
    '                                Dim blkColl As DynamicBlockReferencePropertyCollection = blockRef.DynamicBlockReferencePropertyCollection
    '                                If blkColl.Count > 0 Then
    '                                    For Each blkref As DynamicBlockReferenceProperty In blkColl
    '                                        Dim NameF As String = blkref.PropertyName
    '                                        If NameF Like "Расстояние1" Then
    '                                            LCross = LCross + Val(blkref.Value)
    '                                        End If
    '                                        If NameF Like "Расстояние2" Then
    '                                            LCross = LCross + Val(blkref.Value)
    '                                        End If
    '                                    Next
    '                                End If
    '                            End If
    '                            If TempName Like NameBlk1 Then
    '                                Dim PtXY As Point3d = obj.Position
    '                                Erase ArrayData
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(obj, ArrayData, "CROSSWALK")
    '                                ReDim Preserve ArrayPointWord(8, CountArrayPointWord)
    '                                If IsArray(ArrayData) = True Then
    '                                    '0-км для сортировки, 1-КМ+, 2-ПК+, 3-4-смещение, 5-тип перехода, 6-Вид, 7-Длина, 8-Примечание
    '                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                        Dim NameField As String = ArrayData(1, i)
    '                                        Dim val As String = ArrayData(2, i)
    '                                        If IsNothing(val) = True Then val = ""
    '                                        If NameField Like "TYPE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(5, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(5, CountArrayPointWord) = "-"
    '                                            End If
    '                                        ElseIf NameField Like "LENGHT" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(7, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(7, CountArrayPointWord) = "-"
    '                                            End If
    '                                        ElseIf NameField Like "TRAFFICLIGHT" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(6, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(6, CountArrayPointWord) = "-"
    '                                            End If
    '                                        ElseIf NameField Like "NOTE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(8, CountArrayPointWord) = val
    '                                            Else
    '                                                ArrayPointWord(8, CountArrayPointWord) = "-"
    '                                            End If
    '                                        End If
    '                                    Next i
    '                                End If
    '                                If LCross > 0 Then
    '                                    ArrayPointWord(7, CountArrayPointWord) = Math.Round(LCross, 2)
    '                                End If
    '                                '=================================================================================
    '                                'определяем километр
    '                                'записываем существующий километраж
    '                                If IsNothing(AlignEX) = False Then
    '                                    Dim KMPK As Double = -1
    '                                    Dim KMoff As Double = -1
    '                                    Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(AlignEX, PtXY, KMPK, KMoff, 2)
    '                                    If KMoff <> -1 Then
    '                                        Dim km As String = FuncFormatKM(KMPK)
    '                                        ArrayPointWord(1, CountArrayPointWord) = km
    '                                        ArrayPointWord(0, CountArrayPointWord) = KMPK
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        Dim d1 As Double = AlignEX.StartPoint.Convert2d(New Plane).GetVectorTo(PtXY.Convert2d(New Plane)).Length
    '                                        Dim d2 As Double = AlignEX.StartPoint.Convert2d(New Plane).GetVectorTo(PtXY.Convert2d(New Plane)).Length
    '                                        If d1 <= d2 Then
    '                                            ArrayPointWord(0, CountArrayPointWord) = -99999
    '                                        Else
    '                                            ArrayPointWord(0, CountArrayPointWord) = 99999
    '                                        End If
    '                                    End If
    '                                End If
    '                                'записываем пикетаж проектный
    '                                If IsNothing(AlignProject) = False Then
    '                                    Dim PK As Double = -1
    '                                    Dim off As Double = -1
    '                                    Dim boolPK As Boolean = FuncFindPKOffsetToAlign(AlignProject, PtXY, PK, off, 2)
    '                                    If off <> -1 Then
    '                                        Try
    '                                            Dim pkStrt As String = AlignProject.GetStationStringWithEquations(PK)
    '                                            ArrayPointWord(2, CountArrayPointWord) = pkStrt
    '                                            Dim offStr As String = ""
    '                                            If off < 0 Then
    '                                                offStr = Math.Round(off, 2)
    '                                                ArrayPointWord(3, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                                ArrayPointWord(4, CountArrayPointWord) = "-"
    '                                            ElseIf off > 0 Then
    '                                                offStr = Math.Round(off, 2)
    '                                                ArrayPointWord(4, CountArrayPointWord) = String.Format("{0:0.00}", offStr)
    '                                                ArrayPointWord(3, CountArrayPointWord) = "-"
    '                                            End If
    '                                            If IsNothing(AlignEX) = True Then
    '                                                ArrayPointWord(0, CountArrayPointWord) = PK
    '                                            End If
    '                                        Catch ex As System.ArgumentException
    '                                        End Try
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        If IsNothing(AlignEX) = True Then
    '                                            Dim d1 As Double = AlignProject.StartPoint.Convert2d(New Plane).GetVectorTo(PtXY.Convert2d(New Plane)).Length
    '                                            Dim d2 As Double = AlignProject.EndPoint.Convert2d(New Plane).GetVectorTo(PtXY.Convert2d(New Plane)).Length
    '                                            If d1 <= d2 Then
    '                                                ArrayPointWord(0, CountArrayPointWord) = -99999
    '                                            Else
    '                                                ArrayPointWord(0, CountArrayPointWord) = 99999
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    CountArrayPointWord += 1
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If

    '            'делаем сортировку массива по пикетажу
    '            If IsNothing(ArrayPointWord) = False Then
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пешеходных переходов.docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость пешеходных переходов_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                    GoTo Line1
    '                End Try
    '                If ArrayPointWord.GetUpperBound(1) > 0 Then
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                Else
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                End If
    '                '=========================================================================================
    '                'заполняем таблицу
    '                '0-км для сортировки, 1-КМ+,2-3-смещение, 4-ПК+, 5-6-смещение, 7-функциональное назначене, 8-Посадочная площадка, 9-Павильон, 10-Карман, 11-ПСП, 12-Покрытие, 13-Примечание
    '                For i As Integer = 0 To ArrayPointWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        tabl.Rows.Add()
    '                        tabl.Cell(countTableword, 1).Range.Text = "Пешеходный переход" 'наименование
    '                        'километр
    '                        tabl.Cell(countTableword, 2).Range.Text = ArrayPointWord(1, i)
    '                        'пикет
    '                        tabl.Cell(countTableword, 3).Range.Text = ArrayPointWord(2, i)
    '                        tabl.Cell(countTableword, 4).Range.Text = ArrayPointWord(3, i)
    '                        tabl.Cell(countTableword, 5).Range.Text = ArrayPointWord(4, i)
    '                        'тип перехода
    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayPointWord(5, i)
    '                        'вид перехода
    '                        tabl.Cell(countTableword, 7).Range.Text = ArrayPointWord(6, i)
    '                        'протяженность
    '                        tabl.Cell(countTableword, 8).Range.Text = ArrayPointWord(7, i)
    '                        'примечание
    '                        tabl.Cell(countTableword, 9).Range.Text = ArrayPointWord(8, i)
    '                        Exit For
    '                    Next
    '                Next i
    '            End If
    '            acTrans.Commit()
    '        End Using
    '        WApp.Visible = True
    'Line1:
    '        locker.Dispose()
    '        WDoc = Nothing
    '        WApp = Nothing
    '    End Sub
    '    '==================================================================================================
    '    'формирование ведомости пересечений и примыканий автомобильных дорог
    '    <CommandMethod("STPR-VedRoadWays")>
    '    Public Sub VedRoadWays()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayPointWord As String(,) = Nothing
    '        Dim CountArrayPointWord As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость пересечений с автомобильными и железными дорогами.docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '            'Dim rng As Random = New Random
    '            'Dim Numbrng As Integer = rng.Next
    '            ' WDoc.SaveAs2(TempDir & "\Ведомость пересечений с автомобильными и железными дорогами_" & Numbrng & ".docx")
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость пересечений с автомобильными и железными дорогами.docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\Topografiya\MapClassification\"
    '        If Directory.Exists(PutchMapClass) = False Then
    '            PutchMapClass = parentPatch2 & "\Topografiya\MapClassification\"
    '        End If
    '        'создаем таблицу с данными PS
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'блокируем документ
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'забираем в массив все блоки с именем  A-Указатель направления автодорог или любой другой блок с таблицей данных ROADWAYS
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            'выбираем объекты
    '            Dim NameBlk1 As String = "A-Указатель направления автодорог"
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        'фильтруем блоки автокад
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        Dim boolFindObj As Boolean = False
    '                        Erase ArrayData
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf acEnt Is BlockReference Then
    '                            Dim acBlk As BlockReference = acEnt
    '                            Dim blockDyn As BlockTableRecord = Nothing
    '                            Dim obj As BlockReference = acEnt
    '                            Dim TempName As String = obj.Name
    '                            If obj.IsDynamicBlock = True Then
    '                                Dim blockRef As BlockReference = acTrans.GetObject(acEnt.ObjectId, OpenMode.ForRead)
    '                                blockDyn = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
    '                                TempName = blockDyn.Name
    '                            End If
    '                            'если имя блока A-Указатель направления автодорог
    '                            If TempName Like NameBlk1 Then
    '                                boolFindObj = True
    '                                'смотрим таблицу
    '                                Dim boolFindObjTBL As Boolean = FuncFindPSetTable(obj, "ROADS_WAYS")
    '                                If boolFindObjTBL = True Then
    '                                    Dim boolReadPS As Boolean = FuncReadDataSetProperties(obj, ArrayData, "ROADS_WAYS")
    '                                Else
    '                                    boolFindObjTBL = FuncFindPSetTable(obj, "RAILROADS")
    '                                    If boolFindObj = True Then
    '                                        Dim boolReadPS As Boolean = FuncReadDataSetProperties(obj, ArrayData, "RAILROADS")
    '                                    End If
    '                                End If
    '                            End If
    '                            '==================================================================================
    '                            'начальные установки
    '                            Dim PtXY As Point3d = Nothing
    '                            Dim angle As Double = Math.PI / 2
    '                            Dim CogoDesk As String = FuncReadValueAttr(obj, "NAME")
    '                            If CogoDesk.Trim.Length = 0 Then
    '                                CogoDesk = TempName
    '                            End If
    '                            If boolFindObj = True Then
    '                                'определяем угол пересечения
    '                                angle = acBlk.Rotation
    '                                'определяем угол из динамического блока
    '                                Dim blkColl As DynamicBlockReferencePropertyCollection = acBlk.DynamicBlockReferencePropertyCollection
    '                                If blkColl.Count > 0 Then
    '                                    For Each blkref As DynamicBlockReferenceProperty In blkColl
    '                                        Dim NameF As String = blkref.PropertyName
    '                                        Try
    '                                            If NameF.IndexOf("Угол") > -1 Then
    '                                                angle = angle + CDbl(blkref.Value)
    '                                            End If
    '                                        Catch ex As Exception
    '                                        End Try
    '                                    Next
    '                                End If
    '                                angle = Math.Abs(angle)
    '                                If angle = 0 Then
    '                                    angle = Math.PI / 2
    '                                Else
    '                                    Select Case angle
    '                                        Case Math.PI / 2 To Math.PI
    '                                            angle = Math.PI - angle
    '                                        Case Math.PI To (3 * Math.PI) / 2
    '                                            angle = (3 * Math.PI) / 2 - angle
    '                                        Case (3 * Math.PI) / 2 To Math.PI * 2
    '                                            angle = Math.PI * 2 - angle
    '                                    End Select
    '                                End If
    '                                'переводим в градусы
    '                                angle = Math.Abs(90 - (180 * angle) / Math.PI)
    '                            End If
    '                            '=========================================================================================
    '                            'основное тело функции
    '                            If boolFindObj = True Then
    '                                Dim NameRoad As String = ""
    '                                PtXY = obj.Position 'точка вставки блока
    '                                'считываем расширенные данные
    '                                '0-км для сортировки, 1-КМ+,2-3-смещение, 4-ПК+, 5-6-смещение, 7-Наименование дороги, 8-Угол пересечения, 9-Примечание, 10-Тип объекта, 11-Тип покрытия, 12-Километр пересечения
    '                                ReDim Preserve ArrayPointWord(12, CountArrayPointWord)
    '                                ArrayPointWord(7, CountArrayPointWord) = NameRoad
    '                                ArrayPointWord(10, CountArrayPointWord) = "Автомобильная дорога"
    '                                If IsArray(ArrayData) = True Then
    '                                    Dim NoteSumm As String = ""
    '                                    Dim Name As String = NameRoad
    '                                    Dim CategoryRoad As String = ""
    '                                    Dim idNumber As String = ""
    '                                    Dim Paving As String = ""
    '                                    Dim HLine As String = ""
    '                                    Dim CountLine As String = ""
    '                                    Dim KMIntersect As String = ""
    '                                    Dim State As String = ""
    '                                    Dim Owner As String = ""
    '                                    Dim Note As String = ""
    '                                    Dim PowerStr As String = ""
    '                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                        Dim NameField As String = ArrayData(1, i)
    '                                        Dim val As String = ArrayData(2, i)
    '                                        If IsNothing(val) = True Then val = ""
    '                                        If NameField Like "CATEGORY" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                CategoryRoad = val
    '                                            End If
    '                                        ElseIf NameField Like "LABEL" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(10, CountArrayPointWord) = val
    '                                            End If
    '                                        ElseIf NameField Like "NAME" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                CogoDesk = val
    '                                            End If
    '                                        ElseIf NameField Like "IDNUMBER" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                idNumber = val
    '                                            End If
    '                                        ElseIf NameField Like "PAVING" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                ArrayPointWord(11, CountArrayPointWord) = val
    '                                            End If
    '                                        ElseIf NameField Like "HEIGHTLINE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                HLine = val
    '                                            End If
    '                                        ElseIf NameField Like "COUNTLINES" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                CountLine = val
    '                                            End If
    '                                        ElseIf NameField Like "KM_INTERSECT" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                KMIntersect = val
    '                                            End If
    '                                        ElseIf NameField Like "OWNER" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                Owner = val
    '                                            End If
    '                                        ElseIf NameField Like "STATE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                State = val
    '                                            End If
    '                                        ElseIf NameField Like "NOTE" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                Note = val
    '                                            End If
    '                                        ElseIf NameField Like "POWER" Then
    '                                            If val.Trim.Length > 0 Then
    '                                                PowerStr = val
    '                                            End If
    '                                        End If
    '                                    Next i
    '                                    'идентификатор дороги
    '                                    If idNumber.Trim.Length > 0 Then
    '                                        If CogoDesk.Trim.Length > 0 Then
    '                                            NameRoad = CogoDesk & vbLf & idNumber
    '                                        Else
    '                                            NameRoad = idNumber
    '                                        End If
    '                                    Else
    '                                        NameRoad = CogoDesk
    '                                    End If

    '                                    'километр пересечения (примыкания)
    '                                    If KMIntersect.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            ArrayPointWord(12, CountArrayPointWord) = "Км. - " & KMIntersect
    '                                        End If
    '                                    End If
    '                                    'графа примечание
    '                                    If HLine.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & "Ширина полос движения - " & HLine
    '                                        Else
    '                                            NoteSumm = "Ширина полос движения - " & HLine
    '                                        End If
    '                                    End If
    '                                    If CountLine.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & "Число полос движения - " & CountLine
    '                                        Else
    '                                            NoteSumm = "Число полос движения - " & CountLine
    '                                        End If
    '                                    End If
    '                                    If Paving.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & "Тип покрытия - " & Paving
    '                                        Else
    '                                            NoteSumm = "Тип покрытия - " & Paving
    '                                        End If
    '                                    End If
    '                                    If Owner.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & "Владелец - " & Owner
    '                                        Else
    '                                            NoteSumm = "Владелец - " & Owner
    '                                        End If
    '                                    End If
    '                                    If State.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & "Состояние - " & State
    '                                        Else
    '                                            NoteSumm = "Состояние - " & State
    '                                        End If
    '                                    End If
    '                                    If PowerStr.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & PowerStr
    '                                        Else
    '                                            NoteSumm = PowerStr
    '                                        End If
    '                                    End If
    '                                    If Note.Trim.Length > 0 Then
    '                                        If NoteSumm.Trim.Length > 0 Then
    '                                            NoteSumm = NoteSumm & vbLf & Note
    '                                        Else
    '                                            NoteSumm = Note
    '                                        End If
    '                                    End If
    '                                    If NoteSumm.Trim.Length = 0 Then
    '                                        ArrayPointWord(9, CountArrayPointWord) = NoteSumm
    '                                    End If
    '                                End If
    '                                If NameRoad.Trim.Length = 0 Then
    '                                    If CogoDesk.Trim.Length = 0 Then
    '                                        NameRoad = "-"
    '                                    Else
    '                                        NameRoad = CogoDesk
    '                                    End If
    '                                End If
    '                                ArrayPointWord(7, CountArrayPointWord) = NameRoad
    '                                '=================================================================================
    '                                'определяем километр
    '                                'записываем существующий километраж
    '                                '0-пересечение, 1-лево, 2 - право
    '                                Dim TypePer As Integer = -1
    '                                If IsNothing(AlignEX) = False Then
    '                                    Dim KMPK As Double = -1
    '                                    Dim KMoff As Double = -1
    '                                    Dim PkAlign As Boolean = FuncFindPKOffsetToAlign(AlignEX, PtXY, KMPK, KMoff, 2)
    '                                    If KMoff <> -1 Then
    '                                        Dim km As String = FuncFormatKM(KMPK)
    '                                        ArrayPointWord(1, CountArrayPointWord) = km
    '                                        If Math.Abs(KMoff) < 1 Then
    '                                            ArrayPointWord(2, CountArrayPointWord) = "-"
    '                                            ArrayPointWord(3, CountArrayPointWord) = "-"
    '                                            If IsNothing(ArrayPointWord(12, CountArrayPointWord)) = True Then
    '                                                ArrayPointWord(12, CountArrayPointWord) = "Пересечение"
    '                                            Else
    '                                                If ArrayPointWord(12, CountArrayPointWord).Trim.Length > 0 Then
    '                                                    Dim tempStrKm As String = ArrayPointWord(12, CountArrayPointWord)
    '                                                    ArrayPointWord(12, CountArrayPointWord) = "Пересечение" & tempStrKm
    '                                                Else
    '                                                    ArrayPointWord(12, CountArrayPointWord) = "Пересечение"
    '                                                End If
    '                                            End If
    '                                            TypePer = 0
    '                                        Else
    '                                            If KMoff < 0 Then
    '                                                KMoff = Math.Round(KMoff, 2)
    '                                                ArrayPointWord(2, CountArrayPointWord) = String.Format("{0:0.00}", Math.Abs(KMoff))
    '                                                ArrayPointWord(3, CountArrayPointWord) = "-"
    '                                                TypePer = 1
    '                                            ElseIf KMoff > 0 Then
    '                                                KMoff = Math.Round(KMoff, 2)
    '                                                ArrayPointWord(3, CountArrayPointWord) = String.Format("{0:0.00}", Math.Abs(KMoff))
    '                                                ArrayPointWord(2, CountArrayPointWord) = "-"
    '                                                TypePer = 2
    '                                            End If
    '                                            If IsNothing(ArrayPointWord(12, CountArrayPointWord)) = True Then
    '                                                ArrayPointWord(12, CountArrayPointWord) = "Примыкание"
    '                                            Else
    '                                                If ArrayPointWord(12, CountArrayPointWord).Trim.Length > 0 Then
    '                                                    Dim tempStrKm As String = ArrayPointWord(12, CountArrayPointWord)
    '                                                    ArrayPointWord(12, CountArrayPointWord) = "Примыкание" & tempStrKm
    '                                                Else
    '                                                    ArrayPointWord(12, CountArrayPointWord) = "Примыкание"
    '                                                End If
    '                                            End If
    '                                            ArrayPointWord(0, CountArrayPointWord) = KMPK
    '                                        End If
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        ArrayPointWord(0, CountArrayPointWord) = KMPK
    '                                    End If
    '                                End If
    '                                'записываем пикетаж проектный9
    '                                If IsNothing(AlignProject) = False Then
    '                                    Dim PK As Double = -1
    '                                    Dim off As Double = -1
    '                                    Dim boolPK As Boolean = FuncFindPKOffsetToAlign(AlignProject, PtXY, PK, off, 2)
    '                                    If off <> -1 Then
    '                                        Try
    '                                            Dim pkStrt As String = AlignProject.GetStationStringWithEquations(PK)
    '                                            ArrayPointWord(4, CountArrayPointWord) = pkStrt
    '                                            If off < 0 Then
    '                                                off = Math.Round(off, 2)
    '                                                ArrayPointWord(5, CountArrayPointWord) = String.Format("{0:0.00}", Math.Abs(off))
    '                                                ArrayPointWord(6, CountArrayPointWord) = "-"
    '                                                If TypePer = -1 Then
    '                                                    TypePer = 1
    '                                                End If
    '                                            ElseIf off > 0 Then
    '                                                off = Math.Round(off, 2)
    '                                                ArrayPointWord(6, CountArrayPointWord) = String.Format("{0:0.00}", Math.Abs(off))
    '                                                ArrayPointWord(5, CountArrayPointWord) = "-"
    '                                                If TypePer = -1 Then
    '                                                    TypePer = 2
    '                                                End If
    '                                            End If
    '                                            If IsNothing(AlignEX) = True Then
    '                                                ArrayPointWord(0, CountArrayPointWord) = PK
    '                                            End If
    '                                        Catch ex As System.ArgumentException
    '                                        End Try
    '                                    Else
    '                                        'точка за пределами трассы
    '                                        If IsNothing(AlignEX) = True Then
    '                                            ArrayPointWord(0, CountArrayPointWord) = PK
    '                                        End If
    '                                    End If
    '                                    ArrayPointWord(8, CountArrayPointWord) = Math.Round(angle, 0)
    '                                    CountArrayPointWord += 1
    '                                End If
    '                                '========================================================================================================
    '                                'вставляем точку cogo
    '                                'удаляем предыдущие созданные точки
    '                                Dim NumbPoint As Integer() = Nothing
    '                                If TypePer = 0 Then
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(PtXY.Convert2d(New Plane), NumbPoint, 0.02, "Пересечение дорог")
    '                                ElseIf TypePer = 1 Then
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(PtXY.Convert2d(New Plane), NumbPoint, 0.02, "Примыкание дорог (лево)")
    '                                ElseIf TypePer = 2 Then
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(PtXY.Convert2d(New Plane), NumbPoint, 0.02, "Примыкание дорог (право)")
    '                                Else
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(PtXY.Convert2d(New Plane), NumbPoint, 0.02, "Примыкание дорог")
    '                                End If
    '                                If IsArray(NumbPoint) = True Then
    '                                    For k As Integer = 0 To NumbPoint.Length - 1
    '                                        Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '                                        Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '                                        If IsDBNull(PointCogoDel) = False Then
    '                                            PointCogoDel.Erase()
    '                                        End If
    '                                    Next k
    '                                End If
    '                                '===================================================================================
    '                                'создаем COGO - точку
    '                                'генерируем номер для cogo точки
    '                                Dim numbCogo As Long = 0
    '                                For i As Integer = 0 To 100
    '                                    Dim rng As Random = New Random
    '                                    numbCogo = rng.Next(1000500, 1000999)
    '                                    If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '                                        Exit For
    '                                    End If
    '                                Next
    '                                'проециреем нужную высоту
    '                                '========================================================================================
    '                                'находим поверхность рабочую
    '                                Dim elev As Double = PtXY.Z
    '                                Dim ColSurv As ObjectIdCollection = CivilDoc.GetSurfaceIds
    '                                For Each idSurf As ObjectId In ColSurv
    '                                    Dim surf As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(idSurf, OpenMode.ForRead)
    '                                    Try
    '                                        elev = surf.FindElevationAtXY(PtXY.X, PtXY.Y)
    '                                        Exit For
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                Next
    '                                'создаем новую точку
    '                                Dim idCogo As ObjectId = Nothing
    '                                If TypePer = 0 Then
    '                                    idCogo = FuncCreateCOGOPoint(PtXY, CogoDesk, numbCogo, elev,, "Пересечение дорог")
    '                                ElseIf TypePer = 1 Then
    '                                    idCogo = FuncCreateCOGOPoint(PtXY, CogoDesk, numbCogo, elev, , "Примыкание дорог (лево)")
    '                                ElseIf TypePer = 2 Then
    '                                    idCogo = FuncCreateCOGOPoint(PtXY, CogoDesk, numbCogo, elev,, "Примыкание дорог (право)")
    '                                Else
    '                                    idCogo = FuncCreateCOGOPoint(PtXY, CogoDesk, numbCogo, elev,, "Примыкание дорог")
    '                                End If
    '                                Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '                                'присваиваем набор данных
    '                                Dim ArrayODRec As String(,) = Nothing
    '                                ReDim Preserve ArrayODRec(2, 1)
    '                                ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '                                ArrayODRec(1, 0) = "LABEL"
    '                                ArrayODRec(2, 0) = "Указатель направления автодорог"
    '                                ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '                                ArrayODRec(1, 1) = "DESK"
    '                                ArrayODRec(2, 1) = CogoDesk
    '                                If IsNothing(CreateTBLPSID) = False Then
    '                                    'присваиваем таблицу объекту
    '                                    Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '                                    'записываем данные в таблицу
    '                                    Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            'делаем сортировку массива по пикетажу
    '            If IsNothing(ArrayPointWord) = False Then
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пересечений с автомобильными и железными дорогами.docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость пересечений с автомобильными и железными дорогами_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                    GoTo Line1
    '                End Try
    '                If ArrayPointWord.GetUpperBound(1) > 0 Then
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                Else
    '                    Dim sortArray As Boolean = FuncSortArray3(ArrayPointWord, 0)
    '                End If
    '                '=========================================================================================
    '                'заполняем таблицу
    '                '0-км для сортировки, 1-КМ+,2-3-смещение, 4-ПК+, 5-6-смещение, 7-Наименование дороги, 8-Угол пересечения, 9-Примечание, 10-Тип объекта, 11-Тип покрытия
    '                For i As Integer = 0 To ArrayPointWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        tabl.Rows.Add()
    '                        tabl.Cell(countTableword, 1).Range.Text = ArrayPointWord(10, i) 'функциональное назначение
    '                        'tabl.Cell(countTableword, 2).Range.Text = ArrayPointWord(7, i) 'функциональное назначение
    '                        'километр
    '                        tabl.Cell(countTableword, 2).Range.Text = ArrayPointWord(1, i)
    '                        tabl.Cell(countTableword, 3).Range.Text = ArrayPointWord(2, i)
    '                        tabl.Cell(countTableword, 4).Range.Text = ArrayPointWord(3, i)
    '                        'пикет
    '                        tabl.Cell(countTableword, 5).Range.Text = ArrayPointWord(4, i)
    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayPointWord(5, i)
    '                        tabl.Cell(countTableword, 7).Range.Text = ArrayPointWord(6, i)
    '                        'тип пересечения
    '                        tabl.Cell(countTableword, 8).Range.Text = ArrayPointWord(7, i)
    '                        'угол пересечения
    '                        tabl.Cell(countTableword, 10).Range.Text = ArrayPointWord(8, i)
    '                        'тип покрытия
    '                        tabl.Cell(countTableword, 9).Range.Text = ArrayPointWord(12, i)
    '                        'примечание
    '                        tabl.Cell(countTableword, 11).Range.Text = ArrayPointWord(9, i)
    '                        Exit For
    '                    Next
    '                Next i
    '            End If
    '            acTrans.Commit()
    '        End Using
    '        WApp.Visible = True
    'Line1:
    '        locker.Dispose()
    '        WDoc = Nothing
    '        WApp = Nothing
    '    End Sub
    '    '==================================================================================================
    '    'Ведомость пересечения (сближения) трубопороводов
    '    <CommandMethod("STPR-VedNetworkPipe")>
    '    Public Sub VedNetworkPipe()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayDataWord As String(,) = Nothing
    '        Dim CountArrayDataWord As Integer = 0
    '        Dim boolRez As Boolean = False
    '        Dim AngleEX As Double(,) = Nothing
    '        Dim countAngleEx As Integer = 0
    '        Dim AngleProject As Double(,) = Nothing
    '        Dim countAngleProject As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость пересечения и сближения подземных коммуникаций (трубопроводные сети).docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость пересечения и сближения подземных коммуникаций (трубопроводные сети).docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\Topografiya\MapClassification\"
    '        'Путь к файлам классификации
    '        If IO.Directory.Exists(parentPatch2 & "\Topografiya\MapClassification\") Then
    '            PutchMapClass = parentPatch2 & "\Topografiya\MapClassification\"
    '        End If
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim PlineAlignEx As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim PlineAlignProject As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '                PlineAlignProject = acTrans.GetObject(AlignProject.GetPolyline, OpenMode.ForRead)
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'тело программы
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'выбираем объекты
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim colPointEX As Point3dCollection = New Point3dCollection 'точки пересечения с существующей трассой
    '                        Dim colPointProject As Point3dCollection = New Point3dCollection 'точки пересечения с проектной трассой
    '                        'начало и конец километрового положения трубопровода
    '                        Dim km1 As Double = -1
    '                        Dim kmoff1 As Double = -1
    '                        Dim km2 As Double = -1
    '                        Dim kmoff2 As Double = -1
    '                        'начало и конец пикетажного положения трубопровода
    '                        Dim pk1 As Double = -1
    '                        Dim pkoff1 As Double = -1
    '                        Dim pk2 As Double = -1
    '                        Dim pkoff2 As Double = -1
    '                        Dim Name As String = "" 'наименование сети в общем виде например Водопровод
    '                        Dim Purpose As String = "" 'назначение сети например Речной воды или высокого давления
    '                        Dim Mat As String = "" 'материал трубы
    '                        Dim Diametr As Double = 0 'диаметр трубы
    '                        Dim Height As Double = 0 'средняя глубина заложения (высота над поверхностью земли) до верхней образующей
    '                        Dim CountPipes As Integer = 1 'число труб
    '                        Dim State As String = "" 'состояние (действуюжщий, не действующай)
    '                        Dim Owner As String = "" 'владелец
    '                        Dim Adress As String = "" 'адрес
    '                        Dim TypePipe As String = "Подземный" 'тип сети (подземный, наземный)
    '                        Dim Note As String = "" 'примечание
    '                        Dim StyleType As String = ""
    '                        Dim TopLevel As Double = 0
    '                        'чистим массивы
    '                        Erase AngleEX 'массив с углами пересечек по сущестыующей трассе
    '                        Erase AngleProject 'массив с углами пересечек по проектной трассе
    '                        Erase ArrayData 'масив с расширенными данными
    '                        countAngleEx = 0
    '                        countAngleProject = 0
    '                        boolRez = False
    '                        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
    '                            Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
    '                            If FLine.StyleName Like "ИИ-1051_Трубопроводы специальные" Or FLine.StyleName Like "ИИ-1050_Газопровод" Or FLine.StyleName Like "ИИ-1017_Канализация" Or FLine.StyleName Like "ИИ-1048_Водопровод" Or FLine.StyleName Like "ИИ-1049_Теплосеть" Then
    '                                StyleType = FLine.StyleName
    '                                boolRez = True
    '                                Dim pos As Integer = StyleType.IndexOf("_")
    '                                If pos > -1 Then
    '                                    Name = Mid(StyleType, pos)
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "PIPELINE_NETWORKS")
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                'превращаем характерныю линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                colPt = FLine.GetPoints(FeatureLinePointType.PIPoint)
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    'ищем пикетаж (существующий километр) начало участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            '3d полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "L1051" Or NameLayer Like "L1050" Or NameLayer = "L1049" Or NameLayer = "L1048" Or NameLayer = "L1017" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "PIPELINE_NETWORKS")
    '                                If NameLayer Like "L1051" Then
    '                                    Name = "Трубопровод специального назначения"
    '                                ElseIf NameLayer Like "L1050" Then
    '                                    Name = "Газопровод"
    '                                ElseIf NameLayer Like "L1049" Then
    '                                    Name = "Теплосеть"
    '                                ElseIf NameLayer Like "L1048" Then
    '                                    Name = "Водопровод"
    '                                ElseIf NameLayer Like "L1017" Then
    '                                    Name = "Канализация"
    '                                End If
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            '2d полилинии
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            'Dim ArraySlovo As String() = {"Воздухопровод", "Золопровод", "Паропровод", "Ацетиленопровод", "Этиленопровод", "Бензопровод", "Аммиакопровод", "Технологический", "Материалопровод", "Рассолопровод", "Щелочепровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод"}
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "PIPELINE_NETWORKS")
    '                            If IsArray(ArrayData) = True Then
    '                                Dim ArrayName As String() = NameTypeLine.Split
    '                                If IsArray(ArrayName) = True Then
    '                                    For i As Integer = 0 To ArrayName(i)
    '                                        Dim posName As Integer = ArrayName(i).IndexOf("A-")
    '                                        If posName > -1 Then
    '                                            Name = Mid(ArrayName(i), 3)
    '                                        End If
    '                                    Next i
    '                                End If
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка

    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            'полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "PIPELINE_NETWORKS")
    '                            If IsArray(ArrayData) = True Then
    '                                Dim ArrayName As String() = NameTypeLine.Split
    '                                If IsArray(ArrayName) = True Then
    '                                    For i As Integer = 0 To ArrayName(i)
    '                                        Dim posName As Integer = ArrayName(i).IndexOf("A-")
    '                                        If posName > -1 Then
    '                                            Name = Mid(ArrayName(i), 3)
    '                                        End If
    '                                    Next i
    '                                End If
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            'линия
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Line = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "PIPELINE_NETWORKS")
    '                            If IsArray(ArrayData) = True Then
    '                                Dim ArrayName As String() = NameTypeLine.Split
    '                                If IsArray(ArrayName) = True Then
    '                                    For i As Integer = 0 To ArrayName(i)
    '                                        Dim posName As Integer = ArrayName(i).IndexOf("A-")
    '                                        If posName > -1 Then
    '                                            Name = Mid(ArrayName(i), 3)
    '                                        End If
    '                                    Next i
    '                                End If
    '                                'точка начала и конца линии
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка (точка пересечения с трассой)
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ищем угол пересечения (заносим в отдельный массив)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    'ищем точку сближения с трассой (начало)
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                        End If
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-КМ начало, 6,7-КМ конец, 8,9-ПК начало, 10,11-ПК конец, 12-Глубина закладки, 13,14-Угол пересечения, 15-материал, 16-Диаметр, 17-Владелец, 18-Адрес, 19-Примечание
    '                        '===========================================================================
    '                        'записываем OD Характеристики
    '                        If boolRez = True Then
    '                            If IsArray(ArrayData) = True Then
    '                                For j As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                    Dim NameF As String = ArrayData(1, j)
    '                                    Dim val As String = ArrayData(3, j)
    '                                    If NameF Like "MATERIAL" Then
    '                                        Mat = val
    '                                    ElseIf NameF Like "TYPE" Then
    '                                        TypePipe = val
    '                                    ElseIf NameF Like "HEIGHT" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                Height = CDbl(val)
    '                                            Else
    '                                                Height = 0
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "STATE" Then
    '                                        State = val
    '                                    ElseIf NameF Like "PURPOSE" Then
    '                                        Purpose = val
    '                                    ElseIf NameF Like "DIAMETER" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                Diametr = CDbl(val)
    '                                            Else
    '                                                Diametr = 0
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "COUNT_PIPES" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                CountPipes = CInt(val)
    '                                            Else
    '                                                CountPipes = 1
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "TOP_LEVEL" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                TopLevel = CInt(val)
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "OWNER" Then
    '                                        Owner = val
    '                                    ElseIf NameF Like "ADRESS" Then
    '                                        Adress = val
    '                                    ElseIf NameF Like "NOTE" Then
    '                                        Note = val
    '                                    End If
    '                                Next j
    '                            End If
    '                            'начинаем разбор найденных свойств
    '                            ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '                            '1. записываем наименование объекта
    '                            Dim strName As String = Name & " " & TypePipe & " " & Purpose
    '                            If CountPipes > 1 Then
    '                                strName = strName & " " & CountPipes & " тр."
    '                            End If
    '                            ArrayDataWord(1, CountArrayDataWord) = strName
    '                            '2. записываем пересечение КМ
    '                            If colPointEX.Count > 0 Then
    '                                Dim str As String = ""
    '                                Dim km As Double = 99999999
    '                                For Each pt As Point3d In colPointEX
    '                                    Dim kmper As Double = -1
    '                                    Dim offper As Double = -1
    '                                    Try
    '                                        AlignEX.StationOffset(pt.X, pt.Y, kmper, offper)
    '                                        ArrayDataWord(2, CountArrayDataWord) = ArrayDataWord(2, CountArrayDataWord) & " " & FuncFormatKM(kmper)
    '                                        If kmper < km Then
    '                                            km = kmper
    '                                        End If
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                Next
    '                                ArrayDataWord(0, CountArrayDataWord) = km
    '                            Else
    '                                ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '                            End If
    '                            '3. записывваем пересечение с проектной трассой
    '                            If colPointProject.Count > 0 Then
    '                                Dim str As String = ""
    '                                Dim pk As Double = 99999999
    '                                For Each pt As Point3d In colPointProject
    '                                    Dim pkper As Double = -1
    '                                    Dim offper As Double = -1
    '                                    Try
    '                                        AlignProject.StationOffset(pt.X, pt.Y, pk, offper)
    '                                        ArrayDataWord(3, CountArrayDataWord) = ArrayDataWord(3, CountArrayDataWord) & " " & FuncFormatPK2(pk)
    '                                        If pkper < pk Then
    '                                            pk = pkper
    '                                        End If
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    '===================================================================================
    '                                    'удаляем предыдущие созданные точки
    '                                    Dim NumbPoint As Integer() = Nothing
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(pt.Convert2d(New Plane), NumbPoint, 0.02, "Трубопровод подземный")
    '                                    If IsArray(NumbPoint) = True Then
    '                                        For k As Integer = 0 To NumbPoint.Length - 1
    '                                            Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '                                            Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '                                            If IsDBNull(PointCogoDel) = False Then
    '                                                PointCogoDel.Erase()
    '                                            End If
    '                                        Next k
    '                                    End If
    '                                    'создаем COGO - точку
    '                                    'генерируем номер для cogo точки
    '                                    Dim numbCogo As Long = 0
    '                                    For i As Integer = 0 To 100
    '                                        Dim rng As Random = New Random
    '                                        numbCogo = rng.Next(1000500, 1000999)
    '                                        If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '                                            Exit For
    '                                        End If
    '                                    Next
    '                                    'формируем имя для продольного профиля
    '                                    Dim NameDesk As String = ""
    '                                    If Name.Trim.Length > 0 Then
    '                                        NameDesk = Name
    '                                    End If
    '                                    If Mat.Trim.Length > 0 Then
    '                                        Dim TempMat As String = ""
    '                                        Dim pos As Integer = Mat.IndexOf("Сталь")
    '                                        If pos > -1 Then
    '                                            TempMat = "ст."
    '                                        End If
    '                                        pos = Mat.IndexOf("Железоб")
    '                                        If pos > -1 Then
    '                                            TempMat = "ж.б."
    '                                        End If
    '                                        pos = Mat.IndexOf("Чугун")
    '                                        If pos > -1 Then
    '                                            TempMat = "чуг."
    '                                        End If
    '                                        If TempMat Like "" Then
    '                                            TempMat = Mat
    '                                        End If
    '                                        NameDesk = NameDesk & " " & TempMat
    '                                    End If
    '                                    If Diametr > 0 Then
    '                                        NameDesk = NameDesk & " " & Diametr
    '                                    End If
    '                                    Dim idCogo As ObjectId = FuncCreateCOGOPoint(pt, NameDesk, numbCogo, TopLevel,, "Трубопровод подземный")
    '                                    Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '                                    'присваиваем набор данных
    '                                    Dim ArrayODRec As String(,) = Nothing
    '                                    ReDim Preserve ArrayODRec(2, 1)
    '                                    ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '                                    ArrayODRec(1, 0) = "LABEL"
    '                                    ArrayODRec(2, 0) = "Подземный трубопровод на виде профиля"
    '                                    ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '                                    ArrayODRec(1, 1) = "DESK"
    '                                    ArrayODRec(2, 1) = NameDesk
    '                                    If IsNothing(CreateTBLPSID) = False Then
    '                                        'присваиваем таблицу объекту
    '                                        Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '                                        'записываем данные в таблицу
    '                                        Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '                                    End If
    '                                Next

    '                                If IsNothing(AlignEX) = True Then
    '                                    ArrayDataWord(0, CountArrayDataWord) = pk
    '                                End If
    '                            Else
    '                                If IsNothing(AlignEX) = True Then
    '                                    ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '                                End If
    '                            End If
    '                            '3. записываем км c,kb;tybz
    '                            If km1 <> -1 And km2 <> -1 Then
    '                                If km1 > km2 Then
    '                                    ArrayDataWord(4, CountArrayDataWord) = km2
    '                                    ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '                                    ArrayDataWord(6, CountArrayDataWord) = km1
    '                                    ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '                                Else
    '                                    ArrayDataWord(4, CountArrayDataWord) = km1
    '                                    ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '                                    ArrayDataWord(6, CountArrayDataWord) = km2
    '                                    ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '                                End If
    '                            End If
    '                            If pk1 <> -1 And pk2 <> -1 Then
    '                                If pk1 > pk2 Then
    '                                    ArrayDataWord(8, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '                                    ArrayDataWord(10, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '                                Else
    '                                    ArrayDataWord(8, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '                                    ArrayDataWord(10, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '                                End If
    '                            End If
    '                            'глубина закладки
    '                            ArrayDataWord(12, CountArrayDataWord) = String.Format("{0:0.0}", Math.Round(Height, 1))
    '                            'вычисляем угол
    '                            If IsArray(AngleEX) = True Then
    '                                If AngleEX.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleEX)
    '                                End If
    '                                For i As Integer = 0 To AngleEX.GetUpperBound(1)
    '                                    Dim angl As Double = Math.Round((AngleEX(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    ArrayDataWord(13, CountArrayDataWord) = ArrayDataWord(13, CountArrayDataWord) & vbLf & CStr(angl)
    '                                Next i
    '                            End If
    '                            If IsArray(AngleProject) = True Then
    '                                If AngleProject.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleProject)
    '                                End If
    '                                For i As Integer = 0 To AngleProject.GetUpperBound(1)
    '                                    Dim angl As Double = Math.Round((AngleProject(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    ArrayDataWord(14, CountArrayDataWord) = ArrayDataWord(14, CountArrayDataWord) & vbLf & CStr(angl)
    '                                Next i
    '                            End If
    '                            ArrayDataWord(15, CountArrayDataWord) = Mat 'материал трубы
    '                            ArrayDataWord(16, CountArrayDataWord) = Diametr 'диаметр трубы
    '                            ArrayDataWord(17, CountArrayDataWord) = Owner 'владелец
    '                            ArrayDataWord(18, CountArrayDataWord) = Adress 'адрес
    '                            ArrayDataWord(19, CountArrayDataWord) = Note 'примечание
    '                            CountArrayDataWord += 1
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'записываем все в массив
    '            If IsArray(ArrayDataWord) = True Then
    '                'сортируем массив
    '                Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 0)
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пересечения и сближения подземных коммуникаций (трубопроводные сети).docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость подземных трубопроводных сетей_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                    Exit Sub
    '                End Try
    '                For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
    '                            Exit For
    '                        End If
    '                        tabl.Rows.Add()
    '                        '===================================================================================
    '                        'название объекта
    '                        If IsNothing(ArrayDataWord(1, i)) = False Then
    '                            tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(1, i)
    '                        End If
    '                        '===================================================================================
    '                        'километр пересечения
    '                        If IsNothing(ArrayDataWord(2, i)) = False Then
    '                            tabl.Cell(countTableword, 2).Range.Text = ArrayDataWord(2, i)
    '                        End If
    '                        '===================================================================================
    '                        'записываем смещение относительно существующей трассы
    '                        'String.Format("{0:0.00}", Math.Abs(Math.Round(kmoff2, 2)))
    '                        If IsNothing(ArrayDataWord(4, i)) = False And IsNothing(ArrayDataWord(5, i)) = False Then
    '                            Dim kmStart As String = FuncFormatKM(Val(ArrayDataWord(4, i)))
    '                            Dim kmoff1 As Double = Val(ArrayDataWord(5, i))
    '                            If kmoff1 < 0 Then
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'влево
    '                            ElseIf kmoff1 > 0 Then
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart
    '                            End If
    '                        End If
    '                        If IsNothing(ArrayDataWord(6, i)) = False And IsNothing(ArrayDataWord(7, i)) = False Then
    '                            Dim kmEnd As String = FuncFormatKM(Val(ArrayDataWord(6, i)))
    '                            Dim kmoff2 As Double = Val(ArrayDataWord(7, i))
    '                            If kmoff2 < 0 Then
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'влево
    '                            ElseIf kmoff2 > 0 Then
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd
    '                            End If
    '                        End If
    '                        '=================================================================================================
    '                        'записываем угол пересечения
    '                        If IsNothing(ArrayDataWord(13, i)) = False Then
    '                            tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(13, i)
    '                        End If
    '                        '=================================================================================================
    '                        'записываем пикет пересечения
    '                        If IsNothing(ArrayDataWord(3, i)) = False Then
    '                            tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(3, i)
    '                        End If
    '                        '=================================================================================================
    '                        'записываем смещение относительно существующей трассы
    '                        If IsNothing(ArrayDataWord(8, i)) = False And IsNothing(ArrayDataWord(9, i)) = False Then
    '                            Dim pkStart As String = FuncFormatPK2(Val(ArrayDataWord(8, i)))
    '                            Dim pkoff1 As Double = Val(ArrayDataWord(9, i))
    '                            If pkoff1 < 0 Then
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'влево
    '                            ElseIf pkoff1 > 0 Then
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart
    '                            End If
    '                        End If
    '                        If IsNothing(ArrayDataWord(10, i)) = False And IsNothing(ArrayDataWord(11, i)) = False Then
    '                            Dim pkEnd As String = FuncFormatPK2(Val(ArrayDataWord(10, i)))
    '                            Dim pkoff2 As Double = Val(ArrayDataWord(11, i))
    '                            If pkoff2 < 0 Then
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'влево
    '                            ElseIf pkoff2 > 0 Then
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd
    '                            End If
    '                        End If

    '                        '=================================================================================================
    '                        'записываем угол пересечения
    '                        If IsNothing(ArrayDataWord(14, i)) = False Then
    '                            tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(14, i)
    '                        End If
    '                        '===================================================================================================
    '                        If IsNothing(ArrayDataWord(12, i)) = False Then
    '                            tabl.Cell(countTableword, 10).Range.Text = ArrayDataWord(12, i) 'глубина закладки
    '                        End If
    '                        '==================================================================================================
    '                        'материал и диаметр
    '                        Dim strMat As String = ""
    '                        If IsNothing(ArrayDataWord(15, i)) = False Then
    '                            strMat = ArrayDataWord(15, i)
    '                        End If
    '                        Dim strD As String = ""
    '                        If IsNothing(ArrayDataWord(16, i)) = False Then
    '                            strD = ArrayDataWord(16, i)
    '                        End If
    '                        If strMat.Length > 0 And strD.Length > 0 Then
    '                            tabl.Cell(countTableword, 11).Range.Text = strMat & vbLf & strD
    '                        ElseIf strMat.Length > 0 And strD.Length = 0 Then
    '                            tabl.Cell(countTableword, 11).Range.Text = strMat
    '                        ElseIf strMat.Length = 0 And strD.Length > 0 Then
    '                            tabl.Cell(countTableword, 11).Range.Text = strD
    '                        Else
    '                            tabl.Cell(countTableword, 11).Range.Text = "-"
    '                        End If
    '                        '=================================================================================================
    '                        If IsNothing(ArrayDataWord(17, i)) = False Then
    '                            tabl.Cell(countTableword, 12).Range.Text = ArrayDataWord(17, i) '& vbLf & ArrayDataWord(18, i) 'владелец и адрес
    '                        End If
    '                        If IsNothing(ArrayDataWord(19, i)) = False Then
    '                            tabl.Cell(countTableword, 13).Range.Text = ArrayDataWord(19, i) 'описание
    '                        End If
    '                        Exit For
    '                    Next
    '                Next i
    '                WApp.Visible = True
    '            End If
    '            acTrans.Commit()
    '        End Using
    '    End Sub
    '    '==================================================================================================
    '    'Ведомость пересечения (сближения) подземных кабелей
    '    <CommandMethod("STPR-VedNetworkCable")>
    '    Public Sub VedNetworkCable()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayDataWord As String(,) = Nothing
    '        Dim CountArrayDataWord As Integer = 0
    '        Dim boolRez As Boolean = False
    '        Dim AngleEX As Double(,) = Nothing
    '        Dim countAngleEx As Integer = 0
    '        Dim AngleProject As Double(,) = Nothing
    '        Dim countAngleProject As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость пересечения и сближения подземных коммуникаций (кабели).docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость пересечения и сближения подземных коммуникаций (кабели).docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\Topografiya\MapClassification\"
    '        'Путь к файлам классификации
    '        If IO.Directory.Exists(parentPatch2 & "\Topografiya\MapClassification\") Then
    '            PutchMapClass = parentPatch2 & "\Topografiya\MapClassification\"
    '        End If
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim PlineAlignEx As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim PlineAlignProject As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '                PlineAlignProject = acTrans.GetObject(AlignProject.GetPolyline, OpenMode.ForRead)
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'тело программы
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'выбираем объекты
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim colPointEX As Point3dCollection = New Point3dCollection 'точки пересечения с существующей трассой
    '                        Dim colPointProject As Point3dCollection = New Point3dCollection 'точки пересечения с проектной трассой
    '                        'начало и конец километрового положения трубопровода
    '                        Dim km1 As Double = -1
    '                        Dim kmoff1 As Double = -1
    '                        Dim km2 As Double = -1
    '                        Dim kmoff2 As Double = -1
    '                        'начало и конец пикетажного положения трубопровода
    '                        Dim pk1 As Double = -1
    '                        Dim pkoff1 As Double = -1
    '                        Dim pk2 As Double = -1
    '                        Dim pkoff2 As Double = -1
    '                        Dim Name As String = "" 'наименование сети в общем виде например Водопровод
    '                        Dim Purpose As String = "" 'назначение сети например Речной воды или высокого давления
    '                        Dim Mat As String = "" 'марка провода
    '                        Dim Height As Double = 0 'средняя глубина заложения (высота над поверхностью земли) до верхней образующей
    '                        Dim CountPipes As Integer = 1 'число каналов, кабелей
    '                        Dim State As String = "" 'состояние (действуюжщий, не действующай)
    '                        Dim Owner As String = "" 'владелец
    '                        Dim Adress As String = "" 'адрес
    '                        Dim Note As String = "" 'примечание
    '                        Dim StyleType As String = ""
    '                        Dim Power As String = ""
    '                        Dim TopLevel As Double = 0
    '                        'чистим массивы
    '                        Erase AngleEX 'массив с углами пересечек по сущестыующей трассе
    '                        Erase AngleProject 'массив с углами пересечек по проектной трассе
    '                        Erase ArrayData 'масив с расширенными данными
    '                        countAngleEx = 0
    '                        countAngleProject = 0
    '                        boolRez = False
    '                        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
    '                            Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
    '                            If FLine.StyleName Like "ИИ-1052_Кабель связи подземный" Or FLine.StyleName Like "ИИ-1016_Кабель электрический подземный" Then
    '                                StyleType = FLine.StyleName
    '                                boolRez = True
    '                                Name = "Кабель"
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "CABLES_UNDERGROUND")
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                'превращаем характерныю линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                colPt = FLine.GetPoints(FeatureLinePointType.PIPoint)
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    'ищем пикетаж (существующий километр) начало участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(acPoly) = False Then
    '                                            Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            '3d полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "L1052" Or NameLayer Like "L1016" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "CABLES_UNDERGROUND")
    '                                Name = "Кабель"
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            '2d полилинии
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            'Dim ArraySlovo As String() = {"Воздухопровод", "Золопровод", "Паропровод", "Ацетиленопровод", "Этиленопровод", "Бензопровод", "Аммиакопровод", "Технологический", "Материалопровод", "Рассолопровод", "Щелочепровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод", "Воздухопровод"}
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "CABLES_UNDERGROUND")
    '                            If IsArray(ArrayData) = True Then
    '                                Name = "Кабель"
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка

    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            'полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "CABLES_UNDERGROUND")
    '                            If IsArray(ArrayData) = True Then
    '                                Name = "Кабель"
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim v3d As Vector3d = FLine.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                            'линия
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Line = Ent
    '                            Dim NameTypeLine As String = FLine.Linetype
    '                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "CABLES_UNDERGROUND")
    '                            If IsArray(ArrayData) = True Then
    '                                Name = "Кабель"
    '                                'точка начала и конца линии
    '                                Dim PTStart As Point3d = FLine.StartPoint
    '                                Dim PTEnd As Point3d = FLine.EndPoint
    '                                boolRez = True
    '                                'ищем пикетаж (существующий километр) начало участка (точка пересечения с трассой)
    '                                If IsNothing(AlignEX) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ищем угол пересечения (заносим в отдельный массив)
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(1, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                    'ищем точку сближения с трассой (начало)
    '                                    Try
    '                                        AlignEX.StationOffset(PTStart.X, PTStart.Y, km1, kmoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.StartPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'ищем пикетаж (существующий километр) конец участка
    '                                    Try
    '                                        AlignEX.StationOffset(PTEnd.X, PTEnd.Y, km2, kmoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignEX.EndPoint, False)
    '                                            Try
    '                                                AlignEX.StationOffset(t1.X, t1.Y, 0.1, km2, kmoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                If IsNothing(AlignProject) Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, New Plane, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(1, countAngleProject)
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                    Try
    '                                        AlignProject.StationOffset(PTStart.X, PTStart.Y, pk1, pkoff1)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.StartPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk1, pkoff1)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                    'проектный пикетаж (конец участка)
    '                                    Try
    '                                        AlignProject.StationOffset(PTEnd.X, PTEnd.Y, pk2, pkoff2)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        'конец линии за пределами трассы
    '                                        'ищем пересечение с началом трассы
    '                                        If IsNothing(FLine) = False Then
    '                                            Dim t1 As Point3d = FLine.GetClosestPointTo(AlignProject.EndPoint, False)
    '                                            Try
    '                                                AlignProject.StationOffset(t1.X, t1.Y, 0.1, pk2, pkoff2)
    '                                            Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '                                            End Try
    '                                        End If
    '                                    End Try
    '                                End If
    '                            End If
    '                        End If
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-КМ начало, 6,7-КМ конец, 8,9-ПК начало, 10,11-ПК конец, 12-Глубина закладки, 13,14-Угол пересечения, 15-материал, 16-Диаметр, 17-Владелец, 18-Адрес, 19-Примечание
    '                        '===========================================================================
    '                        'записываем OD Характеристики
    '                        If boolRez = True Then
    '                            If IsArray(ArrayData) = True Then
    '                                For j As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                    Dim NameF As String = ArrayData(1, j)
    '                                    Dim val As String = ArrayData(2, j)
    '                                    If NameF Like "MARK_WIRES" Then
    '                                        Mat = val
    '                                    ElseIf NameF Like "DEPTH" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                Height = CDbl(val)
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "STATE" Then
    '                                        State = val
    '                                    ElseIf NameF Like "PURPOSE" Then
    '                                        Purpose = val
    '                                    ElseIf NameF Like "VOLTAGE" Then
    '                                        Power = val
    '                                    ElseIf NameF Like "NUMB_WIRES" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                CountPipes = CInt(val)
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "TOP_LEVEL" Then
    '                                        Try
    '                                            If IsNumeric(val) = True Then
    '                                                TopLevel = CDbl(val)
    '                                            End If
    '                                        Catch ex As System.InvalidCastException
    '                                        End Try
    '                                    ElseIf NameF Like "OWNER" Then
    '                                        Owner = val
    '                                    ElseIf NameF Like "ADRESS" Then
    '                                        Adress = val
    '                                    ElseIf NameF Like "NOTE" Then
    '                                        Note = val
    '                                    End If
    '                                Next j
    '                            End If
    '                            'начинаем разбор найденных свойств
    '                            ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '                            '1. записываем наименование объекта
    '                            Dim strName As String = Name
    '                            If Purpose Like "Кабельная канализация" Then
    '                                strName = "Кабельная канализация"
    '                                If CountPipes > 0 Then
    '                                    strName = strName & " " & CountPipes & " каналов"
    '                                End If
    '                            Else
    '                                strName = strName & " " & Purpose
    '                                If CountPipes > 0 Then
    '                                    strName = strName & ", " & CountPipes & " кабелей"
    '                                End If
    '                            End If
    '                            ArrayDataWord(1, CountArrayDataWord) = strName
    '                            '2. записываем пересечение КМ
    '                            If colPointEX.Count > 0 Then
    '                                Dim str As String = ""
    '                                Dim km As Double = 99999999
    '                                For Each pt As Point3d In colPointEX
    '                                    Dim kmper As Double = -1
    '                                    Dim offper As Double = -1
    '                                    Try
    '                                        AlignEX.StationOffset(pt.X, pt.Y, kmper, offper)
    '                                        ArrayDataWord(2, CountArrayDataWord) = ArrayDataWord(2, CountArrayDataWord) & " " & FuncFormatKM(kmper)
    '                                        If kmper < km Then
    '                                            km = kmper
    '                                        End If
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                Next
    '                                ArrayDataWord(0, CountArrayDataWord) = km
    '                            Else
    '                                ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '                            End If
    '                            '3. записывваем пересечение с проектной трассой
    '                            If colPointProject.Count > 0 Then
    '                                Dim str As String = ""
    '                                Dim pk As Double = 99999999
    '                                For Each pt As Point3d In colPointProject
    '                                    Dim pkper As Double = -1
    '                                    Dim offper As Double = -1
    '                                    Try
    '                                        AlignProject.StationOffset(pt.X, pt.Y, pk, offper)
    '                                        ArrayDataWord(3, CountArrayDataWord) = ArrayDataWord(3, CountArrayDataWord) & " " & FuncFormatPK2(pk)
    '                                        If pkper < pk Then
    '                                            pk = pkper
    '                                        End If
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    '===================================================================================
    '                                    'создаем COGO - точку
    '                                    '===================================================================================
    '                                    'удаляем предыдущие созданные точки
    '                                    Dim NumbPoint As Integer() = Nothing
    '                                    Dim boolIns As Boolean = FuncFindCogoPoint(pt.Convert2d(New Plane), NumbPoint, 0.02, "Кабели подземные")
    '                                    If IsArray(NumbPoint) = True Then
    '                                        For k As Integer = 0 To NumbPoint.Length - 1
    '                                            Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '                                            Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '                                            If IsDBNull(PointCogoDel) = False Then
    '                                                PointCogoDel.Erase()
    '                                            End If
    '                                        Next k
    '                                    End If
    '                                    'генерируем номер для cogo точки
    '                                    Dim numbCogo As Long = 0
    '                                    For i As Integer = 0 To 100
    '                                        Dim rng As Random = New Random
    '                                        numbCogo = rng.Next(1000500, 1000999)
    '                                        If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '                                            Exit For
    '                                        End If
    '                                    Next
    '                                    Dim idCogo As ObjectId = FuncCreateCOGOPoint(pt, Name, numbCogo, TopLevel,, "Кабели подземные")
    '                                    Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '                                    'присваиваем набор данных
    '                                    Dim ArrayODRec As String(,) = Nothing
    '                                    ReDim Preserve ArrayODRec(2, 1)
    '                                    ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '                                    ArrayODRec(1, 0) = "LABEL"
    '                                    ArrayODRec(2, 0) = "Кабели подземные на виде профиля"
    '                                    ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '                                    ArrayODRec(1, 1) = "DESK"
    '                                    ArrayODRec(2, 1) = Name
    '                                    If IsNothing(CreateTBLPSID) = False Then
    '                                        'присваиваем таблицу объекту
    '                                        Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '                                        'записываем данные в таблицу
    '                                        Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '                                    End If
    '                                Next

    '                                If IsNothing(AlignEX) = True Then
    '                                    ArrayDataWord(0, CountArrayDataWord) = pk
    '                                End If
    '                            Else
    '                                If IsNothing(AlignEX) = True Then
    '                                    ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '                                End If
    '                            End If
    '                            '3. записываем км c,kb;tybz
    '                            If km1 <> -1 And km2 <> -1 Then
    '                                If km1 > km2 Then
    '                                    ArrayDataWord(4, CountArrayDataWord) = km2
    '                                    ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '                                    ArrayDataWord(6, CountArrayDataWord) = km1
    '                                    ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '                                Else
    '                                    ArrayDataWord(4, CountArrayDataWord) = km1
    '                                    ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '                                    ArrayDataWord(6, CountArrayDataWord) = km2
    '                                    ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '                                End If
    '                            End If
    '                            If pk1 <> -1 And pk2 <> -1 Then
    '                                If pk1 > pk2 Then
    '                                    ArrayDataWord(8, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '                                    ArrayDataWord(10, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '                                Else
    '                                    ArrayDataWord(8, CountArrayDataWord) = pk1
    '                                    ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '                                    ArrayDataWord(10, CountArrayDataWord) = pk2
    '                                    ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '                                End If
    '                            End If
    '                            'глубина закладки
    '                            ArrayDataWord(12, CountArrayDataWord) = Math.Round(Height, 1)
    '                            'вычисляем угол
    '                            If IsArray(AngleEX) = True Then
    '                                If AngleEX.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleEX)
    '                                End If
    '                                For i As Integer = 0 To AngleEX.GetUpperBound(1)
    '                                    Dim angl As Double = Math.Round((AngleEX(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    ArrayDataWord(13, CountArrayDataWord) = ArrayDataWord(13, CountArrayDataWord) & vbLf & CStr(angl)
    '                                Next i
    '                            End If
    '                            If IsArray(AngleProject) = True Then
    '                                If AngleProject.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleProject)
    '                                End If
    '                                For i As Integer = 0 To AngleProject.GetUpperBound(1)
    '                                    Dim angl As Double = Math.Round((AngleProject(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    ArrayDataWord(14, CountArrayDataWord) = ArrayDataWord(14, CountArrayDataWord) & vbLf & CStr(angl)
    '                                Next i
    '                            End If
    '                            ArrayDataWord(15, CountArrayDataWord) = Mat 'марка провода
    '                            ArrayDataWord(17, CountArrayDataWord) = Owner 'владелец
    '                            ArrayDataWord(18, CountArrayDataWord) = Adress 'адрес
    '                            ArrayDataWord(19, CountArrayDataWord) = Note 'примечание
    '                            CountArrayDataWord += 1
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'записываем все в массив
    '            If IsArray(ArrayDataWord) = True Then
    '                'сортируем массив
    '                Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 0)
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'выбираем документ Word
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пересечения и сближения подземных коммуникаций (кабели).docx")
    '                    Dim rng As Random = New Random
    '                    Dim Numbrng As Integer = rng.Next
    '                    WDoc.SaveAs2(TempDir & "Ведомость подземных кабельных сетей_" & Numbrng & ".docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                    locker.Dispose()
    '                    WDoc = Nothing
    '                    WApp = Nothing
    '                    Exit Sub
    '                End Try
    '                For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
    '                            Exit For
    '                        End If
    '                        tabl.Rows.Add()
    '                        '===================================================================================
    '                        'название объекта
    '                        If IsNothing(ArrayDataWord(1, i)) = False Then
    '                            tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(1, i)
    '                        End If
    '                        '===================================================================================
    '                        'километр пересечения
    '                        If IsNothing(ArrayDataWord(2, i)) = False Then
    '                            tabl.Cell(countTableword, 2).Range.Text = ArrayDataWord(2, i)
    '                        End If
    '                        '===================================================================================
    '                        'записываем смещение относительно существующей трассы
    '                        If IsNothing(ArrayDataWord(4, i)) = False And IsNothing(ArrayDataWord(5, i)) = False Then
    '                            Dim kmStart As String = FuncFormatKM(Val(ArrayDataWord(4, i)))
    '                            Dim kmoff1 As Double = Val(ArrayDataWord(5, i))
    '                            If kmoff1 < 0 Then
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'влево
    '                            ElseIf kmoff1 > 0 Then
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 3).Range.Text = kmStart
    '                            End If
    '                        End If
    '                        If IsNothing(ArrayDataWord(6, i)) = False And IsNothing(ArrayDataWord(7, i)) = False Then
    '                            Dim kmEnd As String = FuncFormatKM(Val(ArrayDataWord(6, i)))
    '                            Dim kmoff2 As Double = Val(ArrayDataWord(7, i))
    '                            If kmoff2 < 0 Then
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'влево
    '                            ElseIf kmoff2 > 0 Then
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 4).Range.Text = kmEnd
    '                            End If
    '                        End If
    '                        '=================================================================================================
    '                        'записываем угол пересечения
    '                        If IsNothing(ArrayDataWord(13, i)) = False Then
    '                            tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(13, i)
    '                        End If
    '                        '=================================================================================================
    '                        'записываем пикет пересечения
    '                        If IsNothing(ArrayDataWord(3, i)) = False Then
    '                            tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(3, i)
    '                        End If
    '                        '=================================================================================================
    '                        'записываем смещение относительно существующей трассы
    '                        If IsNothing(ArrayDataWord(8, i)) = False And IsNothing(ArrayDataWord(9, i)) = False Then
    '                            Dim pkStart As String = FuncFormatPK2(Val(ArrayDataWord(8, i)))
    '                            Dim pkoff1 As Double = Val(ArrayDataWord(9, i))
    '                            If pkoff1 < 0 Then
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'влево
    '                            ElseIf pkoff1 > 0 Then
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 7).Range.Text = pkStart
    '                            End If
    '                        End If
    '                        If IsNothing(ArrayDataWord(10, i)) = False And IsNothing(ArrayDataWord(11, i)) = False Then
    '                            Dim pkEnd As String = FuncFormatPK2(Val(ArrayDataWord(10, i)))
    '                            Dim pkoff2 As Double = Val(ArrayDataWord(11, i))
    '                            If pkoff2 < 0 Then
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'влево
    '                            ElseIf pkoff2 > 0 Then
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'вправо
    '                            Else
    '                                tabl.Cell(countTableword, 8).Range.Text = pkEnd
    '                            End If
    '                        End If

    '                        '=================================================================================================
    '                        'записываем угол пересечения
    '                        If IsNothing(ArrayDataWord(14, i)) = False Then
    '                            tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(14, i)
    '                        End If
    '                        '===================================================================================================
    '                        'глубина закладки
    '                        If IsNothing(ArrayDataWord(12, i)) = False Then
    '                            tabl.Cell(countTableword, 10).Range.Text = ArrayDataWord(12, i)
    '                        End If
    '                        '==================================================================================================
    '                        'марка провода
    '                        If IsNothing(ArrayDataWord(15, i)) = False Then
    '                            tabl.Cell(countTableword, 11).Range.Text = ArrayDataWord(15, i)
    '                        End If
    '                        '=================================================================================================
    '                        If IsNothing(ArrayDataWord(17, i)) = False Then
    '                            tabl.Cell(countTableword, 12).Range.Text = ArrayDataWord(17, i) '& vbLf & ArrayDataWord(18, i) 'владелец и адрес
    '                        End If
    '                        If IsNothing(ArrayDataWord(19, i)) = False Then
    '                            tabl.Cell(countTableword, 13).Range.Text = ArrayDataWord(19, i) 'описание
    '                        End If
    '                        Exit For
    '                    Next
    '                Next i
    '                WApp.Visible = True
    '            End If
    '            acTrans.Commit()
    '        End Using
    '    End Sub
    '    '==================================================================================================
    '    'Ведомость пересечения воздушных ЛЭП и связи
    '    <CommandMethod("STPR-VedNetworkPowerLines")>
    '    Public Sub VedNetworkPowerLines()
    '        Dim WApp As Word.Application = Nothing
    '        Dim WDoc As Word.Document = Nothing
    '        Dim ArrayDataWord As String(,) = Nothing
    '        Dim CountArrayDataWord As Integer = 0
    '        Dim boolRez As Boolean = False
    '        Dim AngleEX As Double(,) = Nothing
    '        Dim countAngleEx As Integer = 0
    '        Dim AngleProject As Double(,) = Nothing
    '        Dim countAngleProject As Integer = 0
    '        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '        'открываем базу данных чертежа
    '        Dim ed As Editor = acDoc.Editor
    '        Dim db As Database = acDoc.Database
    '        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '        'ищем пути для таблиц
    '        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '        'отделяем имя файла от папки
    '        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '        'определяем имя родительской папки для папки где находится сборка
    '        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '        Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '        If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость пересечения воздушных линий электропередач и связи.docx") = True Then
    '            PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '        Else
    '            If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость пересечения воздушных линий электропередач и связи.docx") = False Then
    '                MsgBox("Шаблон не найден! Макро прервано.")
    '                Exit Sub
    '            End If
    '        End If
    '        Dim PutchMapClass As String = "k:\STPRGeoTools\Topografiya\MapClassification\"
    '        'Путь к файлам классификации
    '        If IO.Directory.Exists(parentPatch2 & "\Topografiya\MapClassification\") Then
    '            PutchMapClass = parentPatch2 & "\Topografiya\MapClassification\"
    '        End If
    '        Dim CreateTBLPSID As ObjectId = Nothing
    '        If File.Exists(PutchMapClass & "\ClassObject.xml") = True Then
    '            CreateTBLPSID = Stroyproekt.funcCreatePropertySetXML("ClassObject.xml", PutchMapClass, "PROFILE_PROJECT")
    '        End If
    '        'временный каталог для автосохранения таблицы
    '        Dim TempDir As String = ""
    '        Try
    '            TempDir = System.IO.Path.GetTempPath()
    '        Catch ex As System.Security.SecurityException
    '            TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '        End Try
    '        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '            'находим ось существующей трассы
    '            Dim AlignProject As Alignment = Nothing
    '            Dim AlignEX As Alignment = Nothing
    '            Dim PlineAlignEx As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim PlineAlignProject As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '            Dim ArrayData As String(,) = Nothing
    '            For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '                Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '                If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '                    AlignEX = TempAlign
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            Next
    '            If IsNothing(AlignEX) = True Then
    '                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                    AlignEX = ent1
    '                    PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '                End If
    '            End If
    '            'выбираем ось проектной трассы
    '            Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '            Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent2
    '                PlineAlignProject = acTrans.GetObject(AlignProject.GetPolyline, OpenMode.ForRead)
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'тело программы
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '            'забираем в массив блоки с провисами проводов
    '            Dim CollObjPrWire As ObjectIdCollection = New ObjectIdCollection
    '            Dim ArrayPrWire As Double(,) = Nothing
    '            Dim CountArrayPrWire As Integer = 0
    '            Dim boolPt As Boolean = FuncFindBLKToPartName("Провис провода", CollObjPrWire)
    '            If CollObjPrWire.Count > 0 Then
    '                For Each id As ObjectId In CollObjPrWire
    '                    Dim blk As BlockReference = acTrans.GetObject(id, OpenMode.ForRead)
    '                    Dim pr As Double = CDbl(FuncReadValueAttr(blk, "TEXT"))
    '                    ReDim Preserve ArrayPrWire(2, CountArrayPrWire)
    '                    ArrayPrWire(0, CountArrayPrWire) = blk.Position.X
    '                    ArrayPrWire(1, CountArrayPrWire) = blk.Position.Y
    '                    ArrayPrWire(2, CountArrayPrWire) = pr
    '                    CountArrayPrWire += 1
    '                Next
    '            End If
    '            'выбираем объекты
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim colPointEX As Point3dCollection = New Point3dCollection 'точки пересечения с существующей трассой
    '                        Dim colPointProject As Point3dCollection = New Point3dCollection 'точки пересечения с проектной трассой
    '                        Dim Mat As String = "" 'марка провода
    '                        Dim CountWires As String = "" 'число проводов
    '                        Dim State As String = "" 'состояние (действуюжщий, не действующай)
    '                        Dim Owner As String = "" 'владелец
    '                        Dim Adress As String = "" 'адрес
    '                        Dim Note As String = "" 'примечание
    '                        Dim StyleType As String = ""
    '                        Dim Power As String = "" 'напряжение
    '                        Dim Name As String = "" 'наименование линии (общее) напиример Воздушная линия ЛЭП или Воздушная линия связи
    '                        Dim Name1 As String = "" 'наменование линии
    '                        Dim Name2 As String = "" 'используется если линия двухцепная
    '                        Dim SlakWires As Double = 0 'провис провода
    '                        Dim Temp As String = "" 'температура
    '                        Dim DataSurv As String = "" 'дата съемки
    '                        Dim TypePer = 0
    '                        Dim CogoDesk As String = "ЛЭП"
    '                        Dim DownLewel As Double = -1
    '                        'чистим массивы
    '                        Erase AngleEX 'массив с углами пересечек по сущестыующей трассе и координатами левого и правого узла линии
    '                        Erase AngleProject 'массив с углами пересечек по проектной трассе и координатами левого и правого узла линии
    '                        Erase ArrayData 'масив с расширенными данными
    '                        countAngleEx = 0
    '                        countAngleProject = 0
    '                        boolRez = False
    '                        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '                        If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
    '                            Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
    '                            If FLine.StyleName Like "ИИ-1021_Воздушная линия связи" Or FLine.StyleName Like "ИИ-1014_Воздушная ЛЭП" Then
    '                                StyleType = FLine.StyleName
    '                                boolRez = True
    '                                Dim pos As Integer = StyleType.IndexOf("_")
    '                                If pos > -1 Then
    '                                    Name = Mid(StyleType, pos + 2)
    '                                End If
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "AIR_LINES_NETWORK")
    '                                'превращаем характерныю линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                colPt = FLine.GetPoints(FeatureLinePointType.PIPoint)
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(7, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle 'угол пересечения
    '                                            'координаты первой вершины
    '                                            AngleEX(2, countAngleEx) = pt1.X
    '                                            AngleEX(3, countAngleEx) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleEX(4, countAngleEx) = pt2.X
    '                                            AngleEX(5, countAngleEx) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleEX(6, countAngleEx) = pt.X
    '                                            AngleEX(7, countAngleEx) = pt.Y
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(8, countAngleProject) '8 элемент отметка провиса нижнего провода
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            'координаты первой вершины
    '                                            AngleProject(2, countAngleProject) = pt1.X
    '                                            AngleProject(3, countAngleProject) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleProject(4, countAngleProject) = pt2.X
    '                                            AngleProject(5, countAngleProject) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleProject(6, countAngleProject) = pt.X
    '                                            AngleProject(7, countAngleProject) = pt.Y
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                            End If
    '                            '3d полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "L1021" Or NameLayer Like "L1014" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "AIR_LINES_NETWORK")
    '                                If NameLayer Like "L1021" Then
    '                                    Name = "Воздушная ЛЭП"
    '                                Else
    '                                    Name = "Воздушная линия связи"
    '                                End If
    '                                'превращаем 3d линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                For Each vid As ObjectId In FLine
    '                                    Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForRead)
    '                                    colPt.Add(v3d.Position)
    '                                Next
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(7, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle 'угол пересечения
    '                                            'координаты первой вершины
    '                                            AngleEX(2, countAngleEx) = pt1.X
    '                                            AngleEX(3, countAngleEx) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleEX(4, countAngleEx) = pt2.X
    '                                            AngleEX(5, countAngleEx) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleEX(6, countAngleEx) = pt.X
    '                                            AngleEX(7, countAngleEx) = pt.Y
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(8, countAngleProject) '8 элемент отметка провиса нижнего провода
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            'координаты первой вершины
    '                                            AngleProject(2, countAngleProject) = pt1.X
    '                                            AngleProject(3, countAngleProject) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleProject(4, countAngleProject) = pt2.X
    '                                            AngleProject(5, countAngleProject) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleProject(6, countAngleProject) = pt.X
    '                                            AngleProject(7, countAngleProject) = pt.Y
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                            End If
    '                            '2d полилинии
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Or NameLayer Like "ИИ-ТОПО-Воздушные линии связи" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "AIR_LINES_NETWORK")
    '                                If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Then
    '                                    Name = "Воздушная ЛЭП"
    '                                Else
    '                                    Name = "Воздушная линия связи"
    '                                End If
    '                                'превращаем 2d линию в полилинию
    '                                Dim colPt As Point3dCollection = New Point3dCollection
    '                                For Each vid As ObjectId In FLine
    '                                    Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForRead)
    '                                    colPt.Add(v3d.Position)
    '                                Next
    '                                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '                                If colPt.Count > 1 Then
    '                                    acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '                                    For i As Integer = 0 To colPt.Count - 1
    '                                        acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '                                    Next
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(7, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle 'угол пересечения
    '                                            'координаты первой вершины
    '                                            AngleEX(2, countAngleEx) = pt1.X
    '                                            AngleEX(3, countAngleEx) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleEX(4, countAngleEx) = pt2.X
    '                                            AngleEX(5, countAngleEx) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleEX(6, countAngleEx) = pt.X
    '                                            AngleEX(7, countAngleEx) = pt.Y
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(8, countAngleProject) '8 элемент отметка провиса нижнего провода
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            'координаты первой вершины
    '                                            AngleProject(2, countAngleProject) = pt1.X
    '                                            AngleProject(3, countAngleProject) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleProject(4, countAngleProject) = pt2.X
    '                                            AngleProject(5, countAngleProject) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleProject(6, countAngleProject) = pt.X
    '                                            AngleProject(7, countAngleProject) = pt.Y
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                            End If
    '                            'полилиния
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '                            Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
    '                            Dim NameLayer As String = acPoly.Layer
    '                            If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Or NameLayer Like "ИИ-ТОПО-Воздушные линии связи" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(acPoly, ArrayData, "AIR_LINES_NETWORK")
    '                                If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Then
    '                                    Name = "Воздушная ЛЭП"
    '                                Else
    '                                    Name = "Воздушная линия связи"
    '                                End If
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(7, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle 'угол пересечения
    '                                            'координаты первой вершины
    '                                            AngleEX(2, countAngleEx) = pt1.X
    '                                            AngleEX(3, countAngleEx) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleEX(4, countAngleEx) = pt2.X
    '                                            AngleEX(5, countAngleEx) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleEX(6, countAngleEx) = pt.X
    '                                            AngleEX(7, countAngleEx) = pt.Y
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = Nothing
    '                                            Dim pt2 As Point3d = Nothing
    '                                            Dim DistPer As Double = acPoly.GetDistAtPoint(pt)
    '                                            For i As Integer = 0 To acPoly.NumberOfVertices - 1
    '                                                Dim TempDist As Double = acPoly.GetDistAtPoint(acPoly.GetPoint3dAt(i))
    '                                                If TempDist >= DistPer Then
    '                                                    pt1 = acPoly.GetPoint3dAt(i - 1)
    '                                                    pt2 = acPoly.GetPoint3dAt(i)
    '                                                    Exit For
    '                                                End If
    '                                            Next i

    '                                            Dim v3d As Vector3d = acPoly.GetFirstDerivative(pt)
    '                                            Dim a1 As Double = v3d.AngleOnPlane(New Plane)
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(8, countAngleProject) '8 элемент отметка провиса нижнего провода
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            'координаты первой вершины
    '                                            AngleProject(2, countAngleProject) = pt1.X
    '                                            AngleProject(3, countAngleProject) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleProject(4, countAngleProject) = pt2.X
    '                                            AngleProject(5, countAngleProject) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleProject(6, countAngleProject) = pt.X
    '                                            AngleProject(7, countAngleProject) = pt.Y
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                            End If
    '                            'линия
    '                        ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Line Then
    '                            Dim FLine As Autodesk.AutoCAD.DatabaseServices.Line = Ent
    '                            Dim NameLayer As String = FLine.Layer
    '                            If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Or NameLayer Like "ИИ-ТОПО-Воздушные линии связи" Then
    '                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(FLine, ArrayData, "AIR_LINES_NETWORK")
    '                                If NameLayer Like "ИИ-ТОПО-ЛЭП высокого напряжения" Or NameLayer Like "ИИ-ТОПО-ЛЭП низкого напряжения" Then
    '                                    Name = "Воздушная ЛЭП"
    '                                Else
    '                                    Name = "Воздушная линия связи"
    '                                End If

    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignEX) = False Then
    '                                    AlignEX.IntersectWith(FLine, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '                                    'ещем касательную к полилинии
    '                                    If colPointEX.Count > 0 Then
    '                                        For Each pt As Point3d In colPointEX
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = FLine.StartPoint
    '                                            Dim pt2 As Point3d = FLine.EndPoint

    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignEx.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleEX(7, countAngleEx)
    '                                            AngleEX(0, countAngleEx) = dist
    '                                            AngleEX(1, countAngleEx) = dAngle 'угол пересечения
    '                                            'координаты первой вершины
    '                                            AngleEX(2, countAngleEx) = pt1.X
    '                                            AngleEX(3, countAngleEx) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleEX(4, countAngleEx) = pt2.X
    '                                            AngleEX(5, countAngleEx) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleEX(6, countAngleEx) = pt.X
    '                                            AngleEX(7, countAngleEx) = pt.Y
    '                                            countAngleEx += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                'проектный пикетаж (начало участка)
    '                                '1. Ищем пересечение с осью трассы
    '                                If IsNothing(AlignProject) = False Then
    '                                    AlignProject.IntersectWith(FLine, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '                                    If colPointProject.Count > 0 Then
    '                                        For Each pt As Point3d In colPointProject
    '                                            'ищем левую и правую опоры
    '                                            Dim pt1 As Point3d = FLine.StartPoint
    '                                            Dim pt2 As Point3d = FLine.EndPoint

    '                                            Dim a1 As Double = FLine.Angle
    '                                            Dim v3d2 As Vector3d = PlineAlignProject.GetFirstDerivative(pt)
    '                                            Dim a2 As Double = v3d2.AngleOnPlane(New Plane)
    '                                            Dim dAngle As Double = Math.Abs(a2 - a1)
    '                                            Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '                                            ReDim Preserve AngleProject(8, countAngleProject) '8 элемент отметка провиса нижнего провода
    '                                            AngleProject(0, countAngleProject) = dist
    '                                            AngleProject(1, countAngleProject) = dAngle
    '                                            'координаты первой вершины
    '                                            AngleProject(2, countAngleProject) = pt1.X
    '                                            AngleProject(3, countAngleProject) = pt1.Y
    '                                            'координаты второй вершины
    '                                            AngleProject(4, countAngleProject) = pt2.X
    '                                            AngleProject(5, countAngleProject) = pt2.Y
    '                                            'координаты точки пересечения
    '                                            AngleProject(6, countAngleProject) = pt.X
    '                                            AngleProject(7, countAngleProject) = pt.Y
    '                                            countAngleProject += 1
    '                                        Next
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                        '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-угол пересечения, 6,7-расстояние до правой и левой опоры сущ.трассы, 8,9-расстояние до правой и левой опоры пр.трассы, 10-количество проводов, 11,12-характеристика опор, 13-Марка провода и провис 14-Владелец, 15-Адрес, 16-Примечание
    '                        '===========================================================================
    '                        'записываем OD Характеристики
    '                        If boolRez = True Then
    '                            If IsArray(ArrayData) = True Then
    '                                For j As Integer = 0 To ArrayData.GetUpperBound(1)
    '                                    Dim NameF As String = ArrayData(1, j)
    '                                    Dim val As String = ArrayData(2, j)
    '                                    If NameF Like "VOLTAGE" Then 'напряжение линии
    '                                        Power = val
    '                                    ElseIf NameF Like "NAME1" Then 'наименование линии
    '                                        Name1 = val
    '                                    ElseIf NameF Like "NAME2" Then 'наименование линии (цепь №2)
    '                                        Name2 = val
    '                                    ElseIf NameF Like "NUMB_WIRES" Then 'число проводов (кабелей)
    '                                        CountWires = val
    '                                    ElseIf NameF Like "MARK_WIRES" Then 'марка провода
    '                                        Mat = val
    '                                    ElseIf NameF Like "SLAK_WIRES" Then 'провис провода
    '                                        If IsNumeric(val) = True Then
    '                                            If CDbl(val) > 0 Then
    '                                                SlakWires = CDbl(val)
    '                                            End If
    '                                        End If
    '                                    ElseIf NameF Like "DOWN_SLAK_LEVEL" Then 'провис провода
    '                                        If IsNumeric(val) = True Then
    '                                            If CDbl(val) > 0 Then
    '                                                DownLewel = CDbl(val)
    '                                            End If
    '                                        End If
    '                                    ElseIf NameF Like "TEMPERATURE" Then 'температура
    '                                        If IsNumeric(val) = True Then
    '                                            If CDbl(val) > 0 Then
    '                                                Temp = CDbl(val)
    '                                            End If
    '                                        End If
    '                                    ElseIf NameF Like "DATA_SURV" Then 'дата съемки
    '                                        DataSurv = val
    '                                    ElseIf NameF Like "OWNER" Then 'собственник
    '                                        Owner = val
    '                                    ElseIf NameF Like "ADRESS" Then 'адрес владельца
    '                                        Adress = val
    '                                    ElseIf NameF Like "NOTE" Then 'примечание
    '                                        Note = val
    '                                    End If
    '                                Next j
    '                            End If
    '                            '1. записываем наименование объекта
    '                            Dim strName As String = Name
    '                            If strName Like "Воздушная ЛЭП" Then
    '                                TypePer = 0
    '                                CogoDesk = "ЛЭП"
    '                            ElseIf strName Like "Воздушная линия связи" Then
    '                                TypePer = 1
    '                                CogoDesk = "Лин.связи"
    '                            End If
    '                            If Power.Trim.Length > 0 Then
    '                                If IsNumeric(Power) = True Then
    '                                    strName = strName & " напряжением " & Power & "кВ."
    '                                    CogoDesk = CogoDesk & " " & Power & "кВ."
    '                                Else
    '                                    strName = strName & " напряжением " & Power
    '                                    CogoDesk = CogoDesk & " " & Power
    '                                End If
    '                            End If
    '                            'наименование линии
    '                            If Name1.Trim.Length > 0 Then
    '                                strName = strName & vbLf & Name1
    '                            End If
    '                            'наименование цепи 2
    '                            If Name2.Trim.Length > 0 Then
    '                                strName = strName & vbLf & Name2
    '                            End If
    '                            '========================================================================================
    '                            'разбираем массив с существующими пересечениями и запиываем эти характеристики в отдельный массив
    '                            If IsArray(AngleEX) = True Then
    '                                If AngleEX.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleEX)
    '                                End If
    '                                '==================================================================
    '                                'перебираем массив
    '                                For i As Integer = 0 To AngleEX.GetUpperBound(1)
    '                                    'открываем массив
    '                                    'наименование линии
    '                                    ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '                                    ArrayDataWord(1, CountArrayDataWord) = strName
    '                                    'количество проводов
    '                                    ArrayDataWord(10, CountArrayDataWord) = CountWires
    '                                    'координаты точки пересечения
    '                                    Dim pt As Point3d = New Point3d(AngleEX(6, i), AngleEX(7, i), 0)
    '                                    'находим километр пересечения
    '                                    Dim kmper As Double = -1
    '                                    Dim offper As Double = -1
    '                                    Try
    '                                        AlignEX.StationOffset(pt.X, pt.Y, kmper, offper)
    '                                        ArrayDataWord(2, CountArrayDataWord) = FuncFormatKM(kmper)
    '                                        ArrayDataWord(0, CountArrayDataWord) = kmper
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    'форматируем угол пересечения
    '                                    Dim angl As Double = Math.Round((AngleEX(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    ArrayDataWord(4, CountArrayDataWord) = CStr(angl)
    '                                    '=========================================================
    '                                    'ищем левую и правую опору
    '                                    'первая опора
    '                                    Dim x1 As Double = AngleEX(2, i)
    '                                    Dim y1 As Double = AngleEX(3, i)
    '                                    Dim kmpk1 As Double = -1
    '                                    Dim kmoff1 As Double = -1
    '                                    If x1 <> -1 And y1 <> -1 Then
    '                                        Try
    '                                            AlignEX.StationOffset(x1, y1, kmpk1, kmoff1)
    '                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                        If kmpk1 <> -1 And kmoff1 <> -1 Then
    '                                            If kmoff1 < 0 Then
    '                                                kmoff1 = Math.Abs(kmoff1)
    '                                                ArrayDataWord(6, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                                            Else
    '                                                ArrayDataWord(7, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(kmoff1, 2))
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    'вторая опора
    '                                    Dim x2 As Double = AngleEX(4, i) 'предполагается что пересечение одно
    '                                    Dim y2 As Double = AngleEX(5, i)
    '                                    Dim kmpk2 As Double = -1
    '                                    Dim kmoff2 As Double = -1
    '                                    If x2 <> -1 And y2 <> -1 Then
    '                                        Try
    '                                            AlignEX.StationOffset(x2, y2, kmpk2, kmoff2)
    '                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                        If kmpk2 <> -1 And kmoff2 <> -1 Then
    '                                            If kmoff2 < 0 Then
    '                                                ArrayDataWord(6, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(kmoff2, 2))
    '                                            Else
    '                                                ArrayDataWord(7, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(kmoff2, 2))
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    If IsNothing(ArrayDataWord(6, CountArrayDataWord)) = True Then
    '                                        ArrayDataWord(6, CountArrayDataWord) = "-"
    '                                    End If
    '                                    If IsNothing(ArrayDataWord(7, CountArrayDataWord)) = True Then
    '                                        ArrayDataWord(7, CountArrayDataWord) = "-"
    '                                    End If
    '                                    '=======================================================================
    '                                    'находим блоки с левой и правой опорой
    '                                    '============================================================================
    '                                    'ищем левую опору
    '                                    Dim ZLevel As Double = -1
    '                                    Dim ArrayDataPillars As String(,) = Nothing
    '                                    Dim colObj As ObjectIdCollection = New ObjectIdCollection
    '                                    Dim ptL As Point3d = New Point3d(-1, -1, 0)
    '                                    If kmpk1 < 0 Then
    '                                        ptL = New Point3d(x1, y1, 0)
    '                                    Else
    '                                        ptL = New Point3d(x2, y2, 0)
    '                                    End If
    '                                    'находим блок автокада
    '                                    Dim boolPt1 As Boolean = FuncFindBLKToPoint(ptL, colObj,, 0.05)
    '                                    If colObj.Count > 0 Then
    '                                        For j As Integer = 0 To colObj.Count - 1
    '                                            Dim blk As BlockReference = acTrans.GetObject(colObj.Item(j), OpenMode.ForRead)
    '                                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(blk, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                            ZLevel = blk.Position.Z
    '                                            If boolReadPS = True Then
    '                                                Exit For
    '                                            End If
    '                                        Next j
    '                                    Else
    '                                        Dim NumCogo As Integer() = Nothing
    '                                        Dim boolPt2 As Boolean = FuncFindCogoPoint(ptL.Convert2d(New Plane), NumCogo, 0.05)
    '                                        If IsArray(NumCogo) = True Then
    '                                            For j As Integer = 0 To NumCogo.Length - 1
    '                                                Dim IDPtCogo As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumCogo(j))
    '                                                Dim PtCogo As CogoPoint = acTrans.GetObject(IDPtCogo, OpenMode.ForRead)
    '                                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(PtCogo, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                ZLevel = PtCogo.Elevation
    '                                                If boolReadPS = True Then
    '                                                    Exit For
    '                                                End If
    '                                            Next j
    '                                        End If
    '                                    End If
    '                                    '============================================================
    '                                    'характеристика найденной опоры
    '                                    Dim NumbPillars As String = ""
    '                                    Dim DownWire As Double = -1
    '                                    Dim TopWire As Double = -1
    '                                    Dim GRWire As Double = -1
    '                                    Dim VPillars As Double = -1
    '                                    Dim SVWire As Double = -1
    '                                    Dim MatPillars As String = ""
    '                                    If IsArray(ArrayDataPillars) = True Then
    '                                        For j As Integer = 0 To ArrayDataPillars.GetUpperBound(1)
    '                                            Dim NameF As String = ArrayDataPillars(1, j)
    '                                            Dim val As String = ArrayDataPillars(3, j)
    '                                            If NameF Like "PILLAR_NUMBER" Then 'номер опоры
    '                                                NumbPillars = val
    '                                            ElseIf NameF Like "ZERO_LEVEL" Then 'отметка земли
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        ZLevel = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "DOWN_WIRE_LEVEL" Then 'отметка нижнего провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        DownWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "TOP_WIRE_LEVEL" Then 'отметка верхнего провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        TopWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "GROZ_WIRE_LEVEL" Then 'отметка грозотроса
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        GRWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "UP_PILLAR_LEVEL" Then 'отметка верха опоры
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        VPillars = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "DOWN_SV_LEVEL" Then 'отметка вспомогательного провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        SVWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "MATERIAL" Then 'материал опоры
    '                                                MatPillars = val
    '                                            End If
    '                                        Next j
    '                                    End If
    '                                    '===========================================================================
    '                                    'заполняем массив далее
    '                                    'формируем строку с Материал опоры\Номер опоры\Расстояние
    '                                    Dim StrPillars As String = ""
    '                                    If MatPillars.Trim.Length > 0 Then
    '                                        If StrPillars.IndexOf("елезоб") Then
    '                                            StrPillars = "ж.б."
    '                                        ElseIf StrPillars.IndexOf("етал") Then
    '                                            StrPillars = "мет."
    '                                        ElseIf StrPillars.IndexOf("ерев") Then
    '                                            StrPillars = "дер."
    '                                        End If
    '                                    End If
    '                                    If NumbPillars.Trim.Length > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "№" & NumbPillars
    '                                        Else
    '                                            StrPillars = "№" & NumbPillars
    '                                        End If
    '                                    End If
    '                                    If ZLevel > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.земл.- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                        Else
    '                                            StrPillars = "Отм.земли- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                        End If
    '                                    End If
    '                                    If VPillars > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                        Else
    '                                            StrPillars = "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                        End If
    '                                    End If
    '                                    If DownWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                        End If
    '                                    End If
    '                                    If TopWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                        End If
    '                                    End If
    '                                    If GRWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.гр.троса - " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.гр.троса.- " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                        End If
    '                                    End If
    '                                    If SVWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                        End If
    '                                    End If
    '                                    'записываем характеристики левой опоры
    '                                    ArrayDataWord(11, CountArrayDataWord) = StrPillars
    '                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                    'ищем правую опору
    '                                    ZLevel = -1
    '                                    Erase ArrayDataPillars
    '                                    colObj = New ObjectIdCollection
    '                                    Dim ptP As Point3d = New Point3d(-1, -1, 0)
    '                                    If kmpk2 < 0 Then
    '                                        ptP = New Point3d(x2, y2, 0)
    '                                    Else
    '                                        ptP = New Point3d(x1, y1, 0)
    '                                    End If
    '                                    'находим блок автокада
    '                                    boolPt1 = FuncFindBLKToPoint(ptP, colObj,, 0.05)
    '                                    If colObj.Count > 0 Then
    '                                        For j As Integer = 0 To colObj.Count - 1
    '                                            Dim blk As BlockReference = acTrans.GetObject(colObj.Item(j), OpenMode.ForRead)
    '                                            Dim boolReadPS As Boolean = FuncReadDataSetProperties(blk, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                            ZLevel = blk.Position.Z
    '                                            If boolReadPS = True Then
    '                                                Exit For
    '                                            End If
    '                                        Next j
    '                                    Else
    '                                        Dim NumCogo As Integer() = Nothing
    '                                        Dim boolPt2 As Boolean = FuncFindCogoPoint(ptP.Convert2d(New Plane), NumCogo, 0.05)
    '                                        If IsArray(NumCogo) = True Then
    '                                            For j As Integer = 0 To NumCogo.Length - 1
    '                                                Dim IDPtCogo As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumCogo(j))
    '                                                Dim PtCogo As CogoPoint = acTrans.GetObject(IDPtCogo, OpenMode.ForRead)
    '                                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(PtCogo, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                ZLevel = PtCogo.Elevation
    '                                                If boolReadPS = True Then
    '                                                    Exit For
    '                                                End If
    '                                            Next j
    '                                        End If
    '                                    End If
    '                                    '============================================================
    '                                    'характеристика найденной опоры
    '                                    NumbPillars = ""
    '                                    DownWire = -1
    '                                    TopWire = -1
    '                                    GRWire = -1
    '                                    VPillars = -1
    '                                    SVWire = -1
    '                                    MatPillars = ""
    '                                    If IsArray(ArrayDataPillars) = True Then
    '                                        For j As Integer = 0 To ArrayDataPillars.GetUpperBound(1)
    '                                            Dim NameF As String = ArrayDataPillars(1, j)
    '                                            Dim val As String = ArrayDataPillars(3, j)
    '                                            If NameF Like "PILLAR_NUMBER" Then 'номер опоры
    '                                                NumbPillars = val
    '                                            ElseIf NameF Like "ZERO_LEVEL" Then 'отметка земли
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        ZLevel = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "DOWN_WIRE_LEVEL" Then 'отметка нижнего провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        DownWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "TOP_WIRE_LEVEL" Then 'отметка верхнего провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        TopWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "GROZ_WIRE_LEVEL" Then 'отметка грозотроса
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        GRWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "UP_PILLAR_LEVEL" Then 'отметка верха опоры
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        VPillars = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "DOWN_SV_LEVEL" Then 'отметка вспомогательного провода
    '                                                If IsNumeric(val) = True Then
    '                                                    If CDbl(val) > 0 Then
    '                                                        SVWire = CDbl(val)
    '                                                    End If
    '                                                End If
    '                                            ElseIf NameF Like "MATERIAL" Then 'материал опоры
    '                                                MatPillars = val
    '                                            End If
    '                                        Next j
    '                                    End If
    '                                    '===========================================================================
    '                                    'заполняем массив далее
    '                                    'формируем строку с Материал опоры\Номер опоры\Расстояние
    '                                    StrPillars = ""
    '                                    If MatPillars.Trim.Length > 0 Then
    '                                        If StrPillars.IndexOf("елезоб") Then
    '                                            StrPillars = "ж.б."
    '                                        ElseIf StrPillars.IndexOf("етал") Then
    '                                            StrPillars = "мет."
    '                                        ElseIf StrPillars.IndexOf("ерев") Then
    '                                            StrPillars = "дер."
    '                                        End If
    '                                    End If
    '                                    If NumbPillars.Trim.Length > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "№" & NumbPillars
    '                                        Else
    '                                            StrPillars = "№" & NumbPillars
    '                                        End If
    '                                    End If
    '                                    If ZLevel > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.земл.- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                        Else
    '                                            StrPillars = "Отм.земли- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                        End If
    '                                    End If
    '                                    If VPillars > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                        Else
    '                                            StrPillars = "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                        End If
    '                                    End If
    '                                    If DownWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                        End If
    '                                    End If
    '                                    If TopWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                        End If
    '                                    End If
    '                                    If GRWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.гр.троса - " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.гр.троса.- " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                        End If
    '                                    End If
    '                                    If SVWire > 0 Then
    '                                        If StrPillars.Trim.Length > 0 Then
    '                                            StrPillars = StrPillars & vbLf & "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                        Else
    '                                            StrPillars = "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                        End If
    '                                    End If
    '                                    'записываем характеристики правой опоры
    '                                    ArrayDataWord(12, CountArrayDataWord) = StrPillars
    '                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                    'записывает провис провода и температуру
    '                                    Dim LenghtSegment As Double = -1
    '                                    Dim ptTemp1 As Point3d = New Point3d(AngleEX(2, i), AngleEX(3, i), 0)
    '                                    Dim ptTemp2 As Point3d = New Point3d(AngleEX(4, i), AngleEX(5, i), 0)
    '                                    Dim acLine As Line = New Line(ptTemp1, ptTemp2)
    '                                    LenghtSegment = acLine.Length
    '                                    Dim ptTempPer As Point3d = New Point3d(-1, -1, -1)
    '                                    If SlakWires <= 0 Then
    '                                        If IsArray(ArrayPrWire) = True Then
    '                                            For k As Integer = 0 To ArrayPrWire.GetUpperBound(1)
    '                                                Dim ptTempPr As Point3d = New Point3d(ArrayPrWire(0, k), ArrayPrWire(1, k), 0)
    '                                                Dim ptTemp As Point3d = acLine.GetClosestPointTo(ptTempPr, True)
    '                                                Dim distPr As Double = ptTemp.GetVectorTo(ptTempPr).Length
    '                                                If distPr < 0.5 Then
    '                                                    SlakWires = ArrayPrWire(2, k)
    '                                                    ptTempPer = ptTemp
    '                                                    Exit For
    '                                                End If
    '                                            Next k
    '                                        End If
    '                                    End If
    '                                    'шщем отметку
    '                                    If DownLewel <= 0 Then
    '                                        If SlakWires >= 0 Then
    '                                            If ptTempPer.X <> -1 And ptTempPer.Y <> -1 Then
    '                                                '========================================================================================
    '                                                'находим поверхность рабочую
    '                                                Dim elev As Double = 0
    '                                                Dim ColSurv As ObjectIdCollection = CivilDoc.GetSurfaceIds
    '                                                For Each idSurf As ObjectId In ColSurv
    '                                                    Dim surf As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(idSurf, OpenMode.ForRead)
    '                                                    Try
    '                                                        elev = surf.FindElevationAtXY(ptTempPer.X, ptTempPer.Y)
    '                                                        DownLewel = elev + SlakWires
    '                                                        Exit For
    '                                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                                    End Try
    '                                                Next
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    '=================================================================================
    '                                    Dim strMarkWire As String = Mat
    '                                    If SlakWires <> 0 Then
    '                                        strMarkWire = strMarkWire & vbLf & String.Format("{0:0.00}", Math.Round(SlakWires, 2))
    '                                    End If
    '                                    ArrayDataWord(13, CountArrayDataWord) = strMarkWire
    '                                    '================================================================================
    '                                    'дата и температура воздуха
    '                                    Dim StrDataTemp As String = DataSurv
    '                                    If Temp.Trim.Length > 0 Then
    '                                        StrDataTemp = StrDataTemp & vbLf & Temp & Chr(176) & "С"
    '                                    End If
    '                                    ArrayDataWord(14, CountArrayDataWord) = StrDataTemp
    '                                    '=================================================================================
    '                                    'владелец
    '                                    ArrayDataWord(15, CountArrayDataWord) = Owner
    '                                    '==================================================================================
    '                                    'примечание
    '                                    ArrayDataWord(16, CountArrayDataWord) = Note
    '                                    '==================================================================================
    '                                    'длина сегмента
    '                                    ArrayDataWord(17, CountArrayDataWord) = Math.Round(LenghtSegment, 3)
    '                                    CountArrayDataWord += 1
    '                                Next i
    '                            End If
    '                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                            '===============================================================================================
    '                            '========================================================================================
    '                            'разбираем массив с проектными пересечениями и запиываем эти характеристики в отдельный массив
    '                            If IsArray(AngleProject) = True Then
    '                                If AngleProject.GetUpperBound(1) > 0 Then
    '                                    Dim boolSort As Boolean = FuncSortDblArray(AngleProject)
    '                                End If
    '                                '==================================================================
    '                                'перебираем массив
    '                                For i As Integer = 0 To AngleProject.GetUpperBound(1)
    '                                    Dim ptTemp1 As Point3d = New Point3d(AngleProject(2, i), AngleProject(3, i), 0)
    '                                    Dim ptTemp2 As Point3d = New Point3d(AngleProject(4, i), AngleProject(5, i), 0)
    '                                    Dim acLine As Line = New Line(ptTemp1, ptTemp2)
    '                                    Dim LenghtSegment As Double = Math.Round(acLine.Length, 3)
    '                                    Dim PosRec As Integer = -1
    '                                    If IsArray(ArrayDataWord) = True Then
    '                                        For k As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '                                            If LenghtSegment = CDbl(ArrayDataWord(17, k)) Then
    '                                                PosRec = k
    '                                            End If
    '                                        Next k
    '                                    Else
    '                                        ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '                                    End If
    '                                    If PosRec = -1 Then
    '                                        ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '                                    End If
    '                                    If PosRec = -1 Then
    '                                        'наименование линии
    '                                        ArrayDataWord(1, CountArrayDataWord) = strName
    '                                        'количество проводов
    '                                        ArrayDataWord(10, CountArrayDataWord) = CountWires
    '                                    End If
    '                                    'координаты точки пересечения
    '                                    Dim pt As Point3d = New Point3d(AngleProject(6, i), AngleProject(7, i), 0)
    '                                    'находим пикет пересечения
    '                                    Dim pkper As Double = -1
    '                                    Dim pkoffper As Double = -1
    '                                    Try
    '                                        AlignProject.StationOffset(pt.X, pt.Y, pkper, pkoffper)
    '                                        If PosRec = -1 Then
    '                                            ArrayDataWord(3, CountArrayDataWord) = FuncFormatPiketSTEq(AlignProject, pkper)

    '                                        Else
    '                                            ArrayDataWord(3, PosRec) = FuncFormatPiketSTEq(AlignProject, pkper)
    '                                        End If
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    'для сортировки значение используется
    '                                    If IsArray(AngleEX) = False Then
    '                                        ArrayDataWord(0, CountArrayDataWord) = pkper
    '                                    End If
    '                                    'форматируем угол пересечения
    '                                    Dim angl As Double = Math.Round((AngleProject(1, i) * 180) / Math.PI, 0)
    '                                    If angl > 180 Then angl = 360 - angl
    '                                    If PosRec = -1 Then
    '                                        ArrayDataWord(5, CountArrayDataWord) = CStr(angl)
    '                                    Else
    '                                        ArrayDataWord(5, PosRec) = CStr(angl)
    '                                    End If
    '                                    '=========================================================
    '                                    'ищем левую и правую опору
    '                                    'первая опора
    '                                    Dim x1 As Double = AngleProject(2, i)
    '                                    Dim y1 As Double = AngleProject(3, i)
    '                                    Dim pk1 As Double = -1
    '                                    Dim off1 As Double = -1
    '                                    If x1 <> -1 And y1 <> -1 Then
    '                                        Try
    '                                            AlignProject.StationOffset(x1, y1, pk1, off1)
    '                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                        If pk1 <> -1 And off1 <> -1 Then
    '                                            If off1 < 0 Then
    '                                                off1 = Math.Abs(off1)
    '                                                If PosRec = -1 Then
    '                                                    ArrayDataWord(8, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(off1, 2))
    '                                                Else
    '                                                    ArrayDataWord(8, PosRec) = String.Format("{0:0.00}", Math.Round(off1, 2))
    '                                                End If
    '                                            Else
    '                                                If PosRec = -1 Then
    '                                                    ArrayDataWord(9, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(off1, 2))
    '                                                Else
    '                                                    ArrayDataWord(9, PosRec) = String.Format("{0:0.00}", Math.Round(off1, 2))
    '                                                End If
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    'вторая опора
    '                                    Dim x2 As Double = AngleProject(4, i) 'предполагается что пересечение одно
    '                                    Dim y2 As Double = AngleProject(5, i)
    '                                    Dim pk2 As Double = -1
    '                                    Dim off2 As Double = -1
    '                                    If x2 <> -1 And y2 <> -1 Then
    '                                        Try
    '                                            AlignProject.StationOffset(x2, y2, pk2, off2)
    '                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                        End Try
    '                                        If pk2 <> -1 And off2 <> -1 Then
    '                                            If off2 < 0 Then
    '                                                If PosRec = -1 Then
    '                                                    ArrayDataWord(8, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(off2, 2))
    '                                                Else
    '                                                    ArrayDataWord(8, PosRec) = String.Format("{0:0.00}", Math.Round(off2, 2))
    '                                                End If
    '                                            Else
    '                                                If PosRec = -1 Then
    '                                                    ArrayDataWord(9, CountArrayDataWord) = String.Format("{0:0.00}", Math.Round(off2, 2))
    '                                                Else
    '                                                    ArrayDataWord(9, PosRec) = String.Format("{0:0.00}", Math.Round(off2, 2))
    '                                                End If
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    '=======================================================================
    '                                    'находим блоки с левой и правой опорой
    '                                    '============================================================================
    '                                    'ищем левую опору
    '                                    If PosRec = -1 Then
    '                                        Dim ZLevel As Double = -1
    '                                        Dim ArrayDataPillars As String(,) = Nothing
    '                                        Dim colObj As ObjectIdCollection = New ObjectIdCollection
    '                                        Dim ptL As Point3d = New Point3d(-1, -1, 0)
    '                                        If pk1 < 0 Then
    '                                            ptL = New Point3d(x1, y1, 0)
    '                                        Else
    '                                            ptL = New Point3d(x2, y2, 0)
    '                                        End If
    '                                        'находим блок автокада
    '                                        Dim boolPt1 As Boolean = FuncFindBLKToPoint(ptL, colObj,, 0.05)
    '                                        If colObj.Count > 0 Then
    '                                            For j As Integer = 0 To colObj.Count - 1
    '                                                Dim blk As BlockReference = acTrans.GetObject(colObj.Item(j), OpenMode.ForRead)
    '                                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(blk, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                ZLevel = blk.Position.Z
    '                                                If boolReadPS = True Then
    '                                                    Exit For
    '                                                End If
    '                                            Next j
    '                                        Else
    '                                            Dim NumCogo As Integer() = Nothing
    '                                            Dim boolPt2 As Boolean = FuncFindCogoPoint(ptL.Convert2d(New Plane), NumCogo, 0.05)
    '                                            If IsArray(NumCogo) = True Then
    '                                                For j As Integer = 0 To NumCogo.Length - 1
    '                                                    Dim IDPtCogo As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumCogo(j))
    '                                                    Dim PtCogo As CogoPoint = acTrans.GetObject(IDPtCogo, OpenMode.ForRead)
    '                                                    Dim boolReadPS As Boolean = FuncReadDataSetProperties(PtCogo, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                    ZLevel = PtCogo.Elevation
    '                                                    If boolReadPS = True Then
    '                                                        Exit For
    '                                                    End If
    '                                                Next j
    '                                            End If
    '                                        End If
    '                                        '============================================================
    '                                        'характеристика найденной опоры
    '                                        Dim NumbPillars As String = ""
    '                                        Dim DownWire As Double = -1
    '                                        Dim TopWire As Double = -1
    '                                        Dim GRWire As Double = -1
    '                                        Dim VPillars As Double = -1
    '                                        Dim SVWire As Double = -1
    '                                        Dim MatPillars As String = ""
    '                                        If IsArray(ArrayDataPillars) = True Then
    '                                            For j As Integer = 0 To ArrayDataPillars.GetUpperBound(1)
    '                                                Dim NameF As String = ArrayDataPillars(1, j)
    '                                                Dim val As String = ArrayDataPillars(3, j)
    '                                                If NameF Like "PILLAR_NUMBER" Then 'номер опоры
    '                                                    NumbPillars = val
    '                                                ElseIf NameF Like "ZERO_LEVEL" Then 'отметка земли
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            ZLevel = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "DOWN_WIRE_LEVEL" Then 'отметка нижнего провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            DownWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "TOP_WIRE_LEVEL" Then 'отметка верхнего провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            TopWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "GROZ_WIRE_LEVEL" Then 'отметка грозотроса
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            GRWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "UP_PILLAR_LEVEL" Then 'отметка верха опоры
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            VPillars = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "DOWN_SV_LEVEL" Then 'отметка вспомогательного провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            SVWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "MATERIAL" Then 'материал опоры
    '                                                    MatPillars = val
    '                                                End If
    '                                            Next j
    '                                        End If
    '                                        '===========================================================================
    '                                        'заполняем массив далее
    '                                        'формируем строку с Материал опоры\Номер опоры\Расстояние
    '                                        Dim StrPillars As String = ""
    '                                        If MatPillars.Trim.Length > 0 Then
    '                                            If StrPillars.IndexOf("елезоб") Then
    '                                                StrPillars = "ж.б."
    '                                            ElseIf StrPillars.IndexOf("етал") Then
    '                                                StrPillars = "мет."
    '                                            ElseIf StrPillars.IndexOf("ерев") Then
    '                                                StrPillars = "дер."
    '                                            End If
    '                                        End If
    '                                        If NumbPillars.Trim.Length > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "№" & NumbPillars
    '                                            Else
    '                                                StrPillars = "№" & NumbPillars
    '                                            End If
    '                                        End If
    '                                        If ZLevel > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.земл.- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                            Else
    '                                                StrPillars = "Отм.земли- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                            End If
    '                                        End If
    '                                        If VPillars > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                            Else
    '                                                StrPillars = "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                            End If
    '                                        End If
    '                                        If DownWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                            End If
    '                                        End If
    '                                        If TopWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                            End If
    '                                        End If
    '                                        If GRWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.гр.троса - " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.гр.троса.- " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                            End If
    '                                        End If
    '                                        If SVWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                            End If
    '                                        End If
    '                                        'записываем характеристики левой опоры
    '                                        ArrayDataWord(11, CountArrayDataWord) = StrPillars
    '                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                        'ищем правую опору
    '                                        ZLevel = -1
    '                                        Erase ArrayDataPillars
    '                                        colObj = New ObjectIdCollection
    '                                        Dim ptP As Point3d = New Point3d(-1, -1, 0)
    '                                        If pk2 < 0 Then
    '                                            ptP = New Point3d(x2, y2, 0)
    '                                        Else
    '                                            ptP = New Point3d(x1, y1, 0)
    '                                        End If
    '                                        'находим блок автокада
    '                                        boolPt1 = FuncFindBLKToPoint(ptP, colObj,, 0.05)
    '                                        If colObj.Count > 0 Then
    '                                            For j As Integer = 0 To colObj.Count - 1
    '                                                Dim blk As BlockReference = acTrans.GetObject(colObj.Item(j), OpenMode.ForRead)
    '                                                Dim boolReadPS As Boolean = FuncReadDataSetProperties(blk, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                ZLevel = blk.Position.Z
    '                                                If boolReadPS = True Then
    '                                                    Exit For
    '                                                End If
    '                                            Next j
    '                                        Else
    '                                            Dim NumCogo As Integer() = Nothing
    '                                            Dim boolPt2 As Boolean = FuncFindCogoPoint(ptP.Convert2d(New Plane), NumCogo, 0.05)
    '                                            If IsArray(NumCogo) = True Then
    '                                                For j As Integer = 0 To NumCogo.Length - 1
    '                                                    Dim IDPtCogo As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumCogo(j))
    '                                                    Dim PtCogo As CogoPoint = acTrans.GetObject(IDPtCogo, OpenMode.ForRead)
    '                                                    Dim boolReadPS As Boolean = FuncReadDataSetProperties(PtCogo, ArrayDataPillars, "PILLARS_AIR_LINE")
    '                                                    ZLevel = PtCogo.Elevation
    '                                                    If boolReadPS = True Then
    '                                                        Exit For
    '                                                    End If
    '                                                Next j
    '                                            End If
    '                                        End If
    '                                        '============================================================
    '                                        'характеристика найденной опоры
    '                                        NumbPillars = ""
    '                                        DownWire = -1
    '                                        TopWire = -1
    '                                        GRWire = -1
    '                                        VPillars = -1
    '                                        SVWire = -1
    '                                        MatPillars = ""
    '                                        If IsArray(ArrayDataPillars) = True Then
    '                                            For j As Integer = 0 To ArrayDataPillars.GetUpperBound(1)
    '                                                Dim NameF As String = ArrayDataPillars(1, j)
    '                                                Dim val As String = ArrayDataPillars(3, j)
    '                                                If NameF Like "PILLAR_NUMBER" Then 'номер опоры
    '                                                    NumbPillars = val
    '                                                ElseIf NameF Like "ZERO_LEVEL" Then 'отметка земли
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            ZLevel = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "DOWN_WIRE_LEVEL" Then 'отметка нижнего провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            DownWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "TOP_WIRE_LEVEL" Then 'отметка верхнего провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            TopWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "GROZ_WIRE_LEVEL" Then 'отметка грозотроса
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            GRWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "UP_PILLAR_LEVEL" Then 'отметка верха опоры
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            VPillars = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "DOWN_SV_LEVEL" Then 'отметка вспомогательного провода
    '                                                    If IsNumeric(val) = True Then
    '                                                        If CDbl(val) > 0 Then
    '                                                            SVWire = CDbl(val)
    '                                                        End If
    '                                                    End If
    '                                                ElseIf NameF Like "MATERIAL" Then 'материал опоры
    '                                                    MatPillars = val
    '                                                End If
    '                                            Next j
    '                                        End If
    '                                        '===========================================================================
    '                                        'заполняем массив далее
    '                                        'формируем строку с Материал опоры\Номер опоры\Расстояние
    '                                        StrPillars = ""
    '                                        If MatPillars.Trim.Length > 0 Then
    '                                            If StrPillars.IndexOf("елезоб") Then
    '                                                StrPillars = "ж.б."
    '                                            ElseIf StrPillars.IndexOf("етал") Then
    '                                                StrPillars = "мет."
    '                                            ElseIf StrPillars.IndexOf("ерев") Then
    '                                                StrPillars = "дер."
    '                                            End If
    '                                        End If
    '                                        If NumbPillars.Trim.Length > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "№" & NumbPillars
    '                                            Else
    '                                                StrPillars = "№" & NumbPillars
    '                                            End If
    '                                        End If
    '                                        If ZLevel > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.земл.- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                            Else
    '                                                StrPillars = "Отм.земли- " & String.Format("{0:0.00}", Math.Round(ZLevel, 2))
    '                                            End If
    '                                        End If
    '                                        If VPillars > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                            Else
    '                                                StrPillars = "Отм.верха.- " & String.Format("{0:0.00}", Math.Round(VPillars, 2))
    '                                            End If
    '                                        End If
    '                                        If DownWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.н.пр.- " & String.Format("{0:0.00}", Math.Round(DownWire, 2))
    '                                            End If
    '                                        End If
    '                                        If TopWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.в.пр.- " & String.Format("{0:0.00}", Math.Round(TopWire, 2))
    '                                            End If
    '                                        End If
    '                                        If GRWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.гр.троса - " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.гр.троса.- " & String.Format("{0:0.00}", Math.Round(GRWire, 2))
    '                                            End If
    '                                        End If
    '                                        If SVWire > 0 Then
    '                                            If StrPillars.Trim.Length > 0 Then
    '                                                StrPillars = StrPillars & vbLf & "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                            Else
    '                                                StrPillars = "Отм.каб.связи - " & String.Format("{0:0.00}", Math.Round(SVWire, 2))
    '                                            End If
    '                                        End If
    '                                        'записываем характеристики правой опоры
    '                                        ArrayDataWord(12, CountArrayDataWord) = StrPillars
    '                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                                        'записывает провис провода и температуру
    '                                        LenghtSegment = Math.Round(acLine.Length, 3)
    '                                        Dim ptTempPer As Point3d = New Point3d(-1, -1, -1)
    '                                        If SlakWires <= 0 Then
    '                                            If IsArray(ArrayPrWire) = True Then
    '                                                For k As Integer = 0 To ArrayPrWire.GetUpperBound(1)
    '                                                    Dim ptTempPr As Point3d = New Point3d(ArrayPrWire(0, k), ArrayPrWire(1, k), 0)
    '                                                    Dim ptTemp As Point3d = acLine.GetClosestPointTo(ptTempPr, True)
    '                                                    Dim distPr As Double = ptTemp.GetVectorTo(ptTempPr).Length
    '                                                    If distPr < 0.5 Then
    '                                                        SlakWires = ArrayPrWire(2, k)
    '                                                        ptTempPer = ptTemp
    '                                                        Exit For
    '                                                    End If
    '                                                Next k
    '                                            End If
    '                                        End If
    '                                        If DownLewel <= 0 Then
    '                                            If SlakWires >= 0 Then
    '                                                If ptTempPer.X <> -1 And ptTempPer.Y <> -1 Then
    '                                                    '========================================================================================
    '                                                    'находим поверхность рабочую
    '                                                    Dim elev As Double = 0
    '                                                    Dim ColSurv As ObjectIdCollection = CivilDoc.GetSurfaceIds
    '                                                    For Each idSurf As ObjectId In ColSurv
    '                                                        Dim surf As Autodesk.Civil.DatabaseServices.Surface = acTrans.GetObject(idSurf, OpenMode.ForRead)
    '                                                        Try
    '                                                            elev = surf.FindElevationAtXY(ptTempPer.X, ptTempPer.Y)
    '                                                            DownLewel = elev + SlakWires
    '                                                            Exit For
    '                                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                                        End Try
    '                                                    Next
    '                                                End If
    '                                            End If
    '                                        End If
    '                                        '=================================================================================
    '                                        Dim strMarkWire As String = Mat
    '                                        If SlakWires <> 0 Then
    '                                            strMarkWire = strMarkWire & vbLf & Math.Round(SlakWires, 2)
    '                                        End If
    '                                        ArrayDataWord(13, CountArrayDataWord) = strMarkWire
    '                                        '================================================================================
    '                                        'дата и температура воздуха
    '                                        Dim StrDataTemp As String = DataSurv
    '                                        If Temp.Trim.Length > 0 Then
    '                                            StrDataTemp = StrDataTemp & vbLf & Temp & Chr(176) & "С"
    '                                        End If
    '                                        ArrayDataWord(14, CountArrayDataWord) = StrDataTemp
    '                                        '=================================================================================
    '                                        'владелец
    '                                        ArrayDataWord(15, CountArrayDataWord) = Owner
    '                                        '==================================================================================
    '                                        'примечание
    '                                        ArrayDataWord(16, CountArrayDataWord) = Note
    '                                        '==================================================================================
    '                                        'длина сегмента
    '                                        ArrayDataWord(17, CountArrayDataWord) = LenghtSegment
    '                                        CountArrayDataWord += 1
    '                                        '=============================================================================================
    '                                        'вставляем точки сого
    '                                        'находим высоту точки
    '                                        'удаляем предыдущие созданные точки
    '                                        Dim NumbPoint As Integer() = Nothing
    '                                        If TypePer = 0 Then
    '                                            Dim boolIns As Boolean = FuncFindCogoPoint(pt.Convert2d(New Plane), NumbPoint, 0.02, "Воздушная ЛЭП")
    '                                        ElseIf TypePer = 1 Then
    '                                            Dim boolIns As Boolean = FuncFindCogoPoint(pt.Convert2d(New Plane), NumbPoint, 0.02, "Воздушная линия связи")
    '                                        End If
    '                                        If IsArray(NumbPoint) = True Then
    '                                            For k As Integer = 0 To NumbPoint.Length - 1
    '                                                Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '                                                Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '                                                If IsDBNull(PointCogoDel) = False Then
    '                                                    PointCogoDel.Erase()
    '                                                End If
    '                                            Next k
    '                                        End If
    '                                        '===================================================================================
    '                                        'создаем COGO - точку
    '                                        'генерируем номер для cogo точки
    '                                        Dim numbCogo As Long = 0
    '                                        For i1 As Integer = 0 To 100
    '                                            Dim rng As Random = New Random
    '                                            numbCogo = rng.Next(1000500, 1000999)
    '                                            If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '                                                Exit For
    '                                            End If
    '                                        Next
    '                                        'создаем новую точку
    '                                        Dim idCogo As ObjectId = Nothing
    '                                        If TypePer = 0 Then
    '                                            idCogo = FuncCreateCOGOPoint(pt, CogoDesk, numbCogo, DownLewel,, "Воздушная ЛЭП")
    '                                        ElseIf TypePer = 1 Then
    '                                            idCogo = FuncCreateCOGOPoint(pt, CogoDesk, numbCogo, DownLewel,, "Воздушная линия связи")
    '                                        End If
    '                                        Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '                                        'присваиваем набор данных
    '                                        Dim ArrayODRec As String(,) = Nothing
    '                                        ReDim Preserve ArrayODRec(2, 1)
    '                                        ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '                                        ArrayODRec(1, 0) = "LABEL"
    '                                        ArrayODRec(2, 0) = "Указатель направления автодорог"
    '                                        ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '                                        ArrayODRec(1, 1) = "DESK"
    '                                        ArrayODRec(2, 1) = CogoDesk
    '                                        If IsNothing(CreateTBLPSID) = False Then
    '                                            'присваиваем таблицу объекту
    '                                            Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '                                            'записываем данные в таблицу
    '                                            Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '                                        End If
    '                                    End If
    '                                Next i
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '                If IsDBNull(PlineAlignEx) = False Then
    '                    If PlineAlignEx.IsWriteEnabled = False Then
    '                        PlineAlignEx.UpgradeOpen()
    '                    End If
    '                    PlineAlignEx.Erase()
    '                End If
    '                If IsDBNull(PlineAlignProject) = False Then
    '                    If PlineAlignProject.IsWriteEnabled = False Then
    '                        PlineAlignProject.UpgradeOpen()
    '                    End If
    '                    PlineAlignProject.Erase()
    '                End If
    '            End If
    '            acTrans.Commit()
    '        End Using
    '        locker.Dispose()
    '        '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-угол пересечения, 6,7-расстояние до правой и левой опоры сущ.трассы, 8,9-расстояние до правой и левой опоры пр.трассы, 10-количество проводов, 11,12-характеристика опор, 13-Марка провода и провис 14-Владелец, 15-Адрес, 16-Примечание
    '        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        'записываем все в массив
    '        If IsArray(ArrayDataWord) = True Then
    '            'сортируем массив
    '            If ArrayDataWord.GetUpperBound(1) > 0 Then
    '                Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 0)
    '            End If
    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '            'выбираем документ Word
    '            Try
    '                WApp = CreateObject("Word.Application")
    '                WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пересечения воздушных линий электропередач и связи.docx")
    '                Dim rng As Random = New Random
    '                Dim Numbrng As Integer = rng.Next
    '                WDoc.SaveAs2(TempDir & "Ведомость пересечения воздушных линий электропередач и связи_" & Numbrng & ".docx")

    '            Catch
    '                MsgBox("Не удалось найти выбранную таблицу!!!")
    '                WDoc = Nothing
    '                WApp = Nothing
    '                Exit Sub
    '            End Try
    '            For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '                For Each tabl As Word.Table In WDoc.Tables
    '                    Dim countTableword As Integer = tabl.Rows.Count
    '                    If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
    '                        Exit For
    '                    End If
    '                    tabl.Rows.Add()
    '                    '===================================================================================
    '                    'название объекта
    '                    If IsNothing(ArrayDataWord(1, i)) = False Then
    '                        tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(1, i)
    '                    Else
    '                        tabl.Cell(countTableword, 1).Range.Text = "-"
    '                    End If
    '                    '===================================================================================
    '                    'количество проводов
    '                    If IsNothing(ArrayDataWord(10, i)) = False Then
    '                        tabl.Cell(countTableword, 2).Range.Text = ArrayDataWord(10, i)
    '                    Else
    '                        tabl.Cell(countTableword, 2).Range.Text = "-"
    '                    End If
    '                    '===================================================================================
    '                    'километр пересечения
    '                    If IsNothing(ArrayDataWord(2, i)) = False Then
    '                        tabl.Cell(countTableword, 3).Range.Text = ArrayDataWord(2, i)
    '                    Else
    '                        tabl.Cell(countTableword, 3).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем угол пересечения
    '                    If IsNothing(ArrayDataWord(4, i)) = False Then
    '                        tabl.Cell(countTableword, 4).Range.Text = ArrayDataWord(4, i)
    '                    Else
    '                        tabl.Cell(countTableword, 4).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем расстояние от оси лево
    '                    If IsNothing(ArrayDataWord(6, i)) = False Then
    '                        tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(6, i)
    '                    Else
    '                        tabl.Cell(countTableword, 5).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем расстояние от оси право
    '                    If IsNothing(ArrayDataWord(7, i)) = False Then
    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(7, i)
    '                    Else
    '                        tabl.Cell(countTableword, 6).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем пикет пересечения
    '                    If IsNothing(ArrayDataWord(3, i)) = False Then
    '                        tabl.Cell(countTableword, 7).Range.Text = ArrayDataWord(3, i)
    '                    Else
    '                        tabl.Cell(countTableword, 7).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем угол пересечения
    '                    If IsNothing(ArrayDataWord(5, i)) = False Then
    '                        tabl.Cell(countTableword, 8).Range.Text = ArrayDataWord(5, i)
    '                    Else
    '                        tabl.Cell(countTableword, 8).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем расстояние смещения лево
    '                    If IsNothing(ArrayDataWord(8, i)) = False Then
    '                        tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(8, i)
    '                    Else
    '                        tabl.Cell(countTableword, 9).Range.Text = "-"
    '                    End If
    '                    '=================================================================================================
    '                    'записываем расстояние смещения право
    '                    If IsNothing(ArrayDataWord(9, i)) = False Then
    '                        tabl.Cell(countTableword, 10).Range.Text = ArrayDataWord(9, i)
    '                    Else
    '                        tabl.Cell(countTableword, 10).Range.Text = "-"
    '                    End If
    '                    '===================================================================================================
    '                    'характеристики левой опоры
    '                    If IsNothing(ArrayDataWord(11, i)) = False Then
    '                        tabl.Cell(countTableword, 11).Range.Text = ArrayDataWord(11, i)
    '                    Else
    '                        tabl.Cell(countTableword, 11).Range.Text = "-"
    '                    End If
    '                    '==================================================================================================
    '                    'характеристики правой опоры
    '                    If IsNothing(ArrayDataWord(12, i)) = False Then
    '                        tabl.Cell(countTableword, 12).Range.Text = ArrayDataWord(12, i)
    '                    Else
    '                        tabl.Cell(countTableword, 12).Range.Text = "-"
    '                    End If
    '                    '==================================================================================================
    '                    'марка провода
    '                    If IsNothing(ArrayDataWord(13, i)) = False Then
    '                        tabl.Cell(countTableword, 13).Range.Text = ArrayDataWord(13, i)
    '                    Else
    '                        tabl.Cell(countTableword, 13).Range.Text = "-"
    '                    End If
    '                    '==================================================================================================
    '                    'владелец
    '                    If IsNothing(ArrayDataWord(15, i)) = False Then
    '                        tabl.Cell(countTableword, 14).Range.Text = ArrayDataWord(15, i)
    '                    Else
    '                        tabl.Cell(countTableword, 14).Range.Text = "-"
    '                    End If
    '                    '==================================================================================================
    '                    'температура
    '                    If IsNothing(ArrayDataWord(14, i)) = False Then
    '                        tabl.Cell(countTableword, 15).Range.Text = ArrayDataWord(14, i)
    '                    Else
    '                        tabl.Cell(countTableword, 15).Range.Text = "-"
    '                    End If
    '                    '==================================================================================================
    '                    'примечание
    '                    If IsNothing(ArrayDataWord(16, i)) = False Then
    '                        tabl.Cell(countTableword, 16).Range.Text = ArrayDataWord(16, i)
    '                    Else
    '                        tabl.Cell(countTableword, 16).Range.Text = "-"
    '                    End If

    '                    Exit For
    '                Next
    '            Next i
    '            WApp.Visible = True
    '            WDoc = Nothing
    '            WApp = Nothing
    '        End If
    '    End Sub
    '    '==================================================================================================
    '    'формирование дефектной ведомости колодцев подземных коммуникаций 
    '    <CommandMethod("STPR-VEDKS")>
    '    Public Sub DefVedKS()
    '        Try
    '            Dim WApp As Word.Application = Nothing
    '            Dim WDoc As Word.Document = Nothing
    '            Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '            Dim acCurDb As Database = acDoc.Database
    '            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '            Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
    '            Dim resOpt As PromptEntityResult = Nothing
    '            Dim LUKPoints As String(,) = Nothing 'контейнер для точек в отчет
    '            Dim CountLUKPoints As Integer = 0 'счетчик
    '            Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '            Using tr
    '                '=====================================================================================================================
    '                '1. выбираем слои для отчета
    '                Dim GroupPointList1 As GroupPointList2 = New GroupPointList2
    '                Dim CountRow As Integer = 0
    '                Dim ArrayLayer As String() = Nothing
    '                Dim boolListLayer As Integer = ListLayerGeo2(ArrayLayer)
    '                If IsArray(ArrayLayer) = True Then
    '                    'заполняем форму
    '                    For i As Integer = 0 To ArrayLayer.GetUpperBound(0)
    '                        Dim NameLayer As String = ArrayLayer(i)
    '                        If NameLayer.IndexOf("ИИ-ТОПО-") > -1 Then
    '                            GroupPointList1.DataGridView1.Rows.Insert(CountRow)
    '                            GroupPointList1.DataGridView1.Rows(CountRow).Cells(1).Value = NameLayer
    '                            CountRow += 1
    '                        End If
    '                    Next i
    '                End If
    '                '=================================================================================================
    '                'выводим диалоговое окно
    '                GroupPointList1.ShowDialog()
    '                If GroupPointList1.stboolGroup = False Then
    '                    GoTo line1
    '                End If
    '                '======================================================================================================
    '                'забираем точки COGO в коллекцию
    '                Dim PointCOGO1 As CogoPoint = Nothing
    '                '3. находим расширенные ранные
    '                Const UDPname1 As String = "MANHOLE_NUMBER"
    '                Const UDPname3 As String = "MANHOLE_LAND_LEVEL"
    '                Const UDPname4 As String = "MANHOLE_BLOCK_TOP_LEVEL"
    '                Const UDPname5 As String = "MANHOLE_BLOCK_DOWN_LEVEL"
    '                Const UDPname6 As String = "MANHOLE_DEPTH_LEVEL"
    '                Dim UDP1 As UDPString = FindUDPStr(UDPname1)
    '                Dim UDP3 As UDPDouble = FindUDP(UDPname3)
    '                Dim UDP4 As UDPDouble = FindUDP(UDPname4)
    '                Dim UDP5 As UDPDouble = FindUDP(UDPname5)
    '                Dim UDP6 As UDPDouble = FindUDP(UDPname6)
    '                For Each Point1 As ObjectId In CivilDoc.CogoPoints
    '                    PointCOGO1 = Point1.GetObject(OpenMode.ForRead)
    '                    Dim labelStyleId As ObjectId = PointCOGO1.LabelStyleId
    '                    If labelStyleId.IsNull = False Then
    '                        Dim labelStyle As LabelStyle = TryCast(labelStyleId.GetObject(OpenMode.ForWrite), LabelStyle)
    '                        If labelStyle.Name Like "ИИ-Колодцы подземных коммуникаций" Then
    '                            For i As Integer = 0 To GroupPointList1.DataGridView1.RowCount - 1
    '                                Dim boolSelect As Boolean = GroupPointList1.DataGridView1.Rows(i).Cells(0).Value
    '                                If boolSelect = True Then
    '                                    Dim NameLayerSelect As String = GroupPointList1.DataGridView1.Rows(i).Cells(1).Value
    '                                    If NameLayerSelect Like PointCOGO1.Layer Then
    '                                        'находим пользовательские свойства
    '                                        Dim nameUserPoint1 As String = PointCOGO1.GetUDPValue(UDP1) 'номер точки
    '                                        Dim nameUserPoint3 As Double = PointCOGO1.GetUDPValue(UDP3) 'отметка земли
    '                                        Dim nameUserPoint4 As Double = PointCOGO1.GetUDPValue(UDP4) 'отметка верха прокладки
    '                                        Dim nameUserPoint5 As Double = PointCOGO1.GetUDPValue(UDP5) 'отметка низа прокладки
    '                                        Dim nameUserPoint6 As Double = PointCOGO1.GetUDPValue(UDP6) 'дно колодца
    '                                        Dim DeptManhole As Double = Math.Round(PointCOGO1.Elevation - nameUserPoint6, 2) 'дно колодца
    '                                        If nameUserPoint6 = -1 Then
    '                                            DeptManhole = 0
    '                                        End If
    '                                        Dim TopManhole As Double = Math.Round(PointCOGO1.Elevation - nameUserPoint4, 2) 'верх прокладки
    '                                        If nameUserPoint4 = -1 Then
    '                                            TopManhole = 0
    '                                        End If
    '                                        Dim DownManhole As Double = Math.Round(PointCOGO1.Elevation - nameUserPoint5, 2) 'низ прокладки
    '                                        If nameUserPoint5 = -1 Then
    '                                            DownManhole = 0
    '                                        End If
    '                                        Dim LandManhole As Double = Math.Round(PointCOGO1.Elevation - nameUserPoint5, 2) 'отметика земли (мощения)
    '                                        If nameUserPoint3 = -1 Then
    '                                            LandManhole = Math.Round(PointCOGO1.Elevation, 2)
    '                                        End If
    '                                        'считываем расширенные данные
    '                                        'Dim ArrayOD As String() = Nothing
    '                                        'Dim boolReadOD As Boolean = FuncReadOD2(ArrayOD, PointCOGO1, "MANHOLES")
    '                                        ReDim Preserve LUKPoints(14, CountLUKPoints)
    '                                        LUKPoints(0, CountLUKPoints) = nameUserPoint1
    '                                        LUKPoints(1, CountLUKPoints) = Math.Round(PointCOGO1.Northing, 2)
    '                                        LUKPoints(2, CountLUKPoints) = Math.Round(PointCOGO1.Easting, 2)
    '                                        LUKPoints(3, CountLUKPoints) = LandManhole 'отметка земли
    '                                        LUKPoints(4, CountLUKPoints) = Math.Round(PointCOGO1.Elevation, 2)
    '                                        If DeptManhole = 0 Then
    '                                            LUKPoints(5, CountLUKPoints) = "-" 'отметка дна
    '                                        Else
    '                                            LUKPoints(5, CountLUKPoints) = DeptManhole 'отметка земли
    '                                        End If
    '                                        If TopManhole = 0 Then
    '                                            LUKPoints(6, CountLUKPoints) = "-"
    '                                        Else
    '                                            LUKPoints(6, CountLUKPoints) = TopManhole 'верх прокладки
    '                                        End If
    '                                        If DownManhole = 0 Then
    '                                            LUKPoints(7, CountLUKPoints) = "-" 'низ прокладки
    '                                        Else
    '                                            LUKPoints(7, CountLUKPoints) = DownManhole 'низ прокладки
    '                                        End If

    '                                        'If IsArray(ArrayOD) = True Then
    '                                        '    LUKPoints(8, CountLUKPoints) = ArrayOD(9)
    '                                        'Else
    '                                        '    LUKPoints(8, CountLUKPoints) = ""
    '                                        'End If
    '                                        CountLUKPoints += 1
    '                                    End If
    '                                End If
    '                            Next i
    '                        End If
    '                    End If
    '                Next
    '                '========================================================================================================================
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                'обрабатываем блоки автокад
    '                '==============================================================================
    '                'формируем фильтр
    '                Dim BlkSelect As PromptSelectionResult
    '                Dim ssetBlk As SelectionSet = Nothing
    '                Dim ValType1(2) As TypedValue
    '                ValType1.SetValue(New TypedValue(DxfCode.Start, "INSERT"), 0)
    '                ValType1.SetValue(New TypedValue(DxfCode.Visibility, 0), 1)
    '                ValType1.SetValue(New TypedValue(67, 0), 2)
    '                Dim FilterParcel As SelectionFilter = New SelectionFilter(ValType1)
    '                BlkSelect = m_Editor.SelectAll(FilterParcel)
    '                'смотрим что попалось
    '                If BlkSelect.Status = PromptStatus.OK Then
    '                    ssetBlk = BlkSelect.Value
    '                    Dim countObj As Integer = ssetBlk.Count
    '                    If countObj = 0 Then GoTo line4
    '                    For Each idobj As ObjectId In ssetBlk.GetObjectIds
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idobj, OpenMode.ForRead)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is BlockReference Then
    '                                Dim obj As BlockReference = acEnt
    '                                Dim LayerBlk As String = obj.Layer
    '                                For i As Integer = 0 To GroupPointList1.DataGridView1.RowCount - 1
    '                                    Dim boolSelect As Boolean = GroupPointList1.DataGridView1.Rows(i).Cells(0).Value
    '                                    If boolSelect = True Then
    '                                        Dim NameLayerSelect As String = GroupPointList1.DataGridView1.Rows(i).Cells(1).Value
    '                                        If NameLayerSelect Like LayerBlk Then
    '                                            'читаем атрибуты
    '                                            Dim ArrayAtrib As String(,) = Nothing
    '                                            Dim boolReadAttrib As Boolean = FuncReadAllBlockAttr(obj, ArrayAtrib)
    '                                            Dim BLKNumb As String = "-"
    '                                            Dim BlkLand As String = "-"
    '                                            Dim BlkTop As String = "-"
    '                                            Dim BlkDown As String = "-"
    '                                            Dim BlkDept As String = "-"
    '                                            Dim BlkOd As String = "-"
    '                                            If IsArray(ArrayAtrib) = True Then
    '                                                For j As Integer = 0 To ArrayAtrib.GetUpperBound(1)
    '                                                    If ArrayAtrib(0, j) Like "MANHOLE_NUMBER" Then
    '                                                        BLKNumb = ArrayAtrib(1, j)
    '                                                    ElseIf ArrayAtrib(0, j) Like "LAND_LEVEL" Then
    '                                                        BlkLand = ArrayAtrib(1, j)
    '                                                    ElseIf ArrayAtrib(0, j) Like "BLOCK_TOP_LEVEL" Then
    '                                                        BlkTop = ArrayAtrib(1, j)
    '                                                    ElseIf ArrayAtrib(0, j) Like "BLOCK_DOWN_LEVEL" Then
    '                                                        BlkDown = ArrayAtrib(1, j)
    '                                                    ElseIf ArrayAtrib(0, j) Like "DEPTH_LEVEL" Then
    '                                                        BlkDept = ArrayAtrib(1, j)
    '                                                    ElseIf ArrayAtrib(0, j) Like "OD" Then
    '                                                        BlkOd = ArrayAtrib(1, j)
    '                                                    End If
    '                                                Next j
    '                                                '==================================================================================
    '                                                'записываем данные в массив
    '                                                If BlkOd Like "MANHOLES" Then
    '                                                    'считываем расширенные данные
    '                                                    Dim ArrayOD As String() = Nothing
    '                                                    Dim boolReadOD As Boolean = FuncReadOD2(ArrayOD, obj, "MANHOLES")
    '                                                    ReDim Preserve LUKPoints(14, CountLUKPoints)
    '                                                    LUKPoints(0, CountLUKPoints) = BLKNumb
    '                                                    LUKPoints(1, CountLUKPoints) = Math.Round(obj.Position.Y, 2)
    '                                                    LUKPoints(2, CountLUKPoints) = Math.Round(obj.Position.X, 2)
    '                                                    LUKPoints(3, CountLUKPoints) = BlkLand
    '                                                    LUKPoints(4, CountLUKPoints) = Math.Round(obj.Position.Z, 2)
    '                                                    LUKPoints(5, CountLUKPoints) = BlkDept
    '                                                    LUKPoints(6, CountLUKPoints) = BlkTop
    '                                                    LUKPoints(7, CountLUKPoints) = BlkDown
    '                                                    If IsArray(ArrayOD) = True Then
    '                                                        LUKPoints(8, CountLUKPoints) = ArrayOD(9)
    '                                                    Else
    '                                                        LUKPoints(8, CountLUKPoints) = ""
    '                                                    End If
    '                                                    CountLUKPoints += 1
    '                                                End If
    '                                            End If
    '                                        End If
    '                                    End If
    '                                Next i
    '                            End If
    '                        End If
    '                    Next
    '                End If
    '                'сортируем массив
    '                'делаем сортировку массива по пикетажу
    'line4:
    '                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '                '=====================================================================================================================
    '                'полилинии или окружности
    '                '=====================================================================================================================
    '                If IsArray(LUKPoints) = True Then
    '                    Dim Obj3dCollection As DBObjectCollection = New DBObjectCollection
    '                    Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) 'открываем таблицу блоков на чтение
    '                    Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '                    For Each id As ObjectId In btr 'цикл для перебора всех объектов на чертеже
    '                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
    '                        If Not IsDBNull(ent) Then
    '                            Dim NameLineStr As String = ent.Linetype
    '                            If NameLineStr Like "A-Камера подземная на трубопроводах" Or NameLineStr Like "A-Камера наземная на трубопроводах" Then
    '                                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Or TypeOf ent Is Circle Then
    '                                    Obj3dCollection.Add(ent)
    '                                End If
    '                            End If
    '                        End If
    '                    Next
    '                    'перебираем циклом коллекцию
    '                    For Each acEnt As Autodesk.AutoCAD.DatabaseServices.DBObject In Obj3dCollection
    '                        For i As Integer = 0 To LUKPoints.GetUpperBound(1)
    '                            Dim Razmer As String = LUKPoints(8, i)
    '                            If Razmer.Trim.Length = 0 Then
    '                                Dim PT2d As Point2d = New Point2d(Val(LUKPoints(2, i)), Val(LUKPoints(1, i)))
    '                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '                                    Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
    '                                    'смотрим на принадлежнотсь его точке
    '                                    Dim boolCrossPt As Boolean = FuncCrossPoint(Pline, PT2d)
    '                                    If boolCrossPt = True Then
    '                                        Dim a As Double = 0 'длина
    '                                        Dim b As Double = 1000000 'ширина
    '                                        If Pline.NumberOfVertices > 2 Then
    '                                            For i1 As Integer = 0 To Pline.NumberOfVertices - 2
    '                                                Dim Pt1 As Point2d = Pline.GetPoint2dAt(i1)
    '                                                Dim Pt2 As Point2d = Pline.GetPoint2dAt(i1 + 1)
    '                                                Dim vect As Vector2d = Pt1.GetVectorTo(Pt2)
    '                                                If vect.Length > a Then a = vect.Length
    '                                                If vect.Length < b Then b = vect.Length
    '                                            Next i1
    '                                            Dim Pt1k As Point2d = Pline.GetPoint2dAt(0)
    '                                            Dim Pt2k As Point2d = Pline.GetPoint2dAt(Pline.NumberOfVertices - 1)
    '                                            Dim vectk As Vector2d = Pt1k.GetVectorTo(Pt2k)
    '                                            If vectk.Length > a Then a = vectk.Length
    '                                            If vectk.Length < b Then b = vectk.Length
    '                                            'записываем размер колодца
    '                                            LUKPoints(8, i) = Math.Round(a, 2) & "х" & Math.Round(b, 2)
    '                                            GoTo line6
    '                                        End If
    '                                    End If
    '                                ElseIf TypeOf acEnt Is Circle Then
    '                                    Dim circ As Circle = acEnt
    '                                    Dim PtCentr As Point3d = circ.Center
    '                                    Dim PtCentr2d As Point2d = New Point2d(PtCentr.X, PtCentr.Y)
    '                                    Dim vect As Vector2d = PT2d.GetVectorTo(PtCentr2d)
    '                                    If vect.Length < circ.Radius Then
    '                                        'записываем диаметр колодца
    '                                        LUKPoints(8, i) = Math.Round(circ.Diameter, 2)
    '                                        GoTo line6
    '                                    End If
    '                                End If
    '                            End If
    '                        Next i
    'line6:
    '                    Next
    '                End If

    '                'проверяем наличие в них люка из массива
    'line5:

    '                Dim sortArray As Boolean = FuncSortArrayStr(LUKPoints, 0)
    '                'запускаем Ворд
    '                'проверяем есть ли такой шаблон
    '                Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '                Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '                'отделяем имя файла от папки
    '                Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '                'определяем имя родительской папки для папки где находится сборка
    '                Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '                'ищем пути для таблиц
    '                Dim PutchTbl As String = "k\STPRGeoTools\TemplateTables\"
    '                If File.Exists(parentPatch2 & "\TemplateTables\Экспликация смотровых колодцев подземных коммуникаций.docx") = True Then
    '                    PutchTbl = parentPatch2 & "\TemplateTables\"
    '                Else
    '                    If File.Exists("k\STPRGeoTools\TemplateTables\Экспликация смотровых колодцев подземных коммуникаций.docx") = False Then
    '                        MsgBox("Шаблон не найден! Макро прервано.")
    '                        Exit Sub
    '                    End If
    '                End If
    '                Try
    '                    WApp = CreateObject("Word.Application")
    '                    WDoc = WApp.Documents.Open(PutchTbl & "Экспликация смотровых колодцев подземных коммуникаций.docx")
    '                Catch
    '                    MsgBox("Не удалось найти выбранную таблицу!!!")
    '                End Try
    '                WApp.Visible = True
    '                WApp.Dialogs(Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFileSaveAs).Show()
    '                Dim fName As String = WDoc.Name
    '                Dim CountRows As Integer = 1
    '                If IsArray(LUKPoints) = True Then
    '                    'делаем сортировку массива по пикетажу
    '                    Dim tblCount As Integer = WDoc.Tables.Count
    '                    For Each tabl As Word.Table In WDoc.Tables
    '                        Dim countTableword As Integer = tabl.Rows.Count
    '                        For i As Integer = 0 To LUKPoints.GetUpperBound(1)
    '                            tabl.Rows.Add()
    '                            countTableword = countTableword + i
    '                            tabl.Cell(countTableword, 1).Range.Text = LUKPoints(0, i) 'номер люка
    '                            'координата х
    '                            tabl.Cell(countTableword, 2).Range.Text = LUKPoints(1, i)
    '                            'координата y
    '                            tabl.Cell(countTableword, 3).Range.Text = LUKPoints(2, i)
    '                            'отметка земли
    '                            tabl.Cell(countTableword, 4).Range.Text = LUKPoints(3, i)
    '                            'высота крышки люка
    '                            tabl.Cell(countTableword, 5).Range.Text = LUKPoints(4, i)
    '                            'отметка дна
    '                            tabl.Cell(countTableword, 6).Range.Text = LUKPoints(5, i)
    '                            'отметка верха прокладки
    '                            tabl.Cell(countTableword, 7).Range.Text = LUKPoints(6, i)
    '                            'отметка низа прокладки
    '                            tabl.Cell(countTableword, 8).Range.Text = LUKPoints(7, i)
    '                            'размер колодца
    '                            tabl.Cell(countTableword, 9).Range.Text = LUKPoints(8, i)
    '                            CountRows += 1
    '                        Next i
    '                        Exit For
    '                    Next
    '                End If
    '                tr.Commit()
    '            End Using
    'line1:
    '            locker.Dispose()
    '            WDoc = Nothing
    '            WApp = Nothing
    '        Catch
    '        End Try
    '    End Sub
    '    '==================================================================================================
    '    'Ведомость владельцев инженерных коммуникаций
    '    '<CommandMethod("STPR-VedOwnerNetwork")>
    '    'Public Sub VedOwnerNetwork()
    '    '    Dim WApp As Word.Application = Nothing
    '    '    Dim WDoc As Word.Document = Nothing
    '    '    Dim ArrayDataWord As String(,) = Nothing
    '    '    Dim CountArrayDataWord As Integer = 0
    '    '    Dim boolRez As Boolean = False
    '    '    Dim AngleEX As Double(,) = Nothing
    '    '    Dim countAngleEx As Integer = 0
    '    '    Dim AngleProject As Double(,) = Nothing
    '    '    Dim countAngleProject As Integer = 0
    '    '    Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '    '    'открываем базу данных чертежа
    '    '    Dim ed As Editor = acDoc.Editor
    '    '    Dim db As Database = acDoc.Database
    '    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    '    'ищем пути для таблиц
    '    '    Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
    '    '    Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
    '    '    'отделяем имя файла от папки
    '    '    Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
    '    '    'определяем имя родительской папки для папки где находится сборка
    '    '    Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
    '    '    Dim PutchTbl As String = "k:\STPRGeoTools\TemplateTables\Geo_Survey\"
    '    '    If File.Exists(parentPatch2 & "\TemplateTables\Geo_Survey\Ведомость инженерных коммуникаций.docx") = True Then
    '    '        PutchTbl = parentPatch2 & "\TemplateTables\Geo_Survey\"
    '    '    Else
    '    '        If File.Exists("k:\STPRGeoTools\TemplateTables\Geo_Survey\Ведомость инженерных коммуникаций.docx") = False Then
    '    '            MsgBox("Шаблон не найден! Макро прервано.")
    '    '            Exit Sub
    '    '        End If
    '    '    End If
    '    '    'временный каталог для автосохранения таблицы
    '    '    Dim TempDir As String = ""
    '    '    Try
    '    '        TempDir = System.IO.Path.GetTempPath()
    '    '    Catch ex As System.Security.SecurityException
    '    '        TempDir = strGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения", 16384)
    '    '    End Try
    '    '    Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
    '    '    Using acTrans As Transaction = db.TransactionManager.StartTransaction()
    '    '        'находим ось существующей трассы
    '    '        Dim AlignProject As Alignment = Nothing
    '    '        Dim AlignEX As Alignment = Nothing
    '    '        Dim PlineAlignEx As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '    '        Dim PlineAlignProject As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '    '        Dim ArrayData As String(,) = Nothing
    '    '        For Each AlignID As ObjectId In CivilDoc.GetAlignmentIds
    '    '            Dim TempAlign As Alignment = acTrans.GetObject(AlignID, OpenMode.ForRead)
    '    '            If TempAlign.StyleName Like "ИИ-Ось существущей автодороги" Then
    '    '                AlignEX = TempAlign
    '    '                PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '    '            End If
    '    '        Next
    '    '        If IsNothing(AlignEX) = True Then
    '    '            Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу существующей автодороги ")
    '    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '    '                AlignEX = ent1
    '    '                PlineAlignEx = acTrans.GetObject(AlignEX.GetPolyline, OpenMode.ForRead)
    '    '            End If
    '    '        End If
    '    '        'выбираем ось проектной трассы
    '    '        Dim ppr2 As PromptEntityResult = ed.GetEntity("Выберите трассу проектируемой автодороги ")
    '    '        Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ppr2.ObjectId, OpenMode.ForRead)
    '    '        If TypeOf ent2 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '    '            AlignProject = ent2
    '    '            PlineAlignProject = acTrans.GetObject(AlignProject.GetPolyline, OpenMode.ForRead)
    '    '        End If
    '    '        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    '        'тело программы
    '    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
    '    '        Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
    '    '        'выбираем объекты
    '    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '    '        If acSSPrompt.Status = PromptStatus.OK Then
    '    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '    '            For Each acSSObj As SelectedObject In acSSet
    '    '                If Not IsDBNull(acSSObj) Then
    '    '                    Dim colPointEX As Point3dCollection = New Point3dCollection 'точки пересечения с существующей трассой
    '    '                    Dim colPointProject As Point3dCollection = New Point3dCollection 'точки пересечения с проектной трассой
    '    '                    'начало и конец километрового положения трубопровода
    '    '                    Dim km1 As Double = -1
    '    '                    Dim kmoff1 As Double = -1
    '    '                    Dim km2 As Double = -1
    '    '                    Dim kmoff2 As Double = -1
    '    '                    'начало и конец пикетажного положения трубопровода
    '    '                    Dim pk1 As Double = -1
    '    '                    Dim pkoff1 As Double = -1
    '    '                    Dim pk2 As Double = -1
    '    '                    Dim pkoff2 As Double = -1
    '    '                    Dim Name As String = "" 'наименование сети в общем виде например Водопровод
    '    '                    Dim Purpose As String = "" 'назначение сети например Речной воды или высокого давления
    '    '                    Dim Mat As String = "" 'материал трубы
    '    '                    Dim Diametr As Double = 0 'диаметр трубы
    '    '                    Dim Height As Double = 0 'средняя глубина заложения (высота над поверхностью земли) до верхней образующей
    '    '                    Dim CountPipes As Integer = 1 'число труб
    '    '                    Dim State As String = "" 'состояние (действуюжщий, не действующай)
    '    '                    Dim Owner As String = "" 'владелец
    '    '                    Dim Adress As String = "" 'адрес
    '    '                    Dim TypePipe As String = "Подземный" 'тип сети (подземный, наземный)
    '    '                    Dim Note As String = "" 'примечание
    '    '                    Dim StyleType As String = ""
    '    '                    Dim TopLevel As Double = 0
    '    '                    'чистим массивы
    '    '                    Erase AngleEX 'массив с углами пересечек по сущестыующей трассе
    '    '                    Erase AngleProject 'массив с углами пересечек по проектной трассе
    '    '                    Erase ArrayData 'масив с расширенными данными
    '    '                    countAngleEx = 0
    '    '                    countAngleProject = 0
    '    '                    boolRez = False
    '    '                    Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
    '    '                    Dim boolReadPS As Boolean = FuncReadAllDataSetProperties(Ent, ArrayData)
    '    '                    Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
    '    '                    If TypeOf Ent Is Autodesk.Civil.DatabaseServices.FeatureLine Then
    '    '                        Dim FLine As Autodesk.Civil.DatabaseServices.FeatureLine = Ent
    '    '                        Name = FLine.StyleName
    '    '                        'превращаем характерныю линию в полилинию
    '    '                        Dim colPt As Point3dCollection = New Point3dCollection
    '    '                        colPt = FLine.GetPoints(FeatureLinePointType.PIPoint)
    '    '                        If colPt.Count > 1 Then
    '    '                            acPoly = New Autodesk.AutoCAD.DatabaseServices.Polyline()
    '    '                            For i As Integer = 0 To colPt.Count - 1
    '    '                                acPoly.AddVertexAt(i, colPt.Item(i).Convert2d(New Plane), 0, 0, 0)
    '    '                            Next
    '    '                        End If
    '    '                    ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline3d Then
    '    '                        Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline3d = Ent
    '    '                        Dim NameLayer As String = FLine.Layer
    '    '                        If NameLayer Like "L1051" Or NameLayer Like "L1050" Or NameLayer = "L1049" Or NameLayer = "L1048" Or NameLayer = "L1017" Then
    '    '                            If NameLayer Like "L1051" Then
    '    '                                Name = "Трубопровод специального назначения"
    '    '                            ElseIf NameLayer Like "L1050" Then
    '    '                                Name = "Газопровод"
    '    '                            ElseIf NameLayer Like "L1049" Then
    '    '                                Name = "Теплосеть"
    '    '                            ElseIf NameLayer Like "L1048" Then
    '    '                                Name = "Водопровод"
    '    '                            ElseIf NameLayer Like "L1017" Then
    '    '                                Name = "Канализация"
    '    '                            Else
    '    '                                Name = NameLayer
    '    '                            End If
    '    '                        End If
    '    '                        Dim Vert As Integer = 0
    '    '                        For Each vid As ObjectId In FLine
    '    '                            Dim v3d As Autodesk.AutoCAD.DatabaseServices.PolylineVertex3d = acTrans.GetObject(vid, OpenMode.ForRead)
    '    '                            acPoly.AddVertexAt(Vert, v3d.Position.Convert2d(New Plane), 0, 0, 0)
    '    '                        Next
    '    '                    ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
    '    '                        Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline2d = Ent
    '    '                        Dim NameTypeLine As String = FLine.Linetype
    '    '                        Dim posName As Integer = NameTypeLine.IndexOf("A-")
    '    '                        If posName = 0 Then
    '    '                            Name = Mid(NameTypeLine, 3)
    '    '                        End If
    '    '                        Dim Vert As Integer = 0
    '    '                        For Each vid As ObjectId In FLine
    '    '                            Dim v2d As Autodesk.AutoCAD.DatabaseServices.Vertex2d = acTrans.GetObject(vid, OpenMode.ForRead)
    '    '                            acPoly.AddVertexAt(Vert, v2d.Position.Convert2d(New Plane), 0, 0, 0)
    '    '                        Next
    '    '                    ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
    '    '                        Dim FLine As Autodesk.AutoCAD.DatabaseServices.Polyline = Ent
    '    '                        Dim NameTypeLine As String = FLine.Linetype
    '    '                        Dim posName As Integer = NameTypeLine.IndexOf("A-")
    '    '                        If posName = 0 Then
    '    '                            Name = Mid(NameTypeLine, 3)
    '    '                        End If
    '    '                        acPoly = FLine
    '    '                    ElseIf TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.line Then
    '    '                        Dim FLine As Autodesk.AutoCAD.DatabaseServices.Line = Ent
    '    '                        Dim NameTypeLine As String = FLine.Linetype
    '    '                        Dim posName As Integer = NameTypeLine.IndexOf("A-")
    '    '                        If posName = 0 Then
    '    '                            Name = Mid(NameTypeLine, 3)
    '    '                        End If
    '    '                        acPoly.AddVertexAt(0, FLine.StartPoint.Convert2d(New Plane), 0, 0, 0)
    '    '                        acPoly.AddVertexAt(1, FLine.EndPoint.Convert2d(New Plane), 0, 0, 0)
    '    '                    End If
    '    '                    '======================================================================================================
    '    '                    If IsNothing(acPoly) = False Then
    '    '                        '1. Ищем пересечение с осью трассы
    '    '                        If IsNothing(AlignEX) = False Then
    '    '                            AlignEX.IntersectWith(acPoly, Intersect.OnBothOperands, colPointEX, IntPtr.Zero, IntPtr.Zero)
    '    '                            'ещем касательную к полилинии
    '    '                            ReDim Preserve AngleEX(4, countAngleEx)
    '    '                            If colPointEX.Count > 0 Then
    '    '                                For Each pt As Point3d In colPointEX
    '    '                                    Dim dist As Double = PlineAlignEx.GetDistAtPoint(pt)
    '    '                                    AngleEX(0, countAngleEx) = dist
    '    '                                    countAngleEx += 1
    '    '                                Next
    '    '                            Else
    '    '                                Try
    '    '                                    AlignEX.StationOffset(acPoly.StartPoint.X, acPoly.StartPoint.Y, km1, kmoff1)
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                    'конец линии за пределами трассы
    '    '                                    'ищем пересечение с началом трассы
    '    '                                    If IsNothing(acPoly) = False Then
    '    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
    '    '                                        Try
    '    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '    '                                            AngleEX(1, countAngleEx) = km1
    '    '                                            AngleEX(2, countAngleEx) = kmoff1
    '    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '    '                                        End Try
    '    '                                    End If
    '    '                                End Try
    '    '                                Try
    '    '                                    AlignEX.StationOffset(acPoly.EndPoint.X, acPoly.EndPoint.Y, km1, kmoff1)
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                    'конец линии за пределами трассы
    '    '                                    'ищем пересечение с началом трассы
    '    '                                    If IsNothing(acPoly) = False Then
    '    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignEX.StartPoint, False)
    '    '                                        Try
    '    '                                            AlignEX.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '    '                                            AngleEX(3, countAngleEx) = km1
    '    '                                            AngleEX(4, countAngleEx) = kmoff1
    '    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '    '                                        End Try
    '    '                                    End If
    '    '                                End Try
    '    '                            End If
    '    '                        End If
    '    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    '                        'проектный пикетаж (начало участка)
    '    '                        '1. Ищем пересечение с осью трассы
    '    '                        If IsNothing(AlignProject) = False Then
    '    '                            AlignProject.IntersectWith(acPoly, Intersect.OnBothOperands, colPointProject, IntPtr.Zero, IntPtr.Zero)
    '    '                            ReDim Preserve AngleProject(4, countAngleProject)
    '    '                            If colPointProject.Count > 0 Then
    '    '                                For Each pt As Point3d In colPointProject
    '    '                                    Dim dist As Double = PlineAlignProject.GetDistAtPoint(pt)
    '    '                                    AngleProject(0, countAngleProject) = dist
    '    '                                    countAngleProject += 1
    '    '                                Next
    '    '                            Else
    '    '                                Try
    '    '                                    AlignProject.StationOffset(acPoly.StartPoint.X, acPoly.StartPoint.Y, km1, kmoff1)
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                    'конец линии за пределами трассы
    '    '                                    'ищем пересечение с началом трассы
    '    '                                    If IsNothing(acPoly) = False Then
    '    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
    '    '                                        Try
    '    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '    '                                            AngleProject(1, countAngleProject) = km1
    '    '                                            AngleProject(2, countAngleProject) = kmoff1
    '    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '    '                                        End Try
    '    '                                    End If
    '    '                                End Try
    '    '                                Try
    '    '                                    AlignEX.StationOffset(acPoly.EndPoint.X, acPoly.EndPoint.Y, km1, kmoff1)
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                    'конец линии за пределами трассы
    '    '                                    'ищем пересечение с началом трассы
    '    '                                    If IsNothing(acPoly) = False Then
    '    '                                        Dim t1 As Point3d = acPoly.GetClosestPointTo(AlignProject.StartPoint, False)
    '    '                                        Try
    '    '                                            AlignProject.StationOffset(t1.X, t1.Y, 0.1, km1, kmoff1)
    '    '                                            AngleProject(3, countAngleProject) = km1
    '    '                                            AngleProject(4, countAngleProject) = kmoff1
    '    '                                        Catch ex1 As Autodesk.Civil.PointNotOnEntityException
    '    '                                        End Try
    '    '                                    End If
    '    '                                End Try
    '    '                            End If
    '    '                        End If

    '    '                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    '                        '0-ПК для сортировки,1-наименование,2-КМ пересечения,3-ПК пересечения, 4,5-КМ начало, 6,7-КМ конец, 8,9-ПК начало, 10,11-ПК конец, 12-Глубина закладки, 13,14-Угол пересечения, 15-материал, 16-Диаметр, 17-Владелец, 18-Адрес, 19-Примечание
    '    '                        '===========================================================================
    '    '                        'записываем OD Характеристики
    '    '                        If IsArray(ArrayData) = True Then
    '    '                            For j As Integer = 0 To ArrayData.GetUpperBound(1)
    '    '                                Dim NameF As String = ArrayData(1, j)
    '    '                                Dim val As String = ArrayData(3, j)
    '    '                                If NameF Like "MATERIAL" Then
    '    '                                    Mat = val
    '    '                                ElseIf NameF Like "TYPE" Then
    '    '                                    TypePipe = val
    '    '                                ElseIf NameF Like "HEIGHT" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            Height = CDbl(val)
    '    '                                        Else
    '    '                                            Height = 0
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "STATE" Then
    '    '                                    State = val
    '    '                                ElseIf NameF Like "PURPOSE" Then
    '    '                                    Purpose = val
    '    '                                ElseIf NameF Like "DIAMETER" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            Diametr = CDbl(val)
    '    '                                        Else
    '    '                                            Diametr = 0
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "COUNT_PIPES" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            CountPipes = CInt(val)
    '    '                                        Else
    '    '                                            CountPipes = 1
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "TOP_LEVEL" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            TopLevel = CInt(val)
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "OWNER" Then
    '    '                                    Owner = val
    '    '                                ElseIf NameF Like "ADRESS" Then
    '    '                                    Adress = val
    '    '                                ElseIf NameF Like "NOTE" Then
    '    '                                    Note = val
    '    '                                ElseIf NameF Like "VOLTAGE" Then 'напряжение линии
    '    '                                    Power = val
    '    '                                ElseIf NameF Like "NAME1" Then 'наименование линии
    '    '                                    Name1 = val
    '    '                                ElseIf NameF Like "NAME2" Then 'наименование линии (цепь №2)
    '    '                                    Name2 = val
    '    '                                ElseIf NameF Like "NUMB_WIRES" Then 'число проводов (кабелей)
    '    '                                    CountWires = val
    '    '                                ElseIf NameF Like "MARK_WIRES" Then 'марка провода
    '    '                                    Mat = val
    '    '                                ElseIf NameF Like "SLAK_WIRES" Then 'провис провода
    '    '                                    If IsNumeric(val) = True Then
    '    '                                        If CDbl(val) > 0 Then
    '    '                                            SlakWires = CDbl(val)
    '    '                                        End If
    '    '                                    End If
    '    '                                ElseIf NameF Like "DOWN_SLAK_LEVEL" Then 'провис провода
    '    '                                    If IsNumeric(val) = True Then
    '    '                                        If CDbl(val) > 0 Then
    '    '                                            DownLewel = CDbl(val)
    '    '                                        End If
    '    '                                    End If
    '    '                                ElseIf NameF Like "TEMPERATURE" Then 'температура
    '    '                                    If IsNumeric(val) = True Then
    '    '                                        If CDbl(val) > 0 Then
    '    '                                            Temp = CDbl(val)
    '    '                                        End If
    '    '                                    End If
    '    '                                ElseIf NameF Like "DATA_SURV" Then 'дата съемки
    '    '                                    DataSurv = val
    '    '                                ElseIf NameF Like "MARK_WIRES" Then
    '    '                                    Mat = val
    '    '                                ElseIf NameF Like "DEPTH" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            Height = CDbl(val)
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "STATE" Then
    '    '                                    State = val
    '    '                                ElseIf NameF Like "PURPOSE" Then
    '    '                                    Purpose = val
    '    '                                ElseIf NameF Like "VOLTAGE" Then
    '    '                                    Power = val
    '    '                                ElseIf NameF Like "NUMB_WIRES" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            CountPipes = CInt(val)
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "TOP_LEVEL" Then
    '    '                                    Try
    '    '                                        If IsNumeric(val) = True Then
    '    '                                            TopLevel = CDbl(val)
    '    '                                        End If
    '    '                                    Catch ex As System.InvalidCastException
    '    '                                    End Try
    '    '                                ElseIf NameF Like "OWNER" Then
    '    '                                    Owner = val
    '    '                                ElseIf NameF Like "ADRESS" Then
    '    '                                    Adress = val
    '    '                                ElseIf NameF Like "NOTE" Then
    '    '                                    Note = val
    '    '                                End If
    '    '                            Next j
    '    '                        End If
    '    '                        'начинаем разбор найденных свойств
    '    '                        ReDim Preserve ArrayDataWord(20, CountArrayDataWord)
    '    '                        '1. записываем наименование объекта
    '    '                        Dim strName As String = Name & " " & TypePipe & " " & Purpose
    '    '                        If CountPipes > 1 Then
    '    '                            strName = strName & " " & CountPipes & " тр."
    '    '                        End If
    '    '                        ArrayDataWord(1, CountArrayDataWord) = strName
    '    '                        '2. записываем пересечение КМ
    '    '                        If colPointEX.Count > 0 Then
    '    '                            Dim str As String = ""
    '    '                            Dim km As Double = 99999999
    '    '                            For Each pt As Point3d In colPointEX
    '    '                                Dim kmper As Double = -1
    '    '                                Dim offper As Double = -1
    '    '                                Try
    '    '                                    AlignEX.StationOffset(pt.X, pt.Y, kmper, offper)
    '    '                                    ArrayDataWord(2, CountArrayDataWord) = ArrayDataWord(2, CountArrayDataWord) & " " & FuncFormatKM(kmper)
    '    '                                    If kmper < km Then
    '    '                                        km = kmper
    '    '                                    End If
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                End Try
    '    '                            Next
    '    '                            ArrayDataWord(0, CountArrayDataWord) = km
    '    '                        Else
    '    '                            ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '    '                        End If
    '    '                        '3. записывваем пересечение с проектной трассой
    '    '                        If colPointProject.Count > 0 Then
    '    '                            Dim str As String = ""
    '    '                            Dim pk As Double = 99999999
    '    '                            For Each pt As Point3d In colPointProject
    '    '                                Dim pkper As Double = -1
    '    '                                Dim offper As Double = -1
    '    '                                Try
    '    '                                    AlignProject.StationOffset(pt.X, pt.Y, pk, offper)
    '    '                                    ArrayDataWord(3, CountArrayDataWord) = ArrayDataWord(3, CountArrayDataWord) & " " & FuncFormatPK2(pk)
    '    '                                    If pkper < pk Then
    '    '                                        pk = pkper
    '    '                                    End If
    '    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '    '                                End Try
    '    '                                '===================================================================================
    '    '                                'удаляем предыдущие созданные точки
    '    '                                Dim NumbPoint As Integer() = Nothing
    '    '                                Dim boolIns As Boolean = FuncFindCogoPoint(pt.Convert2d(New Plane), NumbPoint, 0.02, "Трубопровод подземный")
    '    '                                If IsArray(NumbPoint) = True Then
    '    '                                    For k As Integer = 0 To NumbPoint.Length - 1
    '    '                                        Dim idpoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumbPoint(k))
    '    '                                        Dim PointCogoDel As CogoPoint = acTrans.GetObject(idpoint, OpenMode.ForWrite)
    '    '                                        If IsDBNull(PointCogoDel) = False Then
    '    '                                            PointCogoDel.Erase()
    '    '                                        End If
    '    '                                    Next k
    '    '                                End If
    '    '                                'создаем COGO - точку
    '    '                                'генерируем номер для cogo точки
    '    '                                Dim numbCogo As Long = 0
    '    '                                For i As Integer = 0 To 100
    '    '                                    Dim rng As Random = New Random
    '    '                                    numbCogo = rng.Next(1000500, 1000999)
    '    '                                    If CivilDoc.CogoPoints.Contains(numbCogo) = False Then
    '    '                                        Exit For
    '    '                                    End If
    '    '                                Next
    '    '                                'формируем имя для продольного профиля
    '    '                                Dim NameDesk As String = ""
    '    '                                If Name.Trim.Length > 0 Then
    '    '                                    NameDesk = Name
    '    '                                End If
    '    '                                If Mat.Trim.Length > 0 Then
    '    '                                    Dim TempMat As String = ""
    '    '                                    Dim pos As Integer = Mat.IndexOf("Сталь")
    '    '                                    If pos > -1 Then
    '    '                                        TempMat = "ст."
    '    '                                    End If
    '    '                                    pos = Mat.IndexOf("Железоб")
    '    '                                    If pos > -1 Then
    '    '                                        TempMat = "ж.б."
    '    '                                    End If
    '    '                                    pos = Mat.IndexOf("Чугун")
    '    '                                    If pos > -1 Then
    '    '                                        TempMat = "чуг."
    '    '                                    End If
    '    '                                    If TempMat Like "" Then
    '    '                                        TempMat = Mat
    '    '                                    End If
    '    '                                    NameDesk = NameDesk & " " & TempMat
    '    '                                End If
    '    '                                If Diametr > 0 Then
    '    '                                    NameDesk = NameDesk & " " & Diametr
    '    '                                End If
    '    '                                Dim idCogo As ObjectId = FuncCreateCOGOPoint(pt, NameDesk, numbCogo, TopLevel,, "Трубопровод подземный")
    '    '                                Dim PointCogo1 As CogoPoint = acTrans.GetObject(idCogo, OpenMode.ForWrite)
    '    '                                'присваиваем набор данных
    '    '                                Dim ArrayODRec As String(,) = Nothing
    '    '                                ReDim Preserve ArrayODRec(2, 1)
    '    '                                ArrayODRec(0, 0) = "PROFILE_PROJECT"
    '    '                                ArrayODRec(1, 0) = "LABEL"
    '    '                                ArrayODRec(2, 0) = "Подземный трубопровод на виде профиля"
    '    '                                ArrayODRec(0, 1) = "PROFILE_PROJECT"
    '    '                                ArrayODRec(1, 1) = "DESK"
    '    '                                ArrayODRec(2, 1) = NameDesk
    '    '                                If IsNothing(CreateTBLPSID) = False Then
    '    '                                    'присваиваем таблицу объекту
    '    '                                    Dim boolAddTLB As Boolean = Stroyproekt.FuncAddDataSetProperties(PointCogo1, CreateTBLPSID)
    '    '                                    'записываем данные в таблицу
    '    '                                    Dim boolWriteData As Boolean = Stroyproekt.FuncWriteAllDataSetProperties(PointCogo1, ArrayODRec)
    '    '                                End If
    '    '                            Next

    '    '                            If IsNothing(AlignEX) = True Then
    '    '                                ArrayDataWord(0, CountArrayDataWord) = pk
    '    '                            End If
    '    '                        Else
    '    '                            If IsNothing(AlignEX) = True Then
    '    '                                ArrayDataWord(0, CountArrayDataWord) = "999999999"
    '    '                            End If
    '    '                        End If
    '    '                        '3. записываем км c,kb;tybz
    '    '                        If km1 <> -1 And km2 <> -1 Then
    '    '                            If km1 > km2 Then
    '    '                                ArrayDataWord(4, CountArrayDataWord) = km2
    '    '                                ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '    '                                ArrayDataWord(6, CountArrayDataWord) = km1
    '    '                                ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '    '                            Else
    '    '                                ArrayDataWord(4, CountArrayDataWord) = km1
    '    '                                ArrayDataWord(5, CountArrayDataWord) = Math.Round(kmoff1, 2)
    '    '                                ArrayDataWord(6, CountArrayDataWord) = km2
    '    '                                ArrayDataWord(7, CountArrayDataWord) = Math.Round(kmoff2, 2)
    '    '                            End If
    '    '                        End If
    '    '                        If pk1 <> -1 And pk2 <> -1 Then
    '    '                            If pk1 > pk2 Then
    '    '                                ArrayDataWord(8, CountArrayDataWord) = pk2
    '    '                                ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '    '                                ArrayDataWord(10, CountArrayDataWord) = pk1
    '    '                                ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '    '                            Else
    '    '                                ArrayDataWord(8, CountArrayDataWord) = pk1
    '    '                                ArrayDataWord(9, CountArrayDataWord) = Math.Round(pkoff1, 2)
    '    '                                ArrayDataWord(10, CountArrayDataWord) = pk2
    '    '                                ArrayDataWord(11, CountArrayDataWord) = Math.Round(pkoff2, 2)
    '    '                            End If
    '    '                        End If
    '    '                        'глубина закладки
    '    '                        ArrayDataWord(12, CountArrayDataWord) = String.Format("{0:0.0}", Math.Round(Height, 1))
    '    '                        'вычисляем угол
    '    '                        If IsArray(AngleEX) = True Then
    '    '                            If AngleEX.GetUpperBound(1) > 0 Then
    '    '                                Dim boolSort As Boolean = FuncSortDblArray(AngleEX)
    '    '                            End If
    '    '                            For i As Integer = 0 To AngleEX.GetUpperBound(1)
    '    '                                Dim angl As Double = Math.Round((AngleEX(1, i) * 180) / Math.PI, 0)
    '    '                                If angl > 180 Then angl = 360 - angl
    '    '                                ArrayDataWord(13, CountArrayDataWord) = ArrayDataWord(13, CountArrayDataWord) & vbLf & CStr(angl)
    '    '                            Next i
    '    '                        End If
    '    '                        If IsArray(AngleProject) = True Then
    '    '                            If AngleProject.GetUpperBound(1) > 0 Then
    '    '                                Dim boolSort As Boolean = FuncSortDblArray(AngleProject)
    '    '                            End If
    '    '                            For i As Integer = 0 To AngleProject.GetUpperBound(1)
    '    '                                Dim angl As Double = Math.Round((AngleProject(1, i) * 180) / Math.PI, 0)
    '    '                                If angl > 180 Then angl = 360 - angl
    '    '                                ArrayDataWord(14, CountArrayDataWord) = ArrayDataWord(14, CountArrayDataWord) & vbLf & CStr(angl)
    '    '                            Next i
    '    '                        End If
    '    '                        ArrayDataWord(15, CountArrayDataWord) = Mat 'материал трубы
    '    '                        ArrayDataWord(16, CountArrayDataWord) = Diametr 'диаметр трубы
    '    '                        ArrayDataWord(17, CountArrayDataWord) = Owner 'владелец
    '    '                        ArrayDataWord(18, CountArrayDataWord) = Adress 'адрес
    '    '                        ArrayDataWord(19, CountArrayDataWord) = Note 'примечание
    '    '                        CountArrayDataWord += 1
    '    '                    End If
    '    '                End If
    '    '            Next
    '    '        End If
    '    '        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    '        'записываем все в массив
    '    '        If IsArray(ArrayDataWord) = True Then
    '    '            'сортируем массив
    '    '            Dim boolSort As Boolean = FuncSortArray3(ArrayDataWord, 0)
    '    '            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    '            'выбираем документ Word
    '    '            Try
    '    '                WApp = CreateObject("Word.Application")
    '    '                WDoc = WApp.Documents.Open(PutchTbl & "Ведомость пересечения и сближения подземных коммуникаций (трубопроводные сети).docx")
    '    '                Dim rng As Random = New Random
    '    '                Dim Numbrng As Integer = rng.Next
    '    '                WDoc.SaveAs2(TempDir & "Ведомость подземных трубопроводных сетей_" & Numbrng & ".docx")
    '    '            Catch
    '    '                MsgBox("Не удалось найти выбранную таблицу!!!")
    '    '                locker.Dispose()
    '    '                WDoc = Nothing
    '    '                WApp = Nothing
    '    '                Exit Sub
    '    '            End Try
    '    '            For i As Integer = 0 To ArrayDataWord.GetUpperBound(1)
    '    '                For Each tabl As Word.Table In WDoc.Tables
    '    '                    Dim countTableword As Integer = tabl.Rows.Count
    '    '                    If countTableword > ArrayDataWord.GetUpperBound(1) + 4 Then
    '    '                        Exit For
    '    '                    End If
    '    '                    tabl.Rows.Add()
    '    '                    '===================================================================================
    '    '                    'название объекта
    '    '                    If IsNothing(ArrayDataWord(1, i)) = False Then
    '    '                        tabl.Cell(countTableword, 1).Range.Text = ArrayDataWord(1, i)
    '    '                    End If
    '    '                    '===================================================================================
    '    '                    'километр пересечения
    '    '                    If IsNothing(ArrayDataWord(2, i)) = False Then
    '    '                        tabl.Cell(countTableword, 2).Range.Text = ArrayDataWord(2, i)
    '    '                    End If
    '    '                    '===================================================================================
    '    '                    'записываем смещение относительно существующей трассы
    '    '                    'String.Format("{0:0.00}", Math.Abs(Math.Round(kmoff2, 2)))
    '    '                    If IsNothing(ArrayDataWord(4, i)) = False And IsNothing(ArrayDataWord(5, i)) = False Then
    '    '                        Dim kmStart As String = FuncFormatKM(Val(ArrayDataWord(4, i)))
    '    '                        Dim kmoff1 As Double = Val(ArrayDataWord(5, i))
    '    '                        If kmoff1 < 0 Then
    '    '                            tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'влево
    '    '                        ElseIf kmoff1 > 0 Then
    '    '                            tabl.Cell(countTableword, 3).Range.Text = kmStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff1)) 'вправо
    '    '                        Else
    '    '                            tabl.Cell(countTableword, 3).Range.Text = kmStart
    '    '                        End If
    '    '                    End If
    '    '                    If IsNothing(ArrayDataWord(6, i)) = False And IsNothing(ArrayDataWord(7, i)) = False Then
    '    '                        Dim kmEnd As String = FuncFormatKM(Val(ArrayDataWord(6, i)))
    '    '                        Dim kmoff2 As Double = Val(ArrayDataWord(7, i))
    '    '                        If kmoff2 < 0 Then
    '    '                            tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'влево
    '    '                        ElseIf kmoff2 > 0 Then
    '    '                            tabl.Cell(countTableword, 4).Range.Text = kmEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(kmoff2)) 'вправо
    '    '                        Else
    '    '                            tabl.Cell(countTableword, 4).Range.Text = kmEnd
    '    '                        End If
    '    '                    End If
    '    '                    '=================================================================================================
    '    '                    'записываем угол пересечения
    '    '                    If IsNothing(ArrayDataWord(13, i)) = False Then
    '    '                        tabl.Cell(countTableword, 5).Range.Text = ArrayDataWord(13, i)
    '    '                    End If
    '    '                    '=================================================================================================
    '    '                    'записываем пикет пересечения
    '    '                    If IsNothing(ArrayDataWord(3, i)) = False Then
    '    '                        tabl.Cell(countTableword, 6).Range.Text = ArrayDataWord(3, i)
    '    '                    End If
    '    '                    '=================================================================================================
    '    '                    'записываем смещение относительно существующей трассы
    '    '                    If IsNothing(ArrayDataWord(8, i)) = False And IsNothing(ArrayDataWord(9, i)) = False Then
    '    '                        Dim pkStart As String = FuncFormatPK2(Val(ArrayDataWord(8, i)))
    '    '                        Dim pkoff1 As Double = Val(ArrayDataWord(9, i))
    '    '                        If pkoff1 < 0 Then
    '    '                            tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'влево
    '    '                        ElseIf pkoff1 > 0 Then
    '    '                            tabl.Cell(countTableword, 7).Range.Text = pkStart & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff1)) 'вправо
    '    '                        Else
    '    '                            tabl.Cell(countTableword, 7).Range.Text = pkStart
    '    '                        End If
    '    '                    End If
    '    '                    If IsNothing(ArrayDataWord(10, i)) = False And IsNothing(ArrayDataWord(11, i)) = False Then
    '    '                        Dim pkEnd As String = FuncFormatPK2(Val(ArrayDataWord(10, i)))
    '    '                        Dim pkoff2 As Double = Val(ArrayDataWord(11, i))
    '    '                        If pkoff2 < 0 Then
    '    '                            tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Лево" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'влево
    '    '                        ElseIf pkoff2 > 0 Then
    '    '                            tabl.Cell(countTableword, 8).Range.Text = pkEnd & vbLf & "Право" & String.Format("{0:0.00}", Math.Abs(pkoff2)) 'вправо
    '    '                        Else
    '    '                            tabl.Cell(countTableword, 8).Range.Text = pkEnd
    '    '                        End If
    '    '                    End If

    '    '                    '=================================================================================================
    '    '                    'записываем угол пересечения
    '    '                    If IsNothing(ArrayDataWord(14, i)) = False Then
    '    '                        tabl.Cell(countTableword, 9).Range.Text = ArrayDataWord(14, i)
    '    '                    End If
    '    '                    '===================================================================================================
    '    '                    If IsNothing(ArrayDataWord(12, i)) = False Then
    '    '                        tabl.Cell(countTableword, 10).Range.Text = ArrayDataWord(12, i) 'глубина закладки
    '    '                    End If
    '    '                    '==================================================================================================
    '    '                    'материал и диаметр
    '    '                    Dim strMat As String = ""
    '    '                    If IsNothing(ArrayDataWord(15, i)) = False Then
    '    '                        strMat = ArrayDataWord(15, i)
    '    '                    End If
    '    '                    Dim strD As String = ""
    '    '                    If IsNothing(ArrayDataWord(16, i)) = False Then
    '    '                        strD = ArrayDataWord(16, i)
    '    '                    End If
    '    '                    If strMat.Length > 0 And strD.Length > 0 Then
    '    '                        tabl.Cell(countTableword, 11).Range.Text = strMat & vbLf & strD
    '    '                    ElseIf strMat.Length > 0 And strD.Length = 0 Then
    '    '                        tabl.Cell(countTableword, 11).Range.Text = strMat
    '    '                    ElseIf strMat.Length = 0 And strD.Length > 0 Then
    '    '                        tabl.Cell(countTableword, 11).Range.Text = strD
    '    '                    Else
    '    '                        tabl.Cell(countTableword, 11).Range.Text = "-"
    '    '                    End If
    '    '                    '=================================================================================================
    '    '                    If IsNothing(ArrayDataWord(17, i)) = False Then
    '    '                        tabl.Cell(countTableword, 12).Range.Text = ArrayDataWord(17, i) '& vbLf & ArrayDataWord(18, i) 'владелец и адрес
    '    '                    End If
    '    '                    If IsNothing(ArrayDataWord(19, i)) = False Then
    '    '                        tabl.Cell(countTableword, 13).Range.Text = ArrayDataWord(19, i) 'описание
    '    '                    End If
    '    '                    Exit For
    '    '                Next
    '    '            Next i
    '    '            WApp.Visible = True
    '    '        End If
    '    '        acTrans.Commit()
    '    '    End Using
    '    'End Sub
    '==================================================================================================
    'Ведомость пересечения (сближения) подземных коммуникаций


End Class
