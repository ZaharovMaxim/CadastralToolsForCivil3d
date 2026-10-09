Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide

Public Class CatalogExportCoordinates
    '=========================================================================================================
    'таблица с координатами участков постоянного и временного отводов в MS Excel, Word, Autocad
    <CommandMethod("ExportCatalogCoordinatesParcels")>
    Public Sub ExportCatalogCoordinatesParcels()
        '========================================================================================
        'загружаем пути к сборкам
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '=========================================================================================
        '=============================================================================================================
        'проверяем и загружаем все стили из шаблона
        '=============================================================================================================
        Dim ArrayFilesDirTemlates As String() = Nothing
        Dim CountPutchTemplateDir As Integer = 0
        Dim PutchTemplateDir As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplateDir = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In Directory.GetFiles(PutchTemplateDir)
                Try
                    Dim File1 As String = Path.GetFileName(fileDwg)
                    Dim ext As String = Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayFilesDirTemlates(CountPutchTemplateDir)
                        ArrayFilesDirTemlates(CountPutchTemplateDir) = File1
                        CountPutchTemplateDir += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim WApp As Word.Application = Nothing
        Dim WDoc As Word.Document = Nothing
        Dim FormReport As FormReportCatalogCoordinates = New FormReportCatalogCoordinates
        FormReport.WApp = WApp
        FormReport.WDoc = WDoc
        If IsArray(ArrayFilesDirTemlates) = True Then
            FormReport.ComboBox2.DataSource = ArrayFilesDirTemlates
        End If
        FormReport.ShowDialog()
        If FormReport.boolRp = False Then Exit Sub
        'округление значений
        Dim RoundZn As Integer = FormReport.NumericUpDown1.Value
        Dim AreaRoundZn As Integer = FormReport.NumericUpDown2.Value
        Dim PutchTemplateFile As String = PutchTemplateDir & FormReport.ComboBox2.Text.Trim
        Dim tablePS As String = FormReport.ComboBox3.Text
        If IsNothing(tablePS) = True Then tablePS = ""
        Dim arrayTableXML As String(,) = Nothing
        ReDim arrayTableXML(1, 0)
        arrayTableXML(0, 0) = tablePS
        arrayTableXML(1, 0) = ""

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех объекто чертежа которым присвоен набор данных PS на всем пространстве чертежа (handle, nameOD), кроме островков
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayOstrovPoligon As String(,) = Nothing 'массив для хранения внутренних островков
        Dim countArrayOstrovPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayPSTables As String() = Nothing
        Dim CountArrayPSTables As Integer = 0

        Dim arrayRecFields As String(,) = Nothing
        Dim countFieds As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim PointInsTbl1 As Point3d = Nothing
        Dim ArrayUstanovki As String(,) = Nothing
        Dim countArrayUstanovki As Integer = 0
        Dim NameDir As String = ""
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'лимиты карты для MapInfo
        Dim MinCoordX As Double = 100000000
        Dim MinCoordY As Double = 100000000
        Dim MaxCoordX As Double = -10000000
        Dim MaxCoordY As Double = -10000000
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        '======================================================================================================================
        Dim IDTablePS As ObjectId = Nothing
        If FormReport.RadioButton1.Checked = True Then
            If FormReport.ComboBox1.Text Like "Таблица координат" Then
                IDTablePS = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_OBJECT")
                If File.Exists(PutchTemplateFile) = True Then
                    If IsNothing(IDTablePS) = True Then
                        Dim boolCreatePs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateFile, "CADASTRAL_OBJECT")
                    ElseIf IDTablePS.IsNull = True Then
                        Dim boolCreatePs As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateFile, "CADASTRAL_OBJECT")
                    End If
                Else
                    MsgBox("Не удалось скопировать из шаблона Набор Характеристик CADASTRAL_OBJECT. Макро прервано!!!")
                    Exit Sub
                End If
            End If
        End If
        Dim docloc As DocumentLock = acDoc.LockDocument()
        Using docloc
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Try
                Using tr
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim GeoPointCoolection As ObjectIdCollection = New ObjectIdCollection
                    Dim sourceBtr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                    For Each id As ObjectId In sourceBtr
                        If id.IsNull = False Then
                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                            Dim idLr As ObjectId = ent.LayerId
                            Dim lyrTemp As LayerTableRecord = idLr.GetObject(OpenMode.ForRead)
                            If lyrTemp.IsFrozen = True Then
                                Continue For
                            End If
                            If lyrTemp.IsOff = True Then
                                Continue For
                            End If
                            'блок
                            If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                                Dim blkPt As BlockReference = ent
                                Dim nameUserBlk As String = AcadFunction.FuncReadNameBlk(blkPt)
                                If nameUserBlk Like "PROJECT_POINT" Then
                                    GeoPointCoolection.Add(id)
                                ElseIf nameUserBlk Like "TEMP_POINT" Then
                                    GeoPointCoolection.Add(id)
                                End If
                            End If
                        End If
                    Next
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '3.выбираем участки
                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        For Each acSSObj As SelectedObject In acSSet
                            If Not IsDBNull(acSSObj) Then
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                If Not IsDBNull(acEnt) Then
                                    '===================================================================================================
                                    'читаем расширенные данные
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, arrayTableXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                            If ArrayNameTableOD(i).Trim Like "CADASTRAL_POINT" Then
                                                Continue For
                                            Else
                                                ReDim Preserve ArrayAllPoligon(1, countArrayAllPoligon)
                                                ArrayAllPoligon(0, countArrayAllPoligon) = acEnt.Handle.ToString 'Handle выбранного контура
                                                ArrayAllPoligon(1, countArrayAllPoligon) = ArrayNameTableOD(i) 'тип таблицы
                                                countArrayAllPoligon += 1
                                                'определеям лимиты карты (только для экспорта в MapInfo)
                                                Dim ext As Extents3d = acEnt.GeometricExtents
                                                If MinCoordX > ext.MinPoint.X Then
                                                    MinCoordX = ext.MinPoint.X
                                                End If
                                                If MinCoordY > ext.MinPoint.Y Then
                                                    MinCoordY = ext.MinPoint.Y
                                                End If
                                                If MaxCoordX < ext.MaxPoint.X Then
                                                    MaxCoordX = ext.MaxPoint.X
                                                End If
                                                If MaxCoordY < ext.MaxPoint.Y Then
                                                    MaxCoordY = ext.MaxPoint.Y
                                                End If
                                                If IsArray(ArrayPSTables) = False Then
                                                    ReDim Preserve ArrayPSTables(CountArrayPSTables)
                                                    ArrayPSTables(CountArrayPSTables) = ArrayNameTableOD(i).Trim
                                                    CountArrayPSTables += 1
                                                Else
                                                    Dim boolFindTB As Boolean = False
                                                    For i1 As Integer = 0 To ArrayPSTables.Length - 1
                                                        If ArrayPSTables(i1) Like ArrayNameTableOD(i).Trim Then
                                                            boolFindTB = True
                                                            Exit For
                                                        End If
                                                    Next i1
                                                    If boolFindTB = False Then
                                                        ReDim Preserve ArrayPSTables(CountArrayPSTables)
                                                        ArrayPSTables(CountArrayPSTables) = ArrayNameTableOD(i).Trim
                                                        CountArrayPSTables += 1
                                                    End If
                                                End If
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                        Next
                    End If
                    'неудача - выходим из программы
                    If IsArray(ArrayAllPoligon) = False Then
                        MsgBox("Не удалось найти объекты-участки!!!")
                        Exit Sub
                    Else
                        If ArrayAllPoligon.GetUpperBound(1) > 1 Then
                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '4.исключаем из выбора контура многоконтурного участка кроме основного контура
                    For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                        Dim PlineHandle As String = ArrayAllPoligon(0, i) 'hg полилинии
                        Dim NameOD As String = ArrayAllPoligon(1, i) 'имя набора данных PS
                        Dim IdPlineSelect As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                        If IsNothing(IdPlineSelect) = False Then
                            If IdPlineSelect.IsNull = False Then
                                Dim EntObject As Autodesk.AutoCAD.DatabaseServices.Entity = IdPlineSelect.GetObject(OpenMode.ForRead)
                                If TypeOf EntObject Is Polyline Or TypeOf EntObject Is Polyline2d Or TypeOf EntObject Is Polyline3d Or TypeOf EntObject Is DBPoint Or TypeOf EntObject Is Circle Or TypeOf EntObject Is FeatureLine Or TypeOf EntObject Is BlockReference Or TypeOf EntObject Is CogoPoint Then
                                    Dim GlobalID As String = FuncPropertySet.FuncReadValueSetProperties(EntObject, "GLOBALID")
                                    Dim LocalID As String = FuncPropertySet.FuncReadValueSetProperties(EntObject, "LOCALID")
                                    Dim ConditionalParcel As String = FuncPropertySet.FuncReadValueSetProperties(EntObject, "ConditionalNumber")
                                    If FormReport.ComboBox4.Text Like "00:00:0000000:0000:ЗУ" Then
                                        Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(EntObject, "Cad_number")
                                        If cad_number.Trim.Length > 0 Then
                                            cad_number = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(cad_number)
                                            ConditionalParcel = cad_number & ":ЗУ" & ConditionalParcel
                                        End If
                                    ElseIf FormReport.ComboBox4.Text Like "00:00:0000000:0000/чзу" Then
                                        Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(EntObject, "Cad_number")
                                        If cad_number.Trim.Length > 0 Then
                                            cad_number = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(cad_number)
                                            ConditionalParcel = cad_number & "/чзу" & ConditionalParcel
                                        End If
                                    Else
                                        ConditionalParcel = FormReport.ComboBox4.Text & ConditionalParcel
                                    End If
                                    Dim numbParcel As Integer = 0
                                    If ConditionalParcel.Trim.Length > 0 Then
                                        'получаем отдельно номер участка для сортировки
                                        numbParcel = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionalParcel)
                                    End If
                                    Dim boolFind As Boolean = False
                                    If FormReport.CheckBox6.Checked = False Then 'игнорировать многоконтурность
                                        If GlobalID.Trim.Length > 2 Then
                                            'ищем глобальный ид в массиве дабы исключить дубли
                                            If IsArray(ArrayPoligon) = True Then
                                                For j As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                                                    If ArrayPoligon(3, j) Like GlobalID Then
                                                        boolFind = True
                                                        Exit For
                                                    End If
                                                Next j
                                            End If
                                        End If
                                        If boolFind = False Then
                                            'записываем результат в массив
                                            ReDim Preserve ArrayPoligon(5, countArrayPoligon)
                                            ArrayPoligon(0, countArrayPoligon) = numbParcel 'порядковый номер
                                            ArrayPoligon(1, countArrayPoligon) = PlineHandle 'handle контура
                                            ArrayPoligon(2, countArrayPoligon) = NameOD 'имя набора данных
                                            ArrayPoligon(3, countArrayPoligon) = GlobalID 'глобальный id 
                                            ArrayPoligon(4, countArrayPoligon) = LocalID 'локальный id 
                                            ArrayPoligon(5, countArrayPoligon) = ConditionalParcel 'обозначение участка
                                            countArrayPoligon += 1
                                        End If
                                    Else
                                        'записываем все участки игнорируя многоконтурность
                                        ReDim Preserve ArrayPoligon(5, countArrayPoligon)
                                        ArrayPoligon(0, countArrayPoligon) = ConditionalParcel 'порядковый номер
                                        ArrayPoligon(1, countArrayPoligon) = PlineHandle 'handle контура
                                        ArrayPoligon(2, countArrayPoligon) = NameOD 'имя набора данных
                                        ArrayPoligon(3, countArrayPoligon) = "" 'глобальный id 
                                        ArrayPoligon(4, countArrayPoligon) = LocalID 'локальный id 
                                        ArrayPoligon(5, countArrayPoligon) = ConditionalParcel 'обозначение участка
                                        countArrayPoligon += 1
                                    End If
                                End If
                            End If
                        End If
                    Next i
                    'сортируем контура
                    If ArrayPoligon.GetUpperBound(1) > 0 Then
                        Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayPoligon, 0)
                    End If
                    '========================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'тело программы
                    'последовательно достаем участки из массива
                    If IsArray(ArrayPoligon) = True Then
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'точка вставки таблицы автокад
                        Dim pointInsTLB As Point3d = New Point3d(0, 0, 0)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'запускаем Excel
                        Dim xlApp As Excel.Application = Nothing
                        Dim xlBook As Excel.Workbook = Nothing
                        Dim xlSheets As Excel.Worksheet = Nothing
                        Dim boolExVisibl As Boolean = False
                        Dim StartRows As Integer = 1
                        If FormReport.RadioButton2.Checked = True Then
                            xlApp = CreateObject("Excel.Application")
                            xlBook = xlApp.Workbooks.Add()
                        End If

                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'Word документ
                        Dim boolWDVisibl As Boolean = False
                        Dim SetPos As Integer = 0

                        If FormReport.RadioButton3.Checked = True And FormReport.ComboBox1.Text Like "Новый файл docx" Then
                            WApp = CreateObject("Word.Application")
                            WDoc = WApp.Documents.Add()
                            SetPos = WDoc.Content.End - 1
                            WApp.Visible = True
                        ElseIf FormReport.RadioButton3.Checked = True And FormReport.ComboBox1.Text Like "Выбрать файл...." Then
                            WApp = FormReport.WApp
                            WDoc = FormReport.WDoc
                            Dim rngSelect As Word.Selection = WApp.Selection
                            SetPos = rngSelect.End - 1
                        ElseIf FormReport.RadioButton3.Checked = True Then
                            WApp = New Microsoft.Office.Interop.Word.Application()
                            Try
                                WApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
                                If IsNothing(WApp) = False Then
                                    For i As Integer = 0 To WApp.Windows.Count - 1
                                        Dim idx As Object = i + 1
                                        Dim WinObj As Word.Window = WApp.Windows.Item(idx)
                                        If WinObj.Document.Name Like FormReport.ComboBox1.Text Then
                                            WDoc = WApp.Documents.Item(idx)
                                            Exit For
                                        End If
                                    Next
                                End If
                                Dim rngSelect As Word.Selection = WApp.Selection
                                SetPos = rngSelect.End - 1
                            Catch ex As COMException
                            End Try
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'mid mif
                        Dim ParcelMif1 As StreamWriter = Nothing
                        Dim ParcelMid1 As StreamWriter = Nothing
                        Dim strFolder As String = ""
                        Dim TableMif As String = ""
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'начинаем перебирать таблицы (Цикл верхнего уровня)
                        If IsArray(ArrayPoligon) = True Then
                            'цикл для перебора объектов с одним типом Набора данных
                            For j As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                                Dim NamePSTable As String = ArrayPoligon(2, j) 'имя таблицы PS
                                'читаем все поля данной таблицы
                                Dim ArrayDataPSFields As String(,) = Nothing
                                Dim booReadPSFields As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NamePSTable, ArrayDataPSFields)
                                'читаем идентификатор объекта
                                Dim hgParcel As String = ArrayPoligon(1, j) 'handle объекта
                                Dim idGlobalParcel As String = ArrayPoligon(3, j) 'глобальный номер участка
                                Dim LocalIdParcel As String = ArrayPoligon(4, j) 'локальный ID
                                Dim ConditionCounterParcel As String = ArrayPoligon(5, j) 'обозначение
                                'массивы для данных
                                Dim ArrayCounter As String(,) = Nothing 'массив с входящими контурами для последующей обработки (0-Номер контура, 1-handle, 2-LocalID)
                                Dim ArrayFindCounter As String(,) = Nothing 'временный массив для найденных контуров
                                'получаем доступ к объекту
                                Dim idhgParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgParcel)
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idhgParcel.GetObject(OpenMode.ForRead)
                                '==================================================================================================================
                                'читаем расщиренные данные с объекта
                                Dim ArrayDataOD As String(,) = Nothing
                                Dim NameTBLXML As String(,) = Nothing
                                ReDim Preserve NameTBLXML(1, 0)
                                NameTBLXML(0, 0) = NamePSTable
                                Dim boolreadOD As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, ArrayDataOD, NameTBLXML)
                                '=====================================================================
                                'погрешность определения точки
                                Dim categoryParcel As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Category")
                                Dim delta As Double = FuncCadastr.FuncReturnDeltaByCategory(categoryParcel)
                                'достаем контура (0-хангл, 1-локальный ид)
                                If idGlobalParcel.Trim.Length > 2 Then
                                    Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(idGlobalParcel, ArrayFindCounter, NamePSTable)
                                    'если участок многоконтурный тогда выбираем все контура
                                    If IsArray(ArrayFindCounter) = True Then
                                        If ArrayFindCounter.GetUpperBound(1) > 0 Then
                                            'это многоконтурный участок ищем обозначения всех контуров
                                            For k As Integer = 0 To ArrayFindCounter.GetUpperBound(1)
                                                Dim LocalTempCounter As String = ArrayFindCounter(1, k)
                                                Dim hgCounter As String = ArrayFindCounter(0, k)
                                                Dim numbCounter As Integer = 0
                                                Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                                                If IsNothing(idCounter) = False Then
                                                    If idCounter.IsNull = False Then
                                                        Dim EntCounter As Autodesk.AutoCAD.DatabaseServices.Entity = idCounter.GetObject(OpenMode.ForRead)
                                                        ConditionCounterParcel = FuncPropertySet.FuncReadValueSetProperties(EntCounter, "ConditionalNumber")
                                                        numbCounter = FuncCadastr.FuncFindNumberCounterCadastralObject(ConditionCounterParcel)
                                                    End If
                                                End If
                                                ReDim Preserve ArrayCounter(2, k)
                                                ArrayCounter(0, k) = numbCounter 'номер контура
                                                ArrayCounter(1, k) = hgCounter 'handle контура
                                                ArrayCounter(2, k) = LocalTempCounter 'handle контура
                                            Next k
                                            'сортируем контура
                                            If ArrayCounter.GetUpperBound(1) > 0 Then
                                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCounter, 0)
                                            End If
                                        Else
                                            'участок у которого есть GLOBALID но найден только один контур
                                            ReDim Preserve ArrayCounter(2, 0)
                                            ArrayCounter(0, 0) = "0"
                                            ArrayCounter(1, 0) = hgParcel 'handle контура
                                            ArrayCounter(2, 0) = LocalIdParcel
                                        End If
                                    Else
                                        'участок у которого есть GLOBALID но контуров не найдено
                                        ReDim Preserve ArrayCounter(2, 0)
                                        ArrayCounter(0, 0) = "0"
                                        ArrayCounter(1, 0) = hgParcel 'handle контура
                                        ArrayCounter(2, 0) = LocalIdParcel
                                    End If
                                Else
                                    'участок не многоконтурный
                                    ReDim Preserve ArrayCounter(2, 0)
                                    ArrayCounter(0, 0) = "0"
                                    ArrayCounter(1, 0) = hgParcel 'handle контура
                                    ArrayCounter(2, 0) = LocalIdParcel
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'начинаем обрабатывать каждый контур в отдельности (ищем для него островки)
                                '============================================================================================
                                Dim ArrayKoordinate As String(,) = Nothing 'массив с точками (координаты)
                                Dim CountArrayKoordinate As Integer = 0 'счетчик координат
                                '============================================================================================
                                'массив пуст
                                If IsArray(ArrayCounter) = False Then
                                    Continue For
                                End If
                                '================================================================================================
                                'достаем координаты узлов поворотных точек контура 
                                Dim CountRegions As Integer = 0
                                Dim CountCoordFirstVertex As Integer = 1
                                Dim TotalArea As Double = 0
                                Dim TotalLenght As Double = 0
                                Dim boolRegion As Boolean = True
                                For j1 As Integer = 0 To ArrayCounter.GetUpperBound(1)
                                    Dim SQRParcel As Double = 0 'площадь участка
                                    Dim LenghtObj As Double = 0 'длина конструкции
                                    Dim numCounter As Integer = Val(ArrayCounter(0, j1))
                                    Dim HandlePLine As String = ArrayCounter(1, j1) 'нангл контура
                                    LocalIdParcel = ArrayCounter(2, j1)
                                    'получаем доступ к полилинии для извлечения координат
                                    Dim id As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                    If IsNothing(id) = False Then
                                        If id.IsNull = False Then
                                            Dim EntPline As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                                            Dim Pt3dColliction As Point3dCollection = New Point3dCollection()
                                            If TypeOf EntPline Is Polyline Then
                                                Dim PlineObj As Autodesk.AutoCAD.DatabaseServices.Polyline = EntPline
                                                For k = 0 To PlineObj.NumberOfVertices - 1
                                                    If k = PlineObj.NumberOfVertices - 1 Then
                                                        If PlineObj.GetPoint3dAt(0) = PlineObj.GetPoint3dAt(k) Then
                                                            Exit For
                                                        End If
                                                    End If
                                                    Pt3dColliction.Add(PlineObj.GetPoint3dAt(k))
                                                Next

                                                If PlineObj.Closed = True Then
                                                    SQRParcel = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                                    boolRegion = True
                                                Else
                                                    LenghtObj = PlineObj.Length
                                                    boolRegion = False
                                                End If
                                            ElseIf TypeOf EntPline Is Polyline2d Then
                                                Dim p2d As Polyline2d = TryCast(EntPline, Polyline2d)
                                                If p2d IsNot Nothing Then
                                                    For Each vId As ObjectId In p2d
                                                        Dim v2d As Vertex2d = CType(tr.GetObject(vId, OpenMode.ForRead), Vertex2d)
                                                        Pt3dColliction.Add(New Point3d(v2d.Position.X, v2d.Position.Y, 0))
                                                    Next
                                                End If
                                                If p2d.Closed = True Then
                                                    SQRParcel = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                                    boolRegion = True
                                                Else
                                                    LenghtObj = p2d.Length
                                                    boolRegion = False
                                                End If
                                            ElseIf TypeOf EntPline Is Polyline3d Then
                                                Dim p3d As Polyline3d = TryCast(EntPline, Polyline3d)
                                                If p3d IsNot Nothing Then
                                                    For Each vId As ObjectId In p3d
                                                        Dim v3d As PolylineVertex3d = CType(tr.GetObject(vId, OpenMode.ForRead), PolylineVertex3d)
                                                        Pt3dColliction.Add(New Point3d(v3d.Position.X, v3d.Position.Y, v3d.Position.Z))
                                                    Next
                                                End If
                                                If p3d.Closed = True Then
                                                    SQRParcel = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                                    boolRegion = True
                                                Else
                                                    LenghtObj = p3d.Length
                                                    boolRegion = False
                                                End If
                                            ElseIf TypeOf EntPline Is FeatureLine Then
                                                Dim fLine As Autodesk.Civil.DatabaseServices.FeatureLine = TryCast(EntPline, Autodesk.Civil.DatabaseServices.FeatureLine)
                                                Pt3dColliction = fLine.GetPoints(FeatureLinePointType.PIPoint)
                                                If fLine.Closed = True Then
                                                    SQRParcel = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                                    boolRegion = True
                                                Else
                                                    LenghtObj = fLine.Length2D
                                                    boolRegion = False
                                                End If
                                            ElseIf TypeOf EntPline Is DBPoint Then
                                                Dim pt As DBPoint = EntPline
                                                Pt3dColliction.Add(New Point3d(pt.Position.X, pt.Position.Y, pt.Position.Z))
                                                boolRegion = False
                                            ElseIf TypeOf EntPline Is BlockReference Then
                                                Dim pt As BlockReference = EntPline
                                                Pt3dColliction.Add(New Point3d(pt.Position.X, pt.Position.Y, pt.Position.Z))
                                                boolRegion = False
                                            ElseIf TypeOf EntPline Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                Dim pt As Autodesk.Civil.DatabaseServices.CogoPoint = EntPline
                                                Pt3dColliction.Add(New Point3d(pt.Easting, pt.Northing, pt.Elevation))
                                                boolRegion = False
                                            ElseIf TypeOf EntPline Is Circle Then
                                                Dim pt As Circle = EntPline
                                                Pt3dColliction.Add(New Point3d(pt.Center.X, pt.Center.Y, pt.Center.Z))
                                                boolRegion = False
                                            ElseIf TypeOf EntPline Is DBText Then
                                                Dim pt As DBText = EntPline
                                                Pt3dColliction.Add(New Point3d(pt.Position.X, pt.Position.Y, pt.Position.Z))
                                                boolRegion = False
                                            Else
                                                Continue For
                                            End If
                                            '============================================================================================
                                            '============================================================================================
                                            'ищем островки
                                            Dim ArrayOstrov As String(,) = Nothing 'массив с островками
                                            If boolRegion = True Then
                                                Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalIdParcel, ArrayOstrov)
                                                If IsArray(ArrayOstrov) = True Then
                                                    If ArrayOstrov.GetUpperBound(1) > 0 Then
                                                        'сортируем массив
                                                        Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayOstrov, 0)
                                                    End If
                                                End If
                                            End If
                                            ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                            ArrayKoordinate(0, CountArrayKoordinate) = numCounter 'записываем краткий номер контура
                                            ArrayKoordinate(6, CountArrayKoordinate) = delta
                                            CountArrayKoordinate += 1
                                            'фиксируем первую координату контура
                                            CountCoordFirstVertex = CountArrayKoordinate
                                            'последовательно перебираем все вершины контура для записи в массив координат поворотных точек
                                            Dim countVert As Integer = 0 'счетчик вершин
                                            For i1 As Integer = 0 To Pt3dColliction.Count - 1
                                                Dim VertCoord As Point3d = Pt3dColliction.Item(i1)
                                                'записываем характеристики в массив
                                                ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                                ArrayKoordinate(0, CountArrayKoordinate) = "" 'номер вершины (пока пустое поле)
                                                ArrayKoordinate(1, CountArrayKoordinate) = VertCoord.X 'кооордината Х
                                                ArrayKoordinate(2, CountArrayKoordinate) = VertCoord.Y 'координата Y
                                                ArrayKoordinate(3, CountArrayKoordinate) = VertCoord.Z 'координата Z
                                                ArrayKoordinate(4, CountArrayKoordinate) = "0.1" 'дельта
                                                ArrayKoordinate(5, CountArrayKoordinate) = "626003000000" 'код закрепления точки
                                                ArrayKoordinate(6, CountArrayKoordinate) = "692006000000" 'метод определения координат
                                                countVert += 1
                                                'находим номер точки
                                                If GeoPointCoolection.Count > 0 Then
                                                    Dim FindPointID As ObjectId = FuncCadastr.FuncFindGeoPoint(GeoPointCoolection, VertCoord.Convert2d(New Plane), FormReport.NumericUpDown3.Value)
                                                    If IsNothing(FindPointID) = False Then
                                                        If FindPointID.IsNull = False Then
                                                            Dim entBlk As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(FindPointID, OpenMode.ForRead)
                                                            If TypeOf entBlk Is BlockReference Then
                                                                Dim blk As BlockReference = entBlk
                                                                Dim ArrayAttrib As String() = Nothing
                                                                Dim FindPointGeo As Boolean = FuncCadastr.FuncReturnBlkArtibuteGeoPoint(blk, ArrayAttrib)
                                                                If IsArray(ArrayAttrib) = True Then
                                                                    ArrayKoordinate(0, CountArrayKoordinate) = ArrayAttrib(0) 'номер точки 1, 2 ...
                                                                    If ArrayAttrib.GetUpperBound(0) > 0 Then
                                                                        ArrayKoordinate(4, CountArrayKoordinate) = ArrayAttrib(1) 'дельта
                                                                    End If
                                                                    If ArrayAttrib.GetUpperBound(0) > 1 Then
                                                                        Dim KodZakr As String = FuncCadastr.FuncZakrepPointKod(ArrayAttrib(2).Trim)
                                                                        ArrayKoordinate(5, CountArrayKoordinate) = KodZakr 'код закрепления точки
                                                                    End If
                                                                    If ArrayAttrib.GetUpperBound(0) > 2 Then
                                                                        If FormReport.RadioButton5.Checked = True Then
                                                                            Dim MetodCoord As String = FuncCadastr.FuncGeopointOpredKod(ArrayAttrib(3).Trim)
                                                                            ArrayKoordinate(6, CountArrayKoordinate) = MetodCoord 'метод определения координат
                                                                        Else
                                                                            ArrayKoordinate(6, CountArrayKoordinate) = ArrayAttrib(3).Trim 'метод определения координат
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                CountArrayKoordinate += 1
                                            Next i1
                                            'если первая и последняя точка не совпадают то дублируем в массив первую координату
                                            If boolRegion = True Then
                                                If Not (ArrayKoordinate(1, CountCoordFirstVertex) Like ArrayKoordinate(1, CountArrayKoordinate - 1) And ArrayKoordinate(2, CountCoordFirstVertex) Like ArrayKoordinate(2, CountArrayKoordinate - 1)) Then
                                                    'повторяем координаты перевой точки
                                                    ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                                    ArrayKoordinate(0, CountArrayKoordinate) = ArrayKoordinate(0, CountCoordFirstVertex) ''номер точки 1, 2 ...
                                                    ArrayKoordinate(1, CountArrayKoordinate) = ArrayKoordinate(1, CountCoordFirstVertex) 'кооордината Х
                                                    ArrayKoordinate(2, CountArrayKoordinate) = ArrayKoordinate(2, CountCoordFirstVertex) 'координата Y
                                                    ArrayKoordinate(3, CountArrayKoordinate) = ArrayKoordinate(3, CountCoordFirstVertex) 'координата Z
                                                    ArrayKoordinate(4, CountArrayKoordinate) = ArrayKoordinate(4, CountCoordFirstVertex) 'дельта
                                                    ArrayKoordinate(5, CountArrayKoordinate) = ArrayKoordinate(5, CountCoordFirstVertex) 'код закрепления точки
                                                    ArrayKoordinate(6, CountArrayKoordinate) = ArrayKoordinate(6, CountCoordFirstVertex) 'метод определения координат
                                                    countVert += 1
                                                    CountArrayKoordinate += 1
                                                End If
                                            End If
                                            If boolRegion = True Then
                                                If j1 = 0 Then
                                                    ArrayKoordinate(2, CountCoordFirstVertex - 1) = "REGION" 'регион
                                                End If
                                            Else
                                                If ArrayKoordinate.GetUpperBound(1) > 1 Then
                                                    If j1 = 0 Then
                                                        If ArrayCounter.GetUpperBound(1) = 0 Then
                                                            ArrayKoordinate(2, CountCoordFirstVertex - 1) = "PLINE" 'полилиния
                                                        Else
                                                            ArrayKoordinate(2, CountCoordFirstVertex - 1) = "PLINE MULTIPLE" 'полилиния
                                                        End If
                                                    End If
                                                Else
                                                    ArrayKoordinate(2, CountCoordFirstVertex - 1) = "POINT" 'точка
                                                End If
                                            End If
                                            ArrayKoordinate(3, CountCoordFirstVertex - 1) = countVert 'записываем количество вершин (для экспорта в mapinfo)
                                            ArrayKoordinate(4, CountCoordFirstVertex - 1) = HandlePLine 'записываем handle контура (для экспорта в Autocad)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'достаем островки
                                            If IsArray(ArrayOstrov) = True Then
                                                For k As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                                                    countVert = 0 'сбрасываем счетчик контуров)
                                                    'получаем доступ к полилинии для извлечения координат
                                                    Dim HandleOstrov As String = ArrayOstrov(1, k)
                                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandleOstrov)
                                                    If IsNothing(idOstrov) = False Then
                                                        If idOstrov.IsNull = False Then
                                                            Dim entOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idOstrov.GetObject(OpenMode.ForRead)
                                                            If TypeOf entOstrov Is Polyline Then
                                                                Dim PlineOstrov As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                                                CountRegions += 1
                                                                Dim ptOstrovVertCollections As Point3dCollection = New Point3dCollection
                                                                For k1 = 0 To PlineOstrov.NumberOfVertices - 1
                                                                    If k1 = PlineOstrov.NumberOfVertices - 1 Then
                                                                        If PlineOstrov.GetPoint3dAt(0) = PlineOstrov.GetPoint3dAt(k1) Then
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                    ptOstrovVertCollections.Add(PlineOstrov.GetPoint3dAt(k1))
                                                                Next
                                                                Dim AreaOstrov As Double = MathFunction.FuncSQRTPolyline(ptOstrovVertCollections, RoundZn, AreaRoundZn)
                                                                SQRParcel = SQRParcel - AreaOstrov 'корректируем площадь участка
                                                                'определяем количество вершин в контуре
                                                                ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                                                ArrayKoordinate(0, CountArrayKoordinate) = "-1" 'флаг что это островок
                                                                CountArrayKoordinate += 1
                                                                CountCoordFirstVertex = CountArrayKoordinate 'фиксируем первую координату
                                                                'последовательно перебираем все вершины контура для записи в массив координат поворотных точек
                                                                For i1 As Integer = 0 To PlineOstrov.NumberOfVertices - 1
                                                                    Dim VertCoord As Point2d = PlineOstrov.GetPoint2dAt(i1)
                                                                    'записываем характеристики в массив
                                                                    ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                                                    ArrayKoordinate(0, CountArrayKoordinate) = "" 'номер вершины (пока пустое поле)
                                                                    ArrayKoordinate(1, CountArrayKoordinate) = VertCoord.X 'кооордината Х
                                                                    ArrayKoordinate(2, CountArrayKoordinate) = VertCoord.Y 'координата Y
                                                                    ArrayKoordinate(3, CountArrayKoordinate) = 0 'координата z
                                                                    ArrayKoordinate(4, CountArrayKoordinate) = "0.1" 'дельта
                                                                    ArrayKoordinate(5, CountArrayKoordinate) = "626003000000" 'код закрепления точки
                                                                    ArrayKoordinate(6, CountArrayKoordinate) = "692006000000" 'метод определения координат
                                                                    countVert += 1
                                                                    'находим номер точки
                                                                    If GeoPointCoolection.Count > 0 Then
                                                                        Dim FindPointID As ObjectId = FuncCadastr.FuncFindGeoPoint(GeoPointCoolection, VertCoord, FormReport.NumericUpDown3.Value)
                                                                        If IsNothing(FindPointID) = False Then
                                                                            If FindPointID.IsNull = False Then
                                                                                Dim entBlk As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(FindPointID, OpenMode.ForRead)
                                                                                If TypeOf entBlk Is BlockReference Then
                                                                                    Dim blk As BlockReference = entBlk
                                                                                    Dim ArrayAttrib As String() = Nothing
                                                                                    Dim FindPointGeo As Boolean = FuncCadastr.FuncReturnBlkArtibuteGeoPoint(blk, ArrayAttrib)
                                                                                    If IsArray(ArrayAttrib) = True Then
                                                                                        ArrayKoordinate(0, CountArrayKoordinate) = ArrayAttrib(0) 'номер точки 1, 2 ...
                                                                                        If ArrayAttrib.GetUpperBound(0) > 0 Then
                                                                                            ArrayKoordinate(4, CountArrayKoordinate) = ArrayAttrib(1) 'дельта
                                                                                        End If
                                                                                        If ArrayAttrib.GetUpperBound(0) > 1 Then
                                                                                            Dim KodZakr As String = FuncCadastr.FuncZakrepPointKod(ArrayAttrib(2).Trim)
                                                                                            ArrayKoordinate(5, CountArrayKoordinate) = KodZakr 'код закрепления точки
                                                                                        End If
                                                                                        If ArrayAttrib.GetUpperBound(0) > 2 Then
                                                                                            If FormReport.RadioButton5.Checked = True Then
                                                                                                Dim MetodCoord As String = FuncCadastr.FuncGeopointOpredKod(ArrayAttrib(3).Trim)
                                                                                                ArrayKoordinate(6, CountArrayKoordinate) = MetodCoord 'метод определения координат
                                                                                            Else
                                                                                                ArrayKoordinate(6, CountArrayKoordinate) = ArrayAttrib(3).Trim 'метод определения координат
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    CountArrayKoordinate += 1
                                                                Next i1
                                                                'если первая и последняя точка не совпадают то дублируем в массив первую координату
                                                                If Not (ArrayKoordinate(1, CountCoordFirstVertex) Like ArrayKoordinate(1, CountArrayKoordinate - 1) And ArrayKoordinate(2, CountCoordFirstVertex) Like ArrayKoordinate(2, CountArrayKoordinate - 1)) Then
                                                                    'повторяем координаты первой точки
                                                                    ReDim Preserve ArrayKoordinate(6, CountArrayKoordinate)
                                                                    ArrayKoordinate(0, CountArrayKoordinate) = ArrayKoordinate(0, CountCoordFirstVertex) 'номер точки 1, 2 ...
                                                                    ArrayKoordinate(1, CountArrayKoordinate) = ArrayKoordinate(1, CountCoordFirstVertex) 'кооордината Х
                                                                    ArrayKoordinate(2, CountArrayKoordinate) = ArrayKoordinate(2, CountCoordFirstVertex) 'координата Y
                                                                    ArrayKoordinate(3, CountArrayKoordinate) = ArrayKoordinate(3, CountCoordFirstVertex) 'координата Y
                                                                    ArrayKoordinate(4, CountArrayKoordinate) = ArrayKoordinate(4, CountCoordFirstVertex) '
                                                                    ArrayKoordinate(5, CountArrayKoordinate) = ArrayKoordinate(5, CountCoordFirstVertex) '
                                                                    ArrayKoordinate(6, CountArrayKoordinate) = ArrayKoordinate(6, CountCoordFirstVertex) '
                                                                    CountArrayKoordinate += 1
                                                                    countVert += 1
                                                                End If
                                                                ArrayKoordinate(3, CountCoordFirstVertex - 1) = countVert 'записываем количество вершин
                                                                ArrayKoordinate(4, CountCoordFirstVertex - 1) = HandleOstrov 'записываем handle контура (для экспорта в Autocad)
                                                            End If
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                            TotalArea += SQRParcel
                                            TotalLenght += LenghtObj
                                            CountRegions += 1
                                        End If
                                    End If
                                Next j1
                                If boolRegion = True Then
                                    ArrayKoordinate(5, 0) = TotalArea 'записываем площадь
                                Else
                                    ArrayKoordinate(5, 0) = TotalLenght 'записываем протяженность
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'записываем результат
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'таблица автокад
                                If FormReport.RadioButton1.Checked = True Then
                                    If FormReport.ComboBox1.Text Like "Таблица координат" Then
                                        Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                                        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForWrite)
                                        Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                        Dim countRows As Integer = 1
                                        Dim countRowsConst As Integer = 1
                                        tb.TableStyle = db.Tablestyle
                                        If FormReport.CheckBox8.Checked = True Then
                                            tb.SetSize(3, 5)
                                            tb.SetRowHeight(8)
                                            tb.Rows(0).TextHeight = 2
                                            tb.Rows(1).TextHeight = 2
                                            tb.Rows(2).TextHeight = 2
                                            tb.Columns(0).Width = 15 'номер точки
                                            tb.Columns(1).Width = 20 'X
                                            tb.Columns(2).Width = 20 'Y
                                            tb.Columns(3).Width = 40 'Метод определения
                                            tb.Columns(4).Width = 20 'Погрешность
                                        Else
                                            tb.SetSize(3, 3)
                                            tb.SetRowHeight(8)
                                            tb.Rows(0).TextHeight = 2
                                            tb.Rows(1).TextHeight = 2
                                            tb.Rows(2).TextHeight = 2
                                            tb.Columns(0).Width = 25 'номер точки
                                            tb.Columns(1).Width = 20 'X
                                            tb.Columns(2).Width = 20 'Y
                                        End If
                                        If j = 0 Then
                                            Dim ppr As PromptPointResult = ed.GetPoint("Укажите точку вставки таблицы: ")
                                            If (ppr.Status = PromptStatus.OK) Then
                                                PointInsTbl1 = ppr.Value
                                            End If
                                        Else
                                            Dim X As Double = PointInsTbl1.X + 100
                                            Dim Y As Double = PointInsTbl1.Y
                                            PointInsTbl1 = New Point3d(X, Y, 0)
                                        End If
                                        tb.Position = PointInsTbl1
                                        'заголовок таблицы
                                        Dim range As CellRange = CellRange.Create(tb, 1, 0, 2, 0)
                                        tb.MergeCells(range)
                                        tb.Cells(1, 0).Value = "Обозначение характерных точек границ"
                                        Dim range1 As CellRange = CellRange.Create(tb, 1, 1, 1, 2)
                                        tb.MergeCells(range1)
                                        tb.Cells(1, 1).Value = "Координаты, м"
                                        tb.Cells(2, 1).Value = "X, м"
                                        tb.Cells(2, 2).Value = "Y, м"
                                        If FormReport.CheckBox8.Checked = True Then
                                            tb.Cells(1, 3).Value = "Метод определения координат"
                                            Dim range2 As CellRange = CellRange.Create(tb, 1, 3, 2, 3)
                                            tb.MergeCells(range2)
                                            tb.Cells(1, 4).Value = "Средняя квадратическая погрешность положения характерных точек (Mt), м"
                                            Dim range3 As CellRange = CellRange.Create(tb, 1, 4, 2, 4)
                                            tb.MergeCells(range3)
                                        End If
                                        ms.AppendEntity(tb)
                                        tr.AddNewlyCreatedDBObject(tb, True)
                                        'присваиваем расширенные данные таблице
                                        Dim NameTabRec As String = "CADASTR_OBJECT"
                                        Dim ArrayOBJ As String(,) = Nothing
                                        ReDim Preserve ArrayOBJ(2, 2)
                                        ArrayOBJ(0, 0) = "CADASTRAL_OBJECT"
                                        ArrayOBJ(1, 0) = "LABEL"
                                        ArrayOBJ(2, 0) = "Таблица координат"
                                        ArrayOBJ(0, 1) = "CADASTRAL_OBJECT"
                                        ArrayOBJ(1, 1) = "Number"
                                        ArrayOBJ(2, 1) = "1"
                                        ArrayOBJ(0, 2) = "CADASTRAL_OBJECT"
                                        ArrayOBJ(1, 2) = "IDRNG"
                                        If idGlobalParcel.Trim.Length > 2 Then
                                            ArrayOBJ(2, 2) = idGlobalParcel
                                        Else
                                            ArrayOBJ(2, 2) = LocalIdParcel
                                        End If
                                        If IsNothing(IDTablePS) = False Then
                                            If IDTablePS.IsNull = False Then
                                                Dim boolAddTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(tb, IDTablePS)
                                                Dim boolXrecord As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(tb, ArrayOBJ)
                                            End If
                                        End If
                                        '===========================================================================================================
                                        Dim flagArea As Boolean = False
                                        tb.Cells(0, 0).Value = "Каталог координат поворотных точек участка с условным номером: " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        If IsArray(ArrayKoordinate) = True Then
                                            Dim count As Integer = 3
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    tb.InsertRows(count, 1, 1)
                                                    tb.Rows(count).TextHeight = 2
                                                    Dim range2 As CellRange = CellRange.Create(tb, count, 0, count, 2)
                                                    tb.MergeCells(range2)
                                                    tb.Cells(count, 0).Value = ""
                                                    count += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    tb.InsertRows(count, 1, 1)
                                                    tb.Rows(count).TextHeight = 2
                                                    Dim range2 As CellRange = CellRange.Create(tb, count, 0, count, 2)
                                                    tb.MergeCells(range2)
                                                    tb.Cells(count, 0).Value = "Контур №" & nt
                                                    count += 1
                                                Else
                                                    tb.InsertRows(count, 1, 1)
                                                    tb.Rows(count).TextHeight = 2
                                                    tb.Cells(count, 0).Value = nt
                                                    Dim X As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        xi = Math.Round(xi, RoundZn)
                                                        tb.Cells(count, 2).Value = MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        yi = Math.Round(yi, RoundZn)
                                                        tb.Cells(count, 1).Value = MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    End If
                                                    If FormReport.CheckBox8.Checked = True Then
                                                        If tb.Columns.Count = 5 Then
                                                            Dim Metod As String = ArrayKoordinate(6, j1)
                                                            If IsNothing(Metod) = False Then
                                                                tb.Cells(count, 3).Value = Metod
                                                            End If
                                                            Dim DeltaXY As String = ArrayKoordinate(4, j1)
                                                            If IsNothing(DeltaXY) = False Then
                                                                Dim di As Double = Val(DeltaXY)
                                                                di = Math.Round(di, RoundZn)
                                                                If di = 0 Then
                                                                    di = 0.1
                                                                End If
                                                                tb.Cells(count, 4).Value = MathFunction.FuncFormatNumb(di, RoundZn)
                                                            End If
                                                        End If
                                                    End If
                                                    count += 1
                                                End If
                                            Next j1
                                            If FormReport.CheckBox7.Checked = False Then
                                                If flagArea = False Then
                                                    tb.InsertRows(count, 1, 1)
                                                    tb.Rows(count).TextHeight = 2
                                                    tb.Cells(count, 0).Value = "Площадь: "
                                                    Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                    tb.Cells(count, 2).Value = Math.Round(TotalArea, AreaRoundZn) & " кв.м"
                                                End If
                                            End If
                                        End If
                                    ElseIf FormReport.ComboBox1.Text Like "Участок в отдельный файл" Then
                                        Dim NameDoc As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        NameDoc = NameDoc.Replace("_pxqc;", "")
                                        NameDoc = NameDoc.Replace("\", "_")
                                        NameDoc = NameDoc.Replace("/", "_")
                                        NameDoc = NameDoc.Replace(":", "_")
                                        NameDoc = NameDoc.Replace("*", "_")
                                        NameDoc = NameDoc.Replace(".", "_")
                                        NameDoc = NameDoc.Replace(",", "_")
                                        If j = 0 Then
                                            'указываем папку для сохранения
                                            NameDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите директорию", 1)
                                        End If
                                        Dim strSave As String = NameDir & "\" & NameDoc & ".dwg"
                                        '=========================================================================================
                                        Dim ArrayFiles As String() = IO.Directory.GetFiles(NameDir)
                                        If IsArray(ArrayFiles) = True Then
                                            For k1 As Integer = 0 To ArrayFiles.Length - 1
                                                Dim ffile As String = ArrayFiles(k1)
                                                If ffile Like strSave Then
                                                    Dim rng As Random = New Random
                                                    NameDoc = NameDoc & rng.Next
                                                    strSave = NameDir & "\" & NameDoc & ".dwg"
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        '=============================================================================================
                                        Dim strTemplatePath As String = parentPatch2 & "\TemplateDWG\Шаблон_acadiso.dwg"
                                        If File.Exists(strTemplatePath) = False Then
                                            MsgBox("Шаблон Шаблон_acadiso.dwg в папке " & parentPatch2 & "\TemplateDWG\ " & ", не найден!!!")
                                            Exit Sub
                                        End If
                                        '=================================================================================================
                                        'заносим объекты в коллекцию
                                        Dim idColl As ObjectIdCollection = New ObjectIdCollection
                                        For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                            Dim nt As String = ArrayKoordinate(0, j1)
                                            If IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt.Length > 0 Then
                                                Dim hgObject As String = ArrayKoordinate(4, j1)
                                                Dim idObject As ObjectId = AcadFunction.FuncReturnObjectID(hgObject)
                                                If IsNothing(idObject) = False Then
                                                    If idObject.IsNull = False Then
                                                        idColl.Add(idObject)
                                                    End If
                                                End If
                                            End If
                                        Next j1
                                        '======================================================================================================
                                        'если коллекция не пустая тогла копируем объекты
                                        If idColl.Count > 0 Then
                                            Using newDb As Database = New Database(False, False)
                                                newDb.ReadDwgFile(strTemplatePath, FileOpenMode.OpenForReadAndWriteNoShare, True, Nothing)
                                                newDb.CloseInput(True)
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
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'ТАБЛИЦА WORD
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                ElseIf FormReport.RadioButton3.Checked = True Then
                                    If Not (FormReport.ComboBox1.Text Like "Новый файл docx") Then
                                        WDoc.Range(SetPos).Font.Name = "Times New Roman"
                                        WDoc.Range(SetPos).Font.Size = 12
                                        '================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем данные
                                        Dim rng As Word.Range = WDoc.Range(SetPos + 1, SetPos + 1)
                                        WDoc.Range(SetPos + 1, SetPos + 1).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        Dim text1 As String = "Условный номер (обозначение): " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        rng.InsertAfter(text1)
                                        Dim l1 As Integer = text1.Length
                                        SetPos = SetPos + l1 + 1
                                        Dim range As Word.Range = WDoc.Range(SetPos, SetPos)
                                        range.Paragraphs.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем таблицу
                                        Dim WDocTbl As Word.Table = Nothing
                                        If FormReport.CheckBox8.Checked = True Then
                                            WDocTbl = WDoc.Tables.Add(range, 3, 5, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)
                                        Else
                                            WDocTbl = WDoc.Tables.Add(range, 3, 3, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)
                                        End If
                                        WDocTbl.Style = "Сетка таблицы"
                                        WDocTbl.ApplyStyleRowBands = True
                                        WDocTbl.Cell(1, 1).Merge(MergeTo:=WDocTbl.Cell(2, 1))
                                        WDocTbl.Cell(1, 1).Range.Text = "Обозначение" & vbLf & "характерных" & vbLf & "точек границ"
                                        WDocTbl.Cell(1, 1).Range.Bold = 1
                                        WDocTbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(1, 1).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(1, 2).Merge(MergeTo:=WDocTbl.Cell(1, 3))
                                        WDocTbl.Cell(1, 2).Range.Text = "Координаты, м"
                                        WDocTbl.Cell(1, 2).Range.Bold = 1
                                        WDocTbl.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(1, 2).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(2, 2).Range.Text = "X"
                                        WDocTbl.Cell(2, 2).Range.Bold = 1
                                        WDocTbl.Cell(2, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(2, 2).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(2, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(2, 3).Range.Text = "Y"
                                        WDocTbl.Cell(2, 3).Range.Bold = 1
                                        WDocTbl.Cell(2, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(2, 3).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(2, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        If FormReport.CheckBox8.Checked = True Then
                                            If WDocTbl.Columns.Count = 5 Then
                                                WDocTbl.Cell(1, 3).Merge(MergeTo:=WDocTbl.Cell(2, 4))
                                                WDocTbl.Cell(1, 3).Range.Text = "Метод определения координат"
                                                WDocTbl.Cell(1, 3).Range.Bold = 1
                                                WDocTbl.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                WDocTbl.Cell(1, 3).Range.Font.Name = "Times New Roman"
                                                WDocTbl.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                                WDocTbl.Cell(1, 4).Merge(MergeTo:=WDocTbl.Cell(2, 5))
                                                WDocTbl.Cell(1, 4).Range.Text = "Средняя квадратическая погрешность положения характерных точек (Mt), м"
                                                WDocTbl.Cell(1, 4).Range.Bold = 1
                                                WDocTbl.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                WDocTbl.Cell(1, 4).Range.Font.Name = "Times New Roman"
                                                WDocTbl.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        If IsArray(ArrayKoordinate) = True Then
                                            Dim count As Integer = 3
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    WDocTbl.Rows.Add()
                                                    WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                                    WDocTbl.Cell(count, 1).Range.Text = ""
                                                    count += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt.Length > 0 Then
                                                    WDocTbl.Rows.Add()
                                                    WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                                    WDocTbl.Cell(count, 1).Range.Text = "Контур №" & nt
                                                    WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                    WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                                                    WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    count += 1
                                                Else
                                                    If j1 = ArrayKoordinate.GetUpperBound(1) Then
                                                        If FormReport.CheckBox7.Checked = False Then
                                                            WDocTbl.Rows.Add()
                                                        End If
                                                    Else
                                                        WDocTbl.Rows.Add()
                                                    End If
                                                    WDocTbl.Cell(count, 1).Range.Text = nt
                                                    Dim X As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        xi = Math.Round(xi, RoundZn)
                                                        WDocTbl.Cell(count, 2).Range.Text = MathFunction.FuncFormatNumb(xi, RoundZn)
                                                        WDocTbl.Cell(count, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 2).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        yi = Math.Round(yi, RoundZn)
                                                        WDocTbl.Cell(count, 3).Range.Text = MathFunction.FuncFormatNumb(yi, RoundZn)
                                                        WDocTbl.Cell(count, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 3).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    If WDocTbl.Columns.Count = 5 Then
                                                        Dim metod As String = ArrayKoordinate(6, j1)
                                                        If IsNothing(metod) = True Then metod = "-"
                                                        If metod.Trim.Length = 0 Then
                                                            metod = "-"
                                                        End If
                                                        WDocTbl.Cell(count, 4).Range.Text = metod
                                                        WDocTbl.Cell(count, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 4).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                                        Dim deltaYX As String = ArrayKoordinate(4, j1)
                                                        If IsNothing(deltaYX) = True Then deltaYX = "0"
                                                        Dim di As Double = Val(deltaYX)
                                                        If di = 0 Then di = 0.1
                                                        di = Math.Round(di, RoundZn)
                                                        WDocTbl.Cell(count, 5).Range.Text = MathFunction.FuncFormatNumb(di, RoundZn)
                                                        WDocTbl.Cell(count, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 5).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    count += 1
                                                End If
                                            Next j1
                                            If FormReport.CheckBox7.Checked = False Then
                                                WDocTbl.Cell(count, 1).Range.Text = "Площадь: "
                                                Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                WDocTbl.Cell(count, 3).Range.Text = Math.Round(TotalArea, AreaRoundZn) & " кв.м"
                                            End If
                                        End If
                                        Try
                                            If idGlobalParcel.Trim.Length > 2 Then
                                                WDocTbl.Descr = "IDRNG=" & idGlobalParcel
                                            Else
                                                WDocTbl.Descr = "IDRNG=" & LocalIdParcel
                                            End If
                                        Catch ex As System.NotImplementedException
                                        End Try

                                        WDocTbl.Range.Select()
                                        WDocTbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                                        WDoc.Range(WDocTbl.Range.Start, WDocTbl.Range.End).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        SetPos = WDocTbl.Range.End
                                        rng = WDoc.Range(0, SetPos + 1)
                                        rng.Paragraphs.Add()
                                    Else
                                        Dim endPosDocument As Integer = WDoc.Content.End
                                        '===========================================================================================
                                        'форматируем данные
                                        Dim deltaSetPos As Integer = endPosDocument - SetPos
                                        If deltaSetPos <= 0 Then
                                            SetPos = SetPos + deltaSetPos - 1
                                        End If
                                        WDoc.Range(SetPos).Font.Name = "Times New Roman"
                                        WDoc.Range(SetPos).Font.Size = 12
                                        WDoc.Range(SetPos).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                        '================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем данные
                                        Dim rng As Word.Range = WDoc.Range(SetPos)
                                        Dim text1 As String = "Условный номер (обозначение): " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        rng = WDoc.Range(SetPos)
                                        rng.InsertAfter(text1)
                                        Dim l1 As Integer = text1.Length
                                        SetPos = SetPos + l1 + 1
                                        endPosDocument = WDoc.Content.End
                                        deltaSetPos = endPosDocument - SetPos
                                        If deltaSetPos <= 0 Then
                                            SetPos = SetPos + deltaSetPos - 1
                                        End If
                                        Dim range As Word.Range = WDoc.Range(SetPos)
                                        range.Paragraphs.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем таблицу
                                        Dim WDocTbl As Word.Table = Nothing
                                        If FormReport.CheckBox8.Checked = True Then
                                            WDocTbl = WDoc.Tables.Add(range, 3, 5, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)
                                        Else
                                            WDocTbl = WDoc.Tables.Add(range, 3, 3, Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitContent)
                                        End If
                                        WDocTbl.Style = "Сетка таблицы"
                                        WDocTbl.ApplyStyleRowBands = True
                                        WDocTbl.Cell(1, 1).Merge(MergeTo:=WDocTbl.Cell(2, 1))
                                        WDocTbl.Cell(1, 1).Range.Text = "Обозначение" & vbLf & "характерных" & vbLf & "точек границ"
                                        WDocTbl.Cell(1, 1).Range.Bold = 1
                                        WDocTbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(1, 1).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(1, 2).Merge(MergeTo:=WDocTbl.Cell(1, 3))
                                        WDocTbl.Cell(1, 2).Range.Text = "Координаты, м"
                                        WDocTbl.Cell(1, 2).Range.Bold = 1
                                        WDocTbl.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(1, 2).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(2, 2).Range.Text = "X"
                                        WDocTbl.Cell(2, 2).Range.Bold = 1
                                        WDocTbl.Cell(2, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(2, 2).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(2, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        WDocTbl.Cell(2, 3).Range.Text = "Y"
                                        WDocTbl.Cell(2, 3).Range.Bold = 1
                                        WDocTbl.Cell(2, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                        WDocTbl.Cell(2, 3).Range.Font.Name = "Times New Roman"
                                        WDocTbl.Cell(2, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        If FormReport.CheckBox8.Checked = True Then
                                            If WDocTbl.Columns.Count = 5 Then
                                                WDocTbl.Cell(1, 3).Merge(MergeTo:=WDocTbl.Cell(2, 4))
                                                WDocTbl.Cell(1, 3).Range.Text = "Метод определения координат"
                                                WDocTbl.Cell(1, 3).Range.Bold = 1
                                                WDocTbl.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                WDocTbl.Cell(1, 3).Range.Font.Name = "Times New Roman"
                                                WDocTbl.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                                WDocTbl.Cell(1, 4).Merge(MergeTo:=WDocTbl.Cell(2, 5))
                                                WDocTbl.Cell(1, 4).Range.Text = "Средняя квадратическая погрешность положения характерных точек (Mt), м"
                                                WDocTbl.Cell(1, 4).Range.Bold = 1
                                                WDocTbl.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                WDocTbl.Cell(1, 4).Range.Font.Name = "Times New Roman"
                                                WDocTbl.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        If IsArray(ArrayKoordinate) = True Then
                                            Dim count As Integer = 3
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    WDocTbl.Rows.Add()
                                                    WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                                    WDocTbl.Cell(count, 1).Range.Text = ""
                                                    count += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt.Length > 0 Then
                                                    WDocTbl.Rows.Add()
                                                    WDocTbl.Cell(count, 1).Merge(MergeTo:=WDocTbl.Cell(count, 3))
                                                    WDocTbl.Cell(count, 1).Range.Text = "Контур №" & nt
                                                    WDocTbl.Cell(count, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                    WDocTbl.Cell(count, 1).Range.Font.Name = "Times New Roman"
                                                    WDocTbl.Cell(count, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    count += 1
                                                Else
                                                    If j1 = ArrayKoordinate.GetUpperBound(1) Then
                                                        If FormReport.CheckBox7.Checked = False Then
                                                            WDocTbl.Rows.Add()
                                                        End If
                                                    Else
                                                        WDocTbl.Rows.Add()
                                                    End If
                                                    WDocTbl.Cell(count, 1).Range.Text = nt
                                                    Dim X As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        xi = Math.Round(xi, RoundZn)
                                                        WDocTbl.Cell(count, 2).Range.Text = MathFunction.FuncFormatNumb(xi, RoundZn)
                                                        WDocTbl.Cell(count, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 2).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        yi = Math.Round(yi, RoundZn)
                                                        WDocTbl.Cell(count, 3).Range.Text = MathFunction.FuncFormatNumb(yi, RoundZn)
                                                        WDocTbl.Cell(count, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 3).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    If WDocTbl.Columns.Count = 5 Then
                                                        Dim metod As String = ArrayKoordinate(6, j1)
                                                        If IsNothing(metod) = True Then metod = "-"
                                                        If metod.Trim.Length = 0 Then
                                                            metod = "-"
                                                        End If
                                                        WDocTbl.Cell(count, 4).Range.Text = metod
                                                        WDocTbl.Cell(count, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 4).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                                        Dim deltaYX As String = ArrayKoordinate(4, j1)
                                                        If IsNothing(deltaYX) = True Then deltaYX = "0"
                                                        Dim di As Double = Val(deltaYX)
                                                        If di = 0 Then di = 0.1
                                                        di = Math.Round(di, RoundZn)
                                                        WDocTbl.Cell(count, 5).Range.Text = MathFunction.FuncFormatNumb(di, RoundZn)
                                                        WDocTbl.Cell(count, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                                        WDocTbl.Cell(count, 5).Range.Font.Name = "Times New Roman"
                                                        WDocTbl.Cell(count, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                                                    End If
                                                    count += 1
                                                End If
                                            Next j1
                                            If FormReport.CheckBox7.Checked = False Then
                                                WDocTbl.Cell(count, 1).Range.Text = "Площадь: "
                                                Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                WDocTbl.Cell(count, 3).Range.Text = Math.Round(TotalArea, AreaRoundZn) & " кв.м"
                                            End If
                                        End If
                                        Try
                                            If idGlobalParcel.Trim.Length > 2 Then
                                                WDocTbl.Descr = "IDRNG=" & idGlobalParcel
                                            Else
                                                WDocTbl.Descr = "IDRNG=" & LocalIdParcel
                                            End If
                                        Catch ex As System.NotImplementedException
                                        End Try

                                        WDocTbl.Range.Select()
                                        WDocTbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                                        WDoc.Range(WDocTbl.Range.Start, WDocTbl.Range.End).Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                                        SetPos = WDocTbl.Range.End
                                        rng = WDoc.Range(SetPos)
                                        rng.Paragraphs.Add()
                                    End If

                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'отчет в Excel
                                ElseIf FormReport.RadioButton2.Checked = True Then
                                    Dim userNumbFormat As String = "0.00"
                                    For i As Integer = 0 To RoundZn
                                        If i = 0 Then
                                            userNumbFormat = "0"
                                        ElseIf i = 1 Then
                                            userNumbFormat = userNumbFormat & ".0"
                                        Else
                                            userNumbFormat = userNumbFormat & "0"
                                        End If
                                    Next i
                                    Dim userSQRTFormat As String = "0.00"
                                    For i As Integer = 0 To AreaRoundZn
                                        If i = 0 Then
                                            userSQRTFormat = "0"
                                        ElseIf i = 1 Then
                                            userSQRTFormat = userSQRTFormat & ".0"
                                        Else
                                            userSQRTFormat = userSQRTFormat & "0"
                                        End If
                                    Next i
                                    '================================================================================
                                    If FormReport.ComboBox1.Text Like "Все участки в один файл" Then
                                        boolExVisibl = True
                                        StartRows += 1
                                        Dim cell1 As Excel.Range 'первая ячейка в строке
                                        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                                        Dim CellRng As Excel.Range 'массив ячеек
                                        xlSheets = xlBook.ActiveSheet
                                        If StartRows = 2 Then
                                            xlSheets.Columns.Range("A:A").ColumnWidth = 16
                                            xlSheets.Columns.Range("B:B").ColumnWidth = 13
                                            xlSheets.Columns.Range("C:C").ColumnWidth = 13
                                            If FormReport.CheckBox8.Checked = True Then
                                                xlSheets.Columns.Range("D:D").ColumnWidth = 25
                                                xlSheets.Columns.Range("E:E").ColumnWidth = 15
                                            End If
                                        End If
                                        Dim flagArea As Boolean = False
                                        If FormReport.CheckBox8.Checked = True Then
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 6 ячейку в строке
                                        Else
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 6 ячейку в строке
                                        End If
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Условный номер (обозначение): " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        '===========================================================================
                                        cell1 = xlSheets.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
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
                                        cell1 = xlSheets.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Координаты, м"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "X"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Y"
                                        If FormReport.CheckBox8.Checked = True Then
                                            '===========================================================================
                                            cell1 = xlSheets.Cells(StartRows + 1, 4) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows + 2, 4) 'ссылка на 6 ячейку в строке
                                            CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                            CellRng.Merge(Type.Missing)
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 1
                                            cell1.Font.Size = 12
                                            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                            cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                            CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            cell1.Value = "Метод определения координат"
                                            CellRng.WrapText = True

                                            '===========================================================================
                                            cell1 = xlSheets.Cells(StartRows + 1, 5) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows + 2, 5) 'ссылка на 6 ячейку в строке
                                            CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                            CellRng.Merge(Type.Missing)
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 1
                                            cell1.Font.Size = 12
                                            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                            cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                            CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            cell1.Value = "Средняя квадратическая погрешность положения характерных точек (Mt), м"
                                            CellRng.WrapText = True
                                        End If
                                        '====================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        StartRows = StartRows + 3
                                        If IsArray(ArrayKoordinate) = True Then
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    'островок1
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = ""
                                                    StartRows += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    'Контур
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = "Контур №" & nt
                                                    StartRows += 1
                                                Else
                                                    'Координаты
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = nt
                                                    Dim X As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        'xi = Math.Round(xi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 3)
                                                        cell1.NumberFormat = userNumbFormat
                                                        cell1.Value = xi 'MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        'yi = Math.Round(yi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 2)
                                                        cell1.NumberFormat = userNumbFormat
                                                        cell1.Value = yi 'MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    End If
                                                    If FormReport.CheckBox8.Checked = True Then
                                                        cell1 = xlSheets.Cells(StartRows, 4) 'ссылка на первую ячейку
                                                        cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                        CellRng.Font.Name = "Times New Roman"
                                                        CellRng.Font.Size = 12
                                                        CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                        CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                        Dim metod As String = ArrayKoordinate(6, j1)
                                                        If IsNothing(metod) = True Then
                                                            metod = "Геодезический метод"
                                                        End If
                                                        If metod.Trim.Length = 0 Then
                                                            metod = "Геодезический метод"
                                                        End If
                                                        xlSheets.Cells(StartRows, 4).Value = metod
                                                        Dim dxy As String = ArrayKoordinate(4, j1)
                                                        If IsNothing(dxy) = True Then
                                                            dxy = "0.1"
                                                        End If
                                                        If dxy.Trim.Length = 0 Then
                                                            dxy = "0.1"
                                                        End If
                                                        If IsNothing(dxy) = False Then
                                                            Dim d As Double = Val(dxy)
                                                            'yi = Math.Round(yi, 2)
                                                            cell1 = xlSheets.Cells(StartRows, 5)
                                                            cell1.NumberFormat = userNumbFormat
                                                            cell1.Value = dxy 'MathFunction.FuncFormatNumb(yi, RoundZn)
                                                        End If
                                                    End If
                                                    StartRows += 1
                                                End If
                                            Next j1
                                            If FormReport.CheckBox7.Checked = False Then
                                                If FormReport.CheckBox8.Checked = True Then
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = "Площадь: "
                                                    Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                    cell1 = xlSheets.Cells(StartRows, 4)
                                                    cell1.NumberFormat = AreaRoundZn
                                                    xlSheets.Cells(StartRows, 2).Value = TotalArea
                                                    xlSheets.Cells(StartRows, 3).Value = "кв.м"
                                                Else
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = "Площадь: "
                                                    Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                    cell1 = xlSheets.Cells(StartRows, 2)
                                                    cell1.NumberFormat = AreaRoundZn
                                                    xlSheets.Cells(StartRows, 2).Value = TotalArea
                                                    xlSheets.Cells(StartRows, 3).Value = "кв.м"
                                                End If
                                                StartRows += 1

                                            End If
                                        End If
                                    ElseIf FormReport.ComboBox1.Text Like "Каждый участок в отдельном листе" Then
                                        boolExVisibl = True
                                        Dim cell1 As Excel.Range 'первая ячейка в строке
                                        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                                        Dim CellRng As Excel.Range 'массив ячеек
                                        'создаем новый лист
                                        Dim NameSheet As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        NameSheet = NameSheet.Replace("_pxqc;", "")
                                        NameSheet = NameSheet.Replace("\", "_")
                                        NameSheet = NameSheet.Replace("/", "_")
                                        NameSheet = NameSheet.Replace(":", "_")
                                        NameSheet = NameSheet.Replace("*", "_")
                                        NameSheet = NameSheet.Replace(".", "_")
                                        NameSheet = NameSheet.Replace(",", "_")
                                        xlSheets = xlBook.Sheets.Add()
                                        Try
                                            xlSheets.Name = NameSheet
                                        Catch ex As Exception
                                        End Try
                                        StartRows += 1
                                        If StartRows = 2 Then
                                            xlSheets.Columns.Range("A:A").ColumnWidth = 16
                                            xlSheets.Columns.Range("B:B").ColumnWidth = 13
                                            xlSheets.Columns.Range("C:C").ColumnWidth = 13
                                        End If
                                        Dim flagArea As Boolean = False
                                        If FormReport.CheckBox1.Checked = True Then
                                            Dim text1 As String = "Условный номер участка: " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 0
                                            cell1.Font.Size = 12
                                            cell1.Value = text1
                                            StartRows += 1
                                        End If
                                        If FormReport.CheckBox1.Checked = False Then
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 6 ячейку в строке
                                            CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                            CellRng.Merge(Type.Missing)
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 1
                                            cell1.Font.Size = 12
                                            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                            cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                            CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            cell1.Value = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        End If
                                        '===========================================================================
                                        cell1 = xlSheets.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
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
                                        cell1 = xlSheets.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Координаты, м"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "X"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Y"
                                        '====================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        StartRows = StartRows + 3
                                        If IsArray(ArrayKoordinate) = True Then
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    'островок
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = ""
                                                    StartRows += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    'Контур
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = "Контур №" & nt
                                                    StartRows += 1
                                                Else
                                                    'Координаты
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = nt
                                                    Dim X As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        xi = Math.Round(xi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 2)
                                                        cell1.Value = MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        yi = Math.Round(yi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 3)
                                                        cell1.Value = MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    End If
                                                    StartRows += 1
                                                End If
                                            Next j1
                                            If flagArea = False Then
                                                cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                CellRng.Font.Name = "Times New Roman"
                                                CellRng.Font.Size = 12
                                                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                xlSheets.Cells(StartRows, 1).Value = "Площадь "
                                                Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                cell1 = xlSheets.Cells(StartRows, 1)
                                                xlSheets.Cells(StartRows, 3).Value = Math.Round(TotalArea, AreaRoundZn) & " кв.м"
                                                StartRows += 1
                                            End If
                                        End If
                                        StartRows = 1
                                    ElseIf FormReport.ComboBox1.Text Like "Каждый участок в отдельный файл" Then
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        boolExVisibl = False
                                        Dim NameDoc As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        NameDoc = NameDoc.Replace("_pxqc;", "")
                                        NameDoc = NameDoc.Replace("\", "_")
                                        NameDoc = NameDoc.Replace("/", "_")
                                        NameDoc = NameDoc.Replace(":", "_")
                                        NameDoc = NameDoc.Replace("*", "_")
                                        NameDoc = NameDoc.Replace(".", "_")
                                        NameDoc = NameDoc.Replace(",", "_")
                                        If j > 0 Then
                                            xlBook = xlApp.Workbooks.Add()
                                        Else
                                            'указываем папку для сохранения
                                            NameDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите директорию", 1)
                                            If IO.Directory.Exists(NameDir) = False Then
                                                xlBook.Close()
                                                xlBook = Nothing
                                                xlApp.Quit()
                                                xlApp = Nothing
                                                Exit Sub
                                            End If
                                        End If
                                        'активный лист
                                        xlSheets = xlBook.ActiveSheet
                                        Dim strSave As String = NameDir & "\" & NameDoc & ".xlsx"
                                        '=========================================================================================
                                        Dim ArrayFiles As String() = IO.Directory.GetFiles(NameDir)
                                        If IsArray(ArrayFiles) = True Then
                                            For k1 As Integer = 0 To ArrayFiles.Length - 1
                                                Dim ffile As String = ArrayFiles(k1)
                                                If ffile Like strSave Then
                                                    Dim rng As Random = New Random
                                                    NameDoc = NameDoc & rng.Next
                                                    strSave = NameDir & "\" & NameDoc & ".xlsx"
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        Dim cell1 As Excel.Range 'первая ячейка в строке
                                        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                                        Dim CellRng As Excel.Range 'массив ячеек
                                        StartRows += 1
                                        If StartRows = 2 Then
                                            xlSheets.Columns.Range("A:A").ColumnWidth = 16
                                            xlSheets.Columns.Range("B:B").ColumnWidth = 13
                                            xlSheets.Columns.Range("C:C").ColumnWidth = 13
                                        End If
                                        Dim FlagArea As Boolean = False
                                        If FormReport.CheckBox1.Checked = True Then
                                            Dim text1 As String = "Условный номер участка: " & NameDoc
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 0
                                            cell1.Font.Size = 12
                                            cell1.Value = text1
                                            StartRows += 1
                                        End If
                                        If FormReport.CheckBox1.Checked = False Then
                                            cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 6 ячейку в строке
                                            CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                            CellRng.Merge(Type.Missing)
                                            cell1.Font.Name = "Times New Roman"
                                            cell1.Font.Bold = 1
                                            cell1.Font.Size = 12
                                            cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                            cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                            CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            cell1.Value = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        End If
                                        '===========================================================================
                                        cell1 = xlSheets.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
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
                                        cell1 = xlSheets.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Координаты, м"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "X"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Y"
                                        '====================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем координаты
                                        StartRows = StartRows + 3
                                        If IsArray(ArrayKoordinate) = True Then
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    'островок
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = ""
                                                    StartRows += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    'Контур
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = "Контур №" & nt
                                                    StartRows += 1
                                                Else
                                                    'Координаты
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = nt
                                                    Dim X As String = ArrayKoordinate(1, j)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        xi = Math.Round(xi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 2)
                                                        cell1.Value = MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(2, j)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        yi = Math.Round(yi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 3)
                                                        cell1.Value = MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    End If
                                                    StartRows += 1
                                                End If
                                            Next j1
                                            If FlagArea = False Then
                                                cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                cell2 = xlSheets.Cells(StartRows, 3) 'ссылка на 3 ячейку в строке
                                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                CellRng.Font.Name = "Times New Roman"
                                                CellRng.Font.Size = 12
                                                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                xlSheets.Cells(StartRows, 1).Value = "Площадь "
                                                Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                cell1 = xlSheets.Cells(StartRows, 1)
                                                xlSheets.Cells(StartRows, 3).Value = Math.Round(TotalArea, AreaRoundZn) & " кв.м"
                                                StartRows += 1
                                            End If
                                        End If
                                        StartRows = 1
                                        xlBook.SaveAs(strSave)
                                        xlBook.Close()
                                        xlBook = Nothing
                                    ElseIf FormReport.ComboBox1.Text Like "МСК-WGS84" Then
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
                                        boolExVisibl = True
                                        StartRows += 1
                                        Dim cell1 As Excel.Range 'первая ячейка в строке
                                        Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                                        Dim CellRng As Excel.Range 'массив ячеек
                                        xlSheets = xlBook.ActiveSheet
                                        If StartRows = 2 Then
                                            xlSheets.Columns.Range("A:A").ColumnWidth = 15
                                            xlSheets.Columns.Range("B:B").ColumnWidth = 15
                                            xlSheets.Columns.Range("C:C").ColumnWidth = 15
                                            xlSheets.Columns.Range("D:D").ColumnWidth = 15
                                            xlSheets.Columns.Range("E:E").ColumnWidth = 15
                                        End If
                                        Dim flagArea As Boolean = False
                                        cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Условный номер (обозначение): " & FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        '===========================================================================
                                        cell1 = xlSheets.Cells(StartRows + 1, 1) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 2, 1) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
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
                                        cell1 = xlSheets.Cells(StartRows + 1, 2) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 1, 3) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Координаты МСК, м"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 1, 4) 'ссылка на первую ячейку
                                        cell2 = xlSheets.Cells(StartRows + 1, 5) 'ссылка на 6 ячейку в строке
                                        CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                        CellRng.Merge(Type.Missing)
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Координаты WGS 84, град."
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 2) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "X"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 3) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "Y"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 4) 'ссылка на первую ячейку
                                        cell1.Font.Name = "Times New Roman"
                                        cell1.Font.Bold = 1
                                        cell1.Font.Size = 12
                                        cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                        cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                        cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                        cell1.Value = "B"
                                        '==============================================================================
                                        cell1 = xlSheets.Cells(StartRows + 2, 5) 'ссылка на первую ячейку
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
                                        StartRows = StartRows + 3
                                        If IsArray(ArrayKoordinate) = True Then
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    Continue For
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    'островок1
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = ""
                                                    StartRows += 1
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    'Контур
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Merge(Type.Missing)
                                                    cell1.Font.Name = "Times New Roman"
                                                    cell1.Font.Size = 12
                                                    cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    cell1.Value = "Контур №" & nt
                                                    StartRows += 1
                                                Else
                                                    'Координаты
                                                    cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                    cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                    CellRng.Font.Name = "Times New Roman"
                                                    CellRng.Font.Size = 12
                                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                    xlSheets.Cells(StartRows, 1).Value = nt
                                                    Dim X As String = ArrayKoordinate(1, j1)
                                                    If IsNothing(X) = False Then
                                                        Dim xi As Double = Val(X)
                                                        'xi = Math.Round(xi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 3)
                                                        cell1.NumberFormat = userNumbFormat
                                                        cell1.Value = xi 'MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    End If
                                                    Dim Y As String = ArrayKoordinate(2, j1)
                                                    If IsNothing(Y) = False Then
                                                        Dim yi As Double = Val(Y)
                                                        'yi = Math.Round(yi, 2)
                                                        cell1 = xlSheets.Cells(StartRows, 2)
                                                        cell1.NumberFormat = userNumbFormat
                                                        cell1.Value = yi 'MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    End If
                                                    Dim UserCoord_to_PseudoMercatorRez As MgCoordinate = UserCoord_to_PseudoMarcator.Transform(Val(X), Val(Y))
                                                    Dim L As Double = UserCoord_to_PseudoMercatorRez.X
                                                    Dim B As Double = UserCoord_to_PseudoMercatorRez.Y
                                                    cell1 = xlSheets.Cells(StartRows, 4)
                                                    cell1.NumberFormat = userNumbFormat
                                                    cell1.Value = B 'MathFunction.FuncFormatNumb(xi, RoundZn)
                                                    cell1 = xlSheets.Cells(StartRows, 5)
                                                    cell1.NumberFormat = userNumbFormat
                                                    cell1.Value = L 'MathFunction.FuncFormatNumb(yi, RoundZn)
                                                    StartRows += 1
                                                End If
                                            Next j1
                                            If FormReport.CheckBox7.Checked = False Then
                                                cell1 = xlSheets.Cells(StartRows, 1) 'ссылка на первую ячейку
                                                cell2 = xlSheets.Cells(StartRows, 5) 'ссылка на 3 ячейку в строке
                                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                                CellRng.Font.Name = "Times New Roman"
                                                CellRng.Font.Size = 12
                                                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                                xlSheets.Cells(StartRows, 1).Value = "Площадь: "
                                                Dim rd As Integer = FormReport.NumericUpDown2.Value
                                                cell1 = xlSheets.Cells(StartRows, 4)
                                                cell1.NumberFormat = AreaRoundZn
                                                xlSheets.Cells(StartRows, 2).Value = TotalArea
                                                xlSheets.Cells(StartRows, 3).Value = "кв.м"
                                                StartRows += 1
                                            End If
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Экспорт в MapInfo (mid, mif)
                                ElseIf FormReport.RadioButton4.Checked = True Then
                                    'все участки в один файл mid, mif
                                    If FormReport.ComboBox1.Text Like "Все участки в один файл" Then
                                        '========================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        If j = 0 Then
                                            strFolder = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку")
                                        End If
                                        If IsArray(ArrayKoordinate) = True Then
                                            Try
                                                If j = 0 Then
                                                    'открываем 2 текстовые файла на запись
                                                    Dim ParcelFile1 As String = strFolder & "\" & NamePSTable & ".mif"
                                                    Dim ParcelFile2 As String = strFolder & "\" & NamePSTable & ".mid"
                                                    ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.UTF8)
                                                    ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.UTF8)
                                                    'заполняем заголовок таблицы Участки.mif
                                                    ParcelMif1.WriteLine("Version 300")
                                                    ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                    ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                    'подсчитываем число полей
                                                    Dim numbField As Integer = 0
                                                    Dim boolArea As Boolean = False
                                                    Dim boolLenght As Boolean = False
                                                    Dim posArea As Integer = FormReport.NumericUpDown4.Value
                                                    Dim posLenght As Integer = FormReport.NumericUpDown5.Value
                                                    If IsArray(ArrayDataPSFields) = True Then
                                                        For k1 As Integer = 0 To ArrayDataPSFields.GetUpperBound(1)
                                                            Dim NameField As String = ArrayDataPSFields(1, k1)
                                                            If NameField Like "GLOBALID" Then
                                                                Continue For
                                                            ElseIf NameField Like "LOCALID" Then
                                                                Continue For
                                                            ElseIf NameField Like "LABEL" Then
                                                                Continue For
                                                                'If FormReport.CheckBox5.Checked = True Then

                                                                'End If
                                                            ElseIf NameField Like "Area" Or NameField Like "AREA" Then
                                                                boolArea = True
                                                            ElseIf NameField Like "Lenght" Or NameField Like "LENGHT" Then
                                                                boolLenght = True
                                                            End If
                                                            numbField += 1
                                                        Next k1
                                                    End If
                                                    If FormReport.CheckBox3.Checked = True Then
                                                        If boolArea = False Then
                                                            numbField += 1
                                                        End If
                                                    End If
                                                    If FormReport.CheckBox4.Checked = True Then
                                                        If boolLenght = False Then
                                                            numbField += 1
                                                        End If
                                                    End If
                                                    'записываем количество полей
                                                    ParcelMif1.WriteLine("Columns " & numbField)
                                                    TableMif = NamePSTable
                                                    'записываем поля
                                                    If IsArray(ArrayDataPSFields) = True Then
                                                        For k1 As Integer = 0 To ArrayDataPSFields.GetUpperBound(1)
                                                            If k1 = posArea - 1 Then
                                                                If boolArea = False And FormReport.CheckBox3.Checked = True Then
                                                                    ParcelMif1.WriteLine("   " & "Area" & " Decimal")
                                                                    ReDim Preserve arrayRecFields(1, countFieds)
                                                                    arrayRecFields(0, countFieds) = "Area"
                                                                    arrayRecFields(1, countFieds) = 1
                                                                    countFieds += 1
                                                                End If
                                                            End If
                                                            If k1 = posLenght - 1 Then
                                                                If boolLenght = False And FormReport.CheckBox4.Checked = True Then
                                                                    ParcelMif1.WriteLine("   " & "Lenght" & " Decimal")
                                                                    ReDim Preserve arrayRecFields(1, countFieds)
                                                                    arrayRecFields(0, countFieds) = "Lenght"
                                                                    arrayRecFields(1, countFieds) = 1
                                                                    countFieds += 1
                                                                End If
                                                            End If

                                                            Dim NameField As String = ArrayDataPSFields(1, k1)
                                                            Dim TypeField As String = ArrayDataPSFields(4, k1)
                                                            If IsNumeric(NameField) = True Then
                                                                NameField = "_" & NameField
                                                            End If
                                                            If NameField Like "GLOBALID" Then
                                                                Continue For
                                                            ElseIf NameField Like "LOCALID" Then
                                                                Continue For
                                                            End If
                                                            If NameField Like "LABEL" Then
                                                                'If FormReport.CheckBox5.Checked = True Then
                                                                Continue For
                                                                'End If
                                                            End If
                                                            If TypeField Like "0" Then
                                                                ParcelMif1.WriteLine("   " & NameField & " integer")
                                                            ElseIf TypeField Like "1" Then
                                                                ParcelMif1.WriteLine("   " & NameField & " Decimal(8,2)")
                                                            Else
                                                                ParcelMif1.WriteLine("   " & NameField & " Char(256)")
                                                            End If
                                                            ReDim Preserve arrayRecFields(1, countFieds)
                                                            arrayRecFields(0, countFieds) = NameField
                                                            arrayRecFields(1, countFieds) = TypeField
                                                            countFieds += 1
                                                        Next k1
                                                    End If
                                                    ParcelMif1.WriteLine("DATA")
                                                    ParcelMif1.WriteLine(" ")
                                                End If
                                                '==================================================================================
                                                Dim area As Double = 0
                                                If IsArray(ArrayKoordinate) = True Then
                                                    Dim strPoint As String = ""
                                                    Dim region As String = ""
                                                    For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                        Dim nt As String = ArrayKoordinate(0, j1)
                                                        region = ArrayKoordinate(2, j1)
                                                        If region Like "REGION" Then
                                                            ParcelMif1.WriteLine(region & " " & CountRegions)
                                                            ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                            strPoint = ""
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            Continue For
                                                        ElseIf region Like "PLINE MULTIPLE" Then
                                                            ParcelMif1.WriteLine(region & " " & CountRegions)
                                                            ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            strPoint = ""
                                                            Continue For
                                                        ElseIf region Like "PLINE" Then
                                                            ParcelMif1.WriteLine(region & " " & ArrayKoordinate(3, j1))
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            strPoint = ""
                                                            Continue For
                                                        ElseIf region Like "POINT" Then
                                                            strPoint = "POINT "
                                                            Continue For
                                                        End If
                                                        If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                            ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                        ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                            ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                        ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                            ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                        Else
                                                            Dim x As Double = Val(ArrayKoordinate(1, j1))
                                                            Dim y As Double = Val(ArrayKoordinate(2, j1))
                                                            'Dim z As Double = Val(ArrayKoordinate(3, j1))
                                                            'If FormReport.CheckBox2.Checked = True Then
                                                            'ParcelMif1.WriteLine(strPoint & Math.Round(x, RoundZn) & " " & Math.Round(y, RoundZn) & " " & Math.Round(z, RoundZn))
                                                            'Else
                                                            ParcelMif1.WriteLine(strPoint & Math.Round(x, RoundZn) & " " & Math.Round(y, RoundZn))
                                                            'End If
                                                        End If
                                                    Next j1
                                                    If ArrayKoordinate(2, 0) Like "REGION" Then
                                                        ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                                        ParcelMif1.WriteLine("BRUSH(2,16777215)")
                                                    ElseIf ArrayKoordinate(2, 0) Like "PLINE" Then
                                                        ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                                    Else
                                                        ParcelMif1.WriteLine("SYMBOL (35,16777215,10)")
                                                    End If
                                                    'записываем mid файл
                                                    Dim StrMid As String = ""
                                                    If IsArray(ArrayDataOD) = True And IsArray(arrayRecFields) = True Then
                                                        For k1 As Integer = 0 To arrayRecFields.GetUpperBound(1)
                                                            Dim NameField As String = arrayRecFields(0, k1)
                                                            Dim TypeField As String = arrayRecFields(1, k1)
                                                            Dim boolFindField As Boolean = False
                                                            For k2 As Integer = 0 To ArrayDataOD.GetUpperBound(1)
                                                                Dim NameField2 As String = ArrayDataOD(1, k2)
                                                                If NameField Like NameField2 Then
                                                                    boolFindField = True
                                                                    Dim NameVal As String = ArrayDataOD(2, k2)
                                                                    If IsNothing(NameVal) = True Then NameVal = " "
                                                                    If NameVal Like " " Then
                                                                        If TypeField Like "0" Then
                                                                            NameVal = "0"
                                                                        ElseIf TypeField Like "1" Then
                                                                            NameVal = "0"
                                                                        End If
                                                                    End If
                                                                    If TypeField Like "2" Then
                                                                        Dim newDataVal As String = NameVal.Replace(";", "")
                                                                    End If
                                                                    If StrMid.Trim.Length = 0 Then
                                                                        StrMid = NameVal
                                                                    Else
                                                                        StrMid = StrMid & ";" & NameVal
                                                                    End If
                                                                    Exit For
                                                                End If
                                                            Next k2
                                                            If boolFindField = False Then
                                                                Dim NameVal As String = " "
                                                                If TypeField Like "0" Then
                                                                    NameVal = "0"
                                                                ElseIf TypeField Like "1" Then
                                                                    NameVal = "0"
                                                                End If
                                                                If NameField Like "Area" Then
                                                                    NameVal = area
                                                                End If
                                                                If StrMid.Trim.Length = 0 Then
                                                                    StrMid = NameVal
                                                                Else
                                                                    StrMid = StrMid & ";" & NameVal
                                                                End If
                                                            End If
                                                        Next k1
                                                        ParcelMid1.WriteLine(StrMid)
                                                        StrMid = ""
                                                    End If
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        End If

                                    ElseIf FormReport.ComboBox1.Text Like "Согласно приказа 433/пр" Then
                                        '========================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim hgObject As String = ArrayKoordinate(4, 0)
                                        Dim idObject As ObjectId = AcadFunction.FuncReturnObjectID(hgObject)
                                        Dim nameTypeLine As String = ""
                                        Dim arrayData As String(,) = Nothing
                                        If IsNothing(idObject) = False Then
                                            If idObject.IsNull = False Then
                                                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = idObject.GetObject(OpenMode.ForRead)
                                                nameTypeLine = ent.Linetype
                                                '======================================================================================================
                                                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(ent, arrayData)
                                                'If IsArray(arrayData) = True Then
                                                '    Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Type_line", 1, 2)
                                                '    If nameObject.Trim.Length > 0 Then
                                                '        nameTypeLine = nameObject
                                                '    End If
                                                'End If
                                            End If
                                        End If
                                        If j = 0 Then
                                            strFolder = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку")
                                            'открываем 2 текстовые файла на запись
                                            If nameTypeLine Like "A-Граница территории в отношении которой осуществляется подготовка проекта планировки" Then
                                                Dim ParcelFile1 As String = strFolder & "\" & "gr_dpt" & ".mif"
                                                Dim ParcelFile2 As String = strFolder & "\" & "gr_dpt" & ".mid"
                                                ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                                                ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)
                                                'заполняем заголовок таблицы Участки.mif
                                                ParcelMif1.WriteLine("Version 300")
                                                ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                ParcelMif1.WriteLine("Columns 1")
                                                ParcelMif1.WriteLine("  GLOBALID Char(50)")
                                                ParcelMif1.WriteLine("Data")
                                            ElseIf nameTypeLine Like "A-Граница зоны планируемого размещения линейных объектов" Or nameTypeLine Like "A-Граница зоны планируемого размещения линейных объектов подлежащих реконструкции в связи с изменением их местоположения" Then
                                                Dim ParcelFile1 As String = strFolder & "\" & "ConstructionZonesBorders" & ".mif"
                                                Dim ParcelFile2 As String = strFolder & "\" & "ConstructionZonesBorders" & ".mid"
                                                ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                                                ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)
                                                'заполняем заголовок таблицы Участки.mif
                                                ParcelMif1.WriteLine("Version 300")
                                                ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                ParcelMif1.WriteLine("Columns 5")
                                                ParcelMif1.WriteLine("  GLOBALID Char(50)")
                                                ParcelMif1.WriteLine("  Name Char(256)")
                                                ParcelMif1.WriteLine("  Area Decimal(8,0)")
                                                ParcelMif1.WriteLine("  BuiltUpArea Decimal(15,2)")
                                                ParcelMif1.WriteLine("  ResidentsNum Integer")
                                                ParcelMif1.WriteLine("Data")
                                            ElseIf nameTypeLine Like "A-Красные линии действующие" Or nameTypeLine Like "A-Красные линии устанавливаемые" Or nameTypeLine Like "A-Красные линии отменяемые" Then
                                                Dim ParcelFile1 As String = strFolder & "\" & "RedLine" & ".mif"
                                                Dim ParcelFile2 As String = strFolder & "\" & "RedLine" & ".mid"
                                                ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                                                ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)
                                                'заполняем заголовок таблицы Участки.mif
                                                ParcelMif1.WriteLine("Version 300")
                                                ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                ParcelMif1.WriteLine("Columns 3")
                                                ParcelMif1.WriteLine("  GLOBALID Char(50)")
                                                ParcelMif1.WriteLine("  Number Char(256)")
                                                ParcelMif1.WriteLine("  Status Char(256)")
                                                ParcelMif1.WriteLine("Data")
                                            ElseIf nameTypeLine Like "A-Граница образуемого земельного участка" Then
                                                Dim ParcelFile1 As String = strFolder & "\" & "FormedLand" & ".mif"
                                                Dim ParcelFile2 As String = strFolder & "\" & "FormedLand" & ".mid"
                                                ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                                                ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)
                                                'заполняем заголовок таблицы Участки.mif
                                                ParcelMif1.WriteLine("Version 300")
                                                ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                ParcelMif1.WriteLine("Columns 9")
                                                ParcelMif1.WriteLine("  GLOBALID Char(50)")
                                                ParcelMif1.WriteLine("  Class Char(256)")
                                                ParcelMif1.WriteLine("  Status Char(256)")
                                                ParcelMif1.WriteLine("  FormingType Char(256)")
                                                ParcelMif1.WriteLine("  NominalNum Char(256)")
                                                ParcelMif1.WriteLine("  Location Char(256)")
                                                ParcelMif1.WriteLine("  PermittedUseType Char(256)")
                                                ParcelMif1.WriteLine("  Area Decimal(8,0)")
                                                ParcelMif1.WriteLine("  Easement Char(256)")
                                                ParcelMif1.WriteLine("Data")
                                            ElseIf nameTypeLine Like "A-Граница образуемого публичного сервитута" Or nameTypeLine Like "A-Граница образуемого сервитута" Then
                                                Dim ParcelFile1 As String = strFolder & "\" & "Easement" & ".mif"
                                                Dim ParcelFile2 As String = strFolder & "\" & "Easement" & ".mid"
                                                ParcelMif1 = New StreamWriter(ParcelFile1, False, System.Text.Encoding.Default)
                                                ParcelMid1 = New StreamWriter(ParcelFile2, False, System.Text.Encoding.Default)
                                                'заполняем заголовок таблицы Участки.mif
                                                ParcelMif1.WriteLine("Version 300")
                                                ParcelMif1.WriteLine("Delimiter " & """" & ";" & """")
                                                ParcelMif1.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                ParcelMif1.WriteLine("Columns 7")
                                                ParcelMif1.WriteLine("  GLOBALID Char(50)")
                                                ParcelMif1.WriteLine("  EasementType Char(256)")
                                                ParcelMif1.WriteLine("  Status Char(256)")
                                                ParcelMif1.WriteLine("  Area Decimal(8,0)")
                                                ParcelMif1.WriteLine("  PublicEasementType Char(256)")
                                                ParcelMif1.WriteLine("  Purpose Char(256)")
                                                ParcelMif1.WriteLine("  EGRNNum Char(256)")
                                                ParcelMif1.WriteLine("Data")
                                            End If
                                            ParcelMif1.WriteLine(" ")
                                        End If
                                        '==================================================================================
                                        Dim area As Double = 0
                                        If IsArray(ArrayKoordinate) = True Then
                                            Dim strPoint As String = ""
                                            Dim region As String = ""
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim nt As String = ArrayKoordinate(0, j1)
                                                region = ArrayKoordinate(2, j1)
                                                If region Like "REGION" Then
                                                    ParcelMif1.WriteLine(region & " " & CountRegions)
                                                    ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                    strPoint = ""
                                                    area = Val(ArrayKoordinate(5, j1))
                                                    Continue For
                                                ElseIf region Like "PLINE MULTIPLE" Then
                                                    ParcelMif1.WriteLine(region & " " & CountRegions)
                                                    ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                    area = Val(ArrayKoordinate(5, j1))
                                                    strPoint = ""
                                                    Continue For
                                                ElseIf region Like "PLINE" Then
                                                    ParcelMif1.WriteLine(region & " " & ArrayKoordinate(3, j1))
                                                    area = Val(ArrayKoordinate(5, j1))
                                                    strPoint = ""
                                                    Continue For
                                                ElseIf region Like "POINT" Then
                                                    strPoint = "POINT "
                                                    Continue For
                                                End If
                                                If IsNothing(ArrayKoordinate(1, j1)) = True And nt Like "0" Then
                                                    ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) > 0 Then
                                                    ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                    ParcelMif1.WriteLine(" " & ArrayKoordinate(3, j1))
                                                Else
                                                    Dim x As Double = Val(ArrayKoordinate(1, j1))
                                                    Dim y As Double = Val(ArrayKoordinate(2, j1))
                                                    'Dim z As Double = Val(ArrayKoordinate(3, j1))
                                                    'If FormReport.CheckBox2.Checked = True Then
                                                    'ParcelMif1.WriteLine(strPoint & Math.Round(x, RoundZn) & " " & Math.Round(y, RoundZn) & " " & Math.Round(z, RoundZn))
                                                    'Else
                                                    ParcelMif1.WriteLine(strPoint & Math.Round(x, RoundZn) & " " & Math.Round(y, RoundZn))
                                                    'End If
                                                End If
                                            Next j1
                                            If ArrayKoordinate(2, 0) Like "REGION" Then
                                                ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                                ParcelMif1.WriteLine("BRUSH(2,16777215)")
                                            ElseIf ArrayKoordinate(2, 0) Like "PLINE" Then
                                                ParcelMif1.WriteLine("Pen(12, 2, 16711680)")
                                            Else
                                                ParcelMif1.WriteLine("SYMBOL (35,16777215,10)")
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'записываем mid файл
                                            Dim objectGUID As Guid = Guid.NewGuid()
                                            Dim StrMid As String = ""
                                            If nameTypeLine Like "A-Граница территории в отношении которой осуществляется подготовка проекта планировки" Then
                                                StrMid = objectGUID.ToString
                                            ElseIf nameTypeLine Like "A-Граница зоны планируемого размещения линейных объектов" Or nameTypeLine Like "A-Граница зоны планируемого размещения линейных объектов подлежащих реконструкции в связи с изменением их местоположения" Then
                                                Dim name As String = MathFunction.FuncFindValueToArray2d(arrayData, "Type_line", 1, 2)
                                                StrMid = objectGUID.ToString & ";" & name & ";" & TotalArea & ";;"
                                            ElseIf nameTypeLine Like "A-Красные линии действующие" Or nameTypeLine Like "A-Красные линии устанавливаемые" Or nameTypeLine Like "A-Красные линии отменяемые" Then
                                                Dim numbRedLine As String = MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2)
                                                Dim status As String = ""
                                                If nameTypeLine Like "A-Красные линии действующие" Then
                                                    status = "7Е.1"
                                                ElseIf nameTypeLine Like "A-Красные линии устанавливаемые" Then
                                                    status = "7Е.2"
                                                ElseIf nameTypeLine Like "A-Красные линии отменяемые" Then
                                                    status = "7Е.3"
                                                End If
                                                StrMid = objectGUID.ToString & ";" & numbRedLine & ";" & status
                                            ElseIf nameTypeLine Like "A-Граница образуемого земельного участка" Then
                                                Dim conditionalNumber As String = MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2)
                                                Dim Cad_number As String = MathFunction.FuncFindValueToArray2d(arrayData, "CadlNumber", 1, 2)
                                                Dim adress As String = MathFunction.FuncFindValueToArray2d(arrayData, "Adress", 1, 2)
                                                Dim category As String = MathFunction.FuncFindValueToArray2d(arrayData, "Category", 1, 2)
                                                Dim utilization As String = MathFunction.FuncFindValueToArray2d(arrayData, "Utilization", 1, 2)
                                                Dim metodCreate As String = MathFunction.FuncFindValueToArray2d(arrayData, "Method_create", 1, 2)
                                                Dim sposobObr As String = MathFunction.FuncFindValueToArray2d(arrayData, "Confiscation", 1, 2)
                                                If metodCreate.IndexOf("осударствен") > -1 Then
                                                    metodCreate = "7G.1"
                                                ElseIf metodCreate.IndexOf("раздел") > -1 Then
                                                    metodCreate = "7G.2"
                                                ElseIf metodCreate.IndexOf("объединен") > -1 Then
                                                    metodCreate = "7G.3"
                                                ElseIf metodCreate.IndexOf("Выдел") > -1 Or metodCreate.IndexOf("выдел") > -1 Then
                                                    metodCreate = "7G.4"
                                                ElseIf metodCreate.IndexOf("ерераспределен") > -1 Then
                                                    metodCreate = "7G.5"
                                                Else
                                                    metodCreate = ""
                                                End If

                                                If sposobObr.IndexOf("зъят") > -1 Then
                                                    If metodCreate Like "7G.1" Then
                                                        sposobObr = "7F.2"
                                                    Else
                                                        If category.IndexOf("насел") > -1 Then
                                                            sposobObr = "7F.6"
                                                        Else
                                                            sposobObr = "7F.7"
                                                        End If
                                                    End If
                                                ElseIf sposobObr.IndexOf("езервиров") > -1 Then
                                                    If category.IndexOf("насел") > -1 Then
                                                        sposobObr = "7F.6"
                                                    Else
                                                        sposobObr = "7F.4"
                                                    End If
                                                ElseIf sposobObr.IndexOf("Измен") > -1 Or sposobObr.IndexOf("измен") > -1 Then
                                                    sposobObr = "7F.3"
                                                ElseIf sposobObr.Trim.Length = 0 Then
                                                    sposobObr = "7F.1"
                                                End If
                                                StrMid = objectGUID.ToString & ";" & sposobObr & ";7В.2;" & metodCreate & ";" & conditionalNumber & ";" & adress & ";" & utilization & ";" & TotalArea & ";"
                                            ElseIf nameTypeLine Like "A-Граница образуемого публичного сервитута" Or nameTypeLine Like "A-Граница образуемого сервитута" Then
                                                Dim purpose As String = MathFunction.FuncFindValueToArray2d(arrayData, "Purpose", 1, 2)
                                                Dim typeServ As String = ""
                                                If nameTypeLine Like "A-Граница образуемого сервитута" Then
                                                    typeServ = "13A.2"
                                                Else
                                                    typeServ = "13A.1"
                                                End If
                                                StrMid = objectGUID.ToString & ";" & typeServ & ";13B.1;" & TotalArea & ";" & "13C.3" & ";" & purpose & ";"
                                            End If
                                            Dim newString As String = StrMid.Replace(vbCrLf, "").Replace(vbLf, "")
                                            ParcelMid1.WriteLine(newString)
                                            StrMid = ""
                                        End If
                                    Else
                                        'каждый участок в свой файл
                                        '========================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        If j = 0 Then
                                            NameDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку")
                                        End If
                                        If IsArray(ArrayDataOD) = True Then
                                            For k1 As Integer = 0 To ArrayDataOD.GetUpperBound(1)
                                                Dim NameField As String = ArrayDataOD(1, k1)
                                                If NameField Like "Номер_сервитута_пуб_серв" Then
                                                    ConditionCounterParcel = ArrayDataOD(2, k1)
                                                    Exit For
                                                End If
                                            Next k1
                                        End If
                                        Dim NameDoc As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                        NameDoc = NameDoc.Replace("\", "_")
                                        NameDoc = NameDoc.Replace("/", "_")
                                        NameDoc = NameDoc.Replace(":", "_")
                                        NameDoc = NameDoc.Replace("*", "_")
                                        NameDoc = NameDoc.Replace(".", "_")
                                        NameDoc = NameDoc.Replace(",", "_")
                                        If IsArray(ArrayKoordinate) = True Then
                                            'открываем 2 текстовые файла на запись
                                            Dim ParcelFile1 As String = NameDir & "\" & NameDoc & ".mif"
                                            Dim ParcelFile2 As String = NameDir & "\" & NameDoc & ".mid"
                                            '=========================================================================================
                                            Dim ArrayFiles As String() = IO.Directory.GetFiles(NameDir)
                                            If IsArray(ArrayFiles) = True Then
                                                For k1 As Integer = 0 To ArrayFiles.Length - 1
                                                    Dim ffile As String = ArrayFiles(k1)
                                                    If ffile Like ParcelFile1 Then
                                                        Dim rng As Random = New Random
                                                        NameDoc = NameDoc & rng.Next
                                                        ParcelFile1 = NameDir & "\" & NameDoc & ".mif"
                                                        ParcelFile2 = NameDir & "\" & NameDoc & ".mid"
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                            Dim ParcelMif As StreamWriter = New StreamWriter(ParcelFile1, False, System.Text.Encoding.UTF8)
                                            Dim ParcelMid As StreamWriter = New StreamWriter(ParcelFile2, False, System.Text.Encoding.UTF8)
                                            'заполняем заголовок таблицы Участки.mif
                                            Try
                                                ParcelMif.WriteLine("Version 300")
                                                ParcelMif.WriteLine("Delimiter " & """" & "," & """")
                                                ParcelMif.WriteLine("CoordSys NonEarth Units " & """" & "m" & """" & " Bounds(" & Math.Round(MinCoordX, 0) - 10 & ", " & Math.Round(MinCoordY, 0) - 10 & ")(" & Math.Round(MaxCoordX, 0) + 10 & ", " & Math.Round(MaxCoordY, 0) + 10 & ")")
                                                'подсчитываем число полей
                                                Dim numbField As Integer = 0
                                                Dim boolArea As Boolean = False
                                                Dim boolLenght As Boolean = False
                                                Dim posArea As Integer = FormReport.NumericUpDown4.Value
                                                Dim posLenght As Integer = FormReport.NumericUpDown5.Value
                                                If IsArray(ArrayDataPSFields) = True Then
                                                    For k1 As Integer = 0 To ArrayDataPSFields.GetUpperBound(1)
                                                        Dim NameField As String = ArrayDataPSFields(1, k1)
                                                        If NameField Like "GLOBALID" Then
                                                            Continue For
                                                        ElseIf NameField Like "LOCALID" Then
                                                            Continue For
                                                        ElseIf NameField Like "LABEL" Then
                                                            Continue For
                                                        ElseIf NameField Like "Area" Or NameField Like "AREA" Then
                                                            boolArea = True
                                                        ElseIf NameField Like "Lenght" Or NameField Like "LENGHT" Then
                                                            boolLenght = True
                                                        End If
                                                        numbField += 1
                                                    Next k1
                                                End If
                                                If FormReport.CheckBox3.Checked = True Then
                                                    If boolArea = False Then
                                                        numbField += 1
                                                    End If
                                                End If
                                                If FormReport.CheckBox4.Checked = True Then
                                                    If boolLenght = False Then
                                                        numbField += 1
                                                    End If
                                                End If

                                                'ParcelMif.WriteLine("Columns " & numbField)
                                                'If IsArray(ArrayDataOD) = True Then
                                                '    For k1 As Integer = 0 To ArrayDataOD.GetUpperBound(1)
                                                '        Dim NameField As String = ArrayDataOD(1, k1)
                                                '        If NameField Like "GLOBALID" Then
                                                '            Continue For
                                                '        ElseIf NameField Like "LOCALID" Then
                                                '            Continue For
                                                '        ElseIf NameField Like "TypeOwner" Then
                                                '            Continue For
                                                '        ElseIf NameField Like "Area" Then
                                                '            'ParcelMif.WriteLine("   " & NameField & " Decimal(10, " & AreaRoundZn & ")")
                                                '            'Continue For
                                                '        End If
                                                '        ParcelMif.WriteLine("   " & NameField & " Char(256)")
                                                '    Next k1
                                                'End If
                                                'ParcelMif.WriteLine("DATA")
                                                'ParcelMif.WriteLine(" ")
                                                '==================================================================================
                                                'записываем регион
                                                ParcelMif.WriteLine("Columns " & numbField)
                                                TableMif = NamePSTable
                                                'записываем поля
                                                If IsArray(ArrayDataPSFields) = True Then
                                                    For k1 As Integer = 0 To ArrayDataPSFields.GetUpperBound(1)
                                                        If k1 = posArea - 1 Then
                                                            If boolArea = False And FormReport.CheckBox3.Checked = True Then
                                                                ParcelMif.WriteLine("   " & "Area" & " Decimal")
                                                                ReDim Preserve arrayRecFields(1, countFieds)
                                                                arrayRecFields(0, countFieds) = "Area"
                                                                arrayRecFields(1, countFieds) = 1
                                                                countFieds += 1
                                                            End If
                                                        End If
                                                        If k1 = posLenght - 1 Then
                                                            If boolLenght = False And FormReport.CheckBox4.Checked = True Then
                                                                ParcelMif.WriteLine("   " & "Lenght" & " Decimal")
                                                                ReDim Preserve arrayRecFields(1, countFieds)
                                                                arrayRecFields(0, countFieds) = "Lenght"
                                                                arrayRecFields(1, countFieds) = 1
                                                                countFieds += 1
                                                            End If
                                                        End If

                                                        Dim NameField As String = ArrayDataPSFields(1, k1)
                                                        Dim TypeField As String = ArrayDataPSFields(4, k1)
                                                        If IsNumeric(NameField) = True Then
                                                            NameField = "_" & NameField
                                                        End If
                                                        If NameField Like "GLOBALID" Then
                                                            Continue For
                                                        ElseIf NameField Like "LOCALID" Then
                                                            Continue For
                                                        End If
                                                        If NameField Like "LABEL" Then
                                                            Continue For
                                                        End If
                                                        If TypeField Like "0" Then
                                                            ParcelMif.WriteLine("   " & NameField & " integer")
                                                        ElseIf TypeField Like "1" Then
                                                            ParcelMif.WriteLine("   " & NameField & " Decimal(8,2)")
                                                        Else
                                                            ParcelMif.WriteLine("   " & NameField & " Char(256)")
                                                        End If
                                                        ReDim Preserve arrayRecFields(1, countFieds)
                                                        arrayRecFields(0, countFieds) = NameField
                                                        arrayRecFields(1, countFieds) = TypeField
                                                        countFieds += 1
                                                    Next k1
                                                End If
                                                ParcelMif.WriteLine("DATA")
                                                ParcelMif.WriteLine(" ")
                                                '==================================================================================
                                                Dim area As Double = 0
                                                '==================================================================================
                                                If IsArray(ArrayKoordinate) = True Then
                                                    Dim strPoint As String = ""
                                                    Dim region As String = ""
                                                    For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                        Dim nt As String = ArrayKoordinate(0, j1)
                                                        region = ArrayKoordinate(2, j1)
                                                        If region Like "REGION" Then
                                                            ParcelMif.WriteLine(region & " " & CountRegions)
                                                            ParcelMif.WriteLine(" " & ArrayKoordinate(3, j1))
                                                            strPoint = ""
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            Continue For
                                                        ElseIf region Like "PLINE MULTIPLE" Then
                                                            ParcelMif.WriteLine(region & " " & CountRegions)
                                                            ParcelMif.WriteLine(" " & ArrayKoordinate(3, j1))
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            strPoint = ""
                                                            Continue For
                                                        ElseIf region Like "PLINE" Then
                                                            ParcelMif.WriteLine(region & " " & ArrayKoordinate(3, j1))
                                                            area = Val(ArrayKoordinate(5, j1))
                                                            strPoint = ""
                                                            Continue For
                                                        ElseIf region Like "POINT" Then
                                                            strPoint = "POINT "
                                                            Continue For
                                                        End If
                                                        If IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And Val(nt) >= 0 Then
                                                            ParcelMif.WriteLine(" " & ArrayKoordinate(3, j1))
                                                        ElseIf IsNothing(ArrayKoordinate(1, j1)) = True And IsNothing(ArrayKoordinate(2, j1)) = True And nt Like "-1" Then
                                                            ParcelMif.WriteLine(" " & ArrayKoordinate(3, j1))
                                                        Else
                                                            Dim x As Double = Val(ArrayKoordinate(1, j1))
                                                            Dim y As Double = Val(ArrayKoordinate(2, j1))
                                                            ParcelMif.WriteLine(Math.Round(x, RoundZn) & " " & Math.Round(y, RoundZn))
                                                        End If
                                                    Next j1
                                                    ParcelMif.WriteLine("Pen(12, 2, 16711680)")
                                                    ParcelMif.WriteLine("BRUSH(2,16777215)")
                                                    Dim StrMid As String = ""
                                                    If IsArray(arrayRecFields) = True Then
                                                        If IsArray(ArrayDataOD) = True Then
                                                            For k1 As Integer = 0 To arrayRecFields.GetUpperBound(1)
                                                                Dim NameField As String = arrayRecFields(0, k1)
                                                                For k2 As Integer = 0 To ArrayDataOD.GetUpperBound(1)
                                                                    Dim destNameField As String = ArrayDataOD(1, k2)
                                                                    Dim NameVal As String = ArrayDataOD(2, k2)
                                                                    If IsNothing(NameVal) = True Then NameVal = ""
                                                                    If destNameField Like NameField Then
                                                                        If StrMid.Trim.Length = 0 Then
                                                                            StrMid = Chr(34) & NameVal & Chr(34)
                                                                        Else
                                                                            StrMid = StrMid & "," & Chr(34) & NameVal & Chr(34)
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            Next k1
                                                        End If
                                                    Else
                                                        StrMid = ""
                                                    End If
                                                    ParcelMid.WriteLine(StrMid)
                                                    ParcelMid.Close()
                                                    ParcelMif.Close()
                                                End If
                                            Catch ex As Exception
                                            End Try
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Экспорт в csv
                                ElseIf FormReport.RadioButton5.Checked = True Then
                                    '========================================================================
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim typeCoord As String = FormReport.ComboBox1.Text
                                    Dim NameDoc As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionCounterParcel)
                                    NameDoc = NameDoc.Replace("\", "_")
                                    NameDoc = NameDoc.Replace("/", "_")
                                    NameDoc = NameDoc.Replace(":", "_")
                                    NameDoc = NameDoc.Replace("*", "_")
                                    NameDoc = NameDoc.Replace(".", "_")
                                    NameDoc = NameDoc.Replace(",", "_")
                                    If j = 0 Then
                                        strFolder = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку")
                                    End If
                                    If IsArray(ArrayKoordinate) = True Then
                                        Try
                                            'открываем 2 текстовые файла на запись
                                            Dim ParcelFile As String = strFolder & "\" & NameDoc & ".csv"
                                            ParcelMif1 = New StreamWriter(ParcelFile, False, System.Text.Encoding.Default)
                                            ParcelMif1.WriteLine("Контур;Префикс номера;Номер;Старый X;Старый Y;Новый X;Новый Y;Метод определения;Формула;Радиус;Погрешность;Описание закрепления;Глубина;Высота;Кадастровый номер")
                                            Dim numbCounter As Integer = 1
                                            Dim numOstrov As Integer = 0
                                            Dim numPoint As Integer = 1
                                            Dim firstNumber As Integer = -1
                                            Dim firstPreficsif As String = ""
                                            Dim firstX As String = ""
                                            Dim firstY As String = ""
                                            Dim deltaGeopoint As String = "0,10"
                                            For j1 As Integer = 0 To ArrayKoordinate.GetUpperBound(1)
                                                Dim n1 As String = ArrayKoordinate(0, j1)
                                                If IsNothing(n1) = True Then
                                                    n1 = ""
                                                End If
                                                Dim n2 As String = ArrayKoordinate(1, j1)
                                                If IsNothing(n2) = True Then
                                                    n2 = ""
                                                End If
                                                Dim n3 As String = ArrayKoordinate(2, j1)
                                                If IsNothing(n3) = True Then
                                                    n3 = ""
                                                End If
                                                Dim n4 As String = ArrayKoordinate(4, j1)
                                                If IsNothing(n4) = True Then
                                                    n4 = ""
                                                End If
                                                n4 = n4.Replace(".", ",")
                                                Dim n5 As String = ArrayKoordinate(5, j1)
                                                If IsNothing(n5) = True Then
                                                    n5 = ""
                                                End If
                                                Dim n6 As String = ArrayKoordinate(6, j1)
                                                If IsNothing(n6) = True Then
                                                    n6 = ""
                                                End If
                                                If n1.Trim.Length = 0 And n2.Trim.Length = 0 Then
                                                    Continue For
                                                ElseIf n1.Trim.Length > 0 And n2.Trim.Length = 0 Then
                                                    If Val(n1) > -1 Then
                                                        ParcelMif1.WriteLine(";;;;;;;;;;;;;;")
                                                        numbCounter = Val(n1)
                                                        deltaGeopoint = n6
                                                        deltaGeopoint = deltaGeopoint.Replace(".", ",")
                                                        numOstrov = 0
                                                        firstNumber = -1
                                                        firstPreficsif = ""
                                                        firstX = ""
                                                        firstY = ""
                                                        Continue For
                                                    Else
                                                        numOstrov += 1
                                                        ParcelMif1.WriteLine(";;;;;;;;;;;;;;")
                                                        firstNumber = -1
                                                        firstPreficsif = ""
                                                        firstX = ""
                                                        firstY = ""
                                                        Continue For
                                                    End If
                                                Else
                                                    Dim prefNumber As String = ""
                                                    If IsNumeric(n1) = False Then
                                                        prefNumber = Left(n1, 1)
                                                        n1 = Mid(n1, 2)
                                                    End If
                                                    If Val(n1) = 0 Then
                                                        n1 = numPoint
                                                        numPoint += 1
                                                    End If
                                                    If firstNumber = -1 Then
                                                        firstPreficsif = prefNumber
                                                        firstNumber = n1
                                                        firstX = ArrayKoordinate(1, j1)
                                                        firstY = ArrayKoordinate(2, j1)
                                                    Else
                                                        If j1 = ArrayKoordinate.GetUpperBound(1) Then
                                                            If ArrayKoordinate(1, j1) Like firstX And ArrayKoordinate(2, j1) Like firstY Then
                                                                prefNumber = firstPreficsif
                                                                n1 = firstNumber
                                                            End If
                                                        End If
                                                    End If
                                                    Dim xish As Double = Val(n2)
                                                    Dim yish As Double = Val(n3)
                                                    xish = Math.Round(xish, RoundZn)
                                                    yish = Math.Round(yish, RoundZn)
                                                    Dim x As String = CStr(xish)
                                                    x = x.Replace(".", ",")
                                                    Dim y As String = CStr(yish)
                                                    y = y.Replace(".", ",")
                                                    If n4.Trim.Length = 0 Then
                                                        n4 = deltaGeopoint
                                                    End If
                                                    If n5.Trim.Length = 0 Then
                                                        n5 = "626003000000"
                                                    End If
                                                    If n6.Trim.Length = 0 Then
                                                        n6 = "692005000000"
                                                    End If
                                                    If typeCoord Like "Уточняемые координаты" Then
                                                        If numOstrov = 0 Then
                                                            ParcelMif1.WriteLine("[" & numbCounter & "];" & prefNumber & ";" & n1 & ";;;" & y & ";" & x & ";" & n6 & ";;;" & n4 & ";" & n5)
                                                        Else
                                                            ParcelMif1.WriteLine("[" & numbCounter & "." & numOstrov & "];" & prefNumber & ";" & n1 & ";;;" & y & ";" & x & ";" & n6 & ";;;" & n4 & ";" & n5)
                                                        End If
                                                    ElseIf typeCoord Like "Исходные координаты" Then
                                                        If numOstrov = 0 Then
                                                            ParcelMif1.WriteLine("[" & numbCounter & "];" & prefNumber & ";" & n1 & ";" & y & ";" & x & ";;;" & n6 & ";;;" & n4 & ";" & n5)
                                                        Else
                                                            ParcelMif1.WriteLine("[" & numbCounter & "." & numOstrov & "];" & prefNumber & ";" & n1 & ";" & y & ";" & x & ";;;" & n6 & ";;;" & n4 & ";" & n5)
                                                        End If
                                                    End If
                                                End If
                                            Next j1
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    ParcelMif1.Close()
                                End If
                            Next j
                            '=================================================================================================================
                            'запуск Word или Excel
                            If FormReport.RadioButton3.Checked = True Then
                                WApp.Visible = True
                                WDoc = Nothing
                                WApp.Quit()
                                WApp = Nothing
                            ElseIf FormReport.RadioButton2.Checked = True Then
                                If boolExVisibl = True Then
                                    xlApp.Visible = True
                                    xlBook = Nothing
                                    xlApp.Quit()
                                    xlApp = Nothing
                                Else
                                    xlApp.Quit()
                                    xlApp = Nothing
                                End If
                            ElseIf FormReport.RadioButton4.Checked = True Then
                                ParcelMid1.Close()
                                ParcelMif1.Close()
                            ElseIf FormReport.RadioButton5.Checked = True Then
                                ParcelMif1.Close()
                            End If
                        End If

                    End If
                    tr.Commit()
                End Using
            Catch ex As System.Exception
            End Try
        End Using
    End Sub
End Class
