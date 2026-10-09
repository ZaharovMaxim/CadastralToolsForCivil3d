Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Drawing
Imports System.IO
Imports Exception = System.Exception
Imports OSGeo.MapGuide
Imports System.Net
Imports System.Windows.Forms

Public Class FormCreateParcels
    Inherits System.Windows.Forms.Form
    Public PutchTemlateDwg As String = ""
    Public PutchTemlateXml As String = ""
    Public CategoryObject As String = "Проект планировки и межевания территории"
    Public dictionaryDataParcels As New Dictionary(Of Integer, String(,)) 'участки
    Public Xmin As Single = 1000000000
    Public Ymin As Single = 1000000000
    Public Xmax As Single = -1000000000
    Public Ymax As Single = -1000000000
    Public Sub New()
        InitializeComponent()
        'начальный путь к файлу-шаблону
        Dim nameFileXML As String = "Стандарт_ADSK.xml"
        Dim nameFileDwg As String = "Стандарт_ADSK.dwg"
        Dim boolFindTeplates As Boolean = FuncFiles.FuncFindTemplateFiles(nameFileXML, nameFileDwg, PutchTemlateXml, PutchTemlateDwg)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. Читаем таблицу с категориями объектов
        Dim arrayTypeObject As String() = Nothing
        Dim boolFindTypeObject As Boolean = TOPOSymbols.FuncReadTablesToXML(PutchTemlateXml, CategoryObject, arrayTypeObject)
        If IsArray(arrayTypeObject) = True Then
            ComboBox8.DataSource = arrayTypeObject
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2. Читаем таблицу с объектов
        Dim tableName As String = ComboBox8.Text
        Dim arrayObject As String() = Nothing
        Dim boolFindObject As Boolean = TOPOSymbols.FuncReadTypeObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Полигон", arrayObject)
        If IsArray(arrayObject) = True Then
            ComboBox12.DataSource = arrayObject
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '3. Читаем характеристики объекта
        Dim deskObject As String = ComboBox12.Text
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolFindDataObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Полигон", deskObject, arrayPropertiesObject)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '4.читаем таблицу PS
        Dim userPSTable As String = "OBJECT"
        If IsArray(arrayPropertiesObject) = True Then
            For i As Integer = 0 To arrayPropertiesObject.GetUpperBound(1)
                If arrayPropertiesObject(0, i) Like "PS" Then
                    userPSTable = arrayPropertiesObject(1, i).Trim
                    Exit For
                End If
            Next i
        End If
        Dim arrayDataField As New Dictionary(Of Integer, String(,)) 'линейные
        Dim NameTablePS As String = TOPOSymbols.FuncReadDataTablePSToXML(PutchTemlateXml, userPSTable, arrayDataField)
        DataGridView2.Tag = userPSTable
        If userPSTable Like "PROJECT_PARCEL" Then
            DataGridView2.ContextMenuStrip = ContextMenuParcel
        ElseIf userPSTable Like "PROJECT_BULDING" Then
            DataGridView2.ContextMenuStrip = ContextMenuBuild
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '4. Префикс номера участка
        Dim ArrayPref As String() = {"00:00:0000000:0000:ЗУ", "Участок №", "0000:ЗУ", "0000/чзу", "Публ.сервитут №", "Сервитут №", "00:00:0000000:0000/чзу", "Часть участка №", "№", ""}
        ComboBox2.DataSource = ArrayPref
        Dim firstOboznParcels As String = ComboBox2.Text
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем датагрид
        If arrayDataField.Count > 0 Then
            DataGridView2.Rows.Clear()
            For i As Integer = 0 To arrayDataField.Count - 1
                DataGridView2.Rows.Add()
                If arrayDataField.ContainsKey(i) = True Then
                    Dim arrayField As String(,) = arrayDataField.Item(i)
                    If IsArray(arrayField) = True Then
                        For j As Integer = 0 To arrayField.GetUpperBound(1)
                            Dim nameF As String = arrayField(0, j)
                            Dim val As String = arrayField(1, j)
                            If nameF Like "Name" Then
                                DataGridView2.Rows(i).Cells(0).Tag = val
                            ElseIf nameF Like "Description" Then
                                DataGridView2.Rows(i).Cells(0).Value = val
                            ElseIf nameF Like "DefaultData" Then
                                DataGridView2.Rows(i).Cells(1).Value = val
                            End If
                            If val Like "ConditionalNumber" Then
                                DataGridView2.Rows(i).Cells(1).Value = firstOboznParcels
                            ElseIf val Like "LABEL" Then
                                Dim countrys As String() = New String() {"США", "ОАЭ", "ЮАР"}
                                Dim s = New Windows.Forms.DataGridViewComboBoxCell()
                                s.DataSource = countrys
                                DataGridView2.Rows(i).Cells(1).Value = deskObject
                            End If
                        Next j
                    End If
                End If
            Next i
            If DataGridView2.Rows(0).Cells(0).Value = Nothing Then
                DataGridView2.Rows.RemoveAt(0)
            End If
        End If
        DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'забираем все мультивыноски
        Dim arrayMleader As String() = Nothing
        Dim boolFindMLeader As Boolean = TOPOSymbols.FuncReadTypeObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Мультивыноска", arrayMleader)
        If IsArray(arrayMleader) = True Then
            ComboBox11.DataSource = arrayMleader
        Else
            RadioButton1.Checked = True
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '5. Кадастровый номер исходного участка
        'MaskedTextBox1.Mask = "00:00:#######:#####"
        'MaskedTextBox1.Text = "00:00:0000000:00000"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'масштаб
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        Dim ScaleM As Integer() = {MAnno, 20000, 15000, 12000, 10000, 8000, 7000, 5000, 2000, 1000, 500, 100, 50}
        ComboBox5.DataSource = ScaleM
        NumericUpDown2.Enabled = False
        Button1.Enabled = False
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальный номер участка
        NumericUpDown4.Value = 1
        'начальный номер контура
        NumericUpDown2.Value = 1
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'предварительное создание участков
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'ссылка на чертеж .dwg
        Dim db As Database = acDoc.Database ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim PlineArray As String(,) = Nothing
        Dim CountPLine As Integer = 0
        Dim tr As Transaction = acDoc.TransactionManager.StartTransaction()
        Using tr
            Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim acSSPrompt As PromptSelectionResult = ed.GetSelection() 'выбираем контура
                Dim acSSet1 As SelectionSet = Nothing
                If acSSPrompt.Status = PromptStatus.OK Then
                    acSSet1 = acSSPrompt.Value
                    'перебираем объекты в коллекции
                    If acSSet1.Count > 0 Then
                        For Each objId1 As ObjectId In acSSet1.GetObjectIds()
                            If Not IsDBNull(objId1) Then
                                Dim ent As Entity = tr.GetObject(objId1, OpenMode.ForRead)
                                If TypeOf ent Is Polyline Then
                                    Dim plineTemp As Polyline = ent
                                    If plineTemp.Closed = False Then
                                        If plineTemp.StartPoint.X <> plineTemp.StartPoint.Y And plineTemp.StartPoint.Y <> plineTemp.EndPoint.Y Then
                                            Continue For
                                        End If
                                    End If
                                    ReDim Preserve PlineArray(1, CountPLine)
                                    PlineArray(0, CountPLine) = 0
                                    PlineArray(1, CountPLine) = ent.Handle.ToString
                                    CountPLine += 1

                                    Dim pt3dmax As Point3d = plineTemp.GeometricExtents.MaxPoint
                                    Dim pt3dmin As Point3d = plineTemp.GeometricExtents.MinPoint
                                    If pt3dmax.X > Xmax Then
                                        Xmax = Math.Round(pt3dmax.X, 4)
                                    End If
                                    If pt3dmax.Y > Ymax Then
                                        Ymax = Math.Round(pt3dmax.Y, 4)
                                    End If
                                    If pt3dmin.X < Xmin Then
                                        Xmin = Math.Round(pt3dmin.X, 4)
                                    End If
                                    If pt3dmin.Y < Ymin Then
                                        Ymin = Math.Round(pt3dmin.Y, 4)
                                    End If
                                End If
                            End If
                        Next
                    End If
                    'если в коллекции более 2 участков то предлагаем пользователю произвести сортировку
                    If CountPLine > 2 Then
                        Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                        pIntOpts.Message = vbCrLf & "Упорядочить полигоны по возранию?"
                        pIntOpts.Keywords.Add("Да")
                        pIntOpts.Keywords.Add("Нет")
                        pIntOpts.Keywords.Default = "Да"
                        pIntOpts.AllowNone = True
                        Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                        If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                            Throw New System.Exception("Прервано пользователем")
                            Exit Sub 'если нажали ESC то выходим из программы
                        End If
                        If pIntRes.StringResult Like "Да" Then
                            Dim per As PromptEntityResult = ed.GetEntity("Выберите начальный полигон: ")
                            If per.Status = PromptStatus.OK Then
                                Dim obj As DBObject = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект
                                If TypeOf obj Is Polyline Then 'проверяем тип объекта, если это полиния
                                    Dim entPLine As Entity = obj
                                    'находим центр габаритного контейнера выбранного полигона
                                    Dim CentrObj As Point3d = MathFunction.FuncFindCenterObjectByGeometricExtents(entPLine)
                                    'проходим циклом по всем объектам и сортируем их в порядке удаления
                                    For i2 As Integer = 0 To PlineArray.GetUpperBound(1)
                                        Dim hgStr As String = PlineArray(1, i2)
                                        Dim PlineObjID As ObjectId = AcadFunction.FuncReturnObjectID(hgStr)
                                        If IsNothing(PlineObjID) = False Then
                                            If PlineObjID.IsNull = False Then
                                                Dim entTempPline As Polyline = tr.GetObject(PlineObjID, OpenMode.ForRead)
                                                'находим центр габаритного контейнера
                                                Dim centrObj2 As Point3d = MathFunction.FuncFindCenterObjectByGeometricExtents(entTempPline)
                                                'находим расстояние до полигонов
                                                Dim Dist As Double = CentrObj.GetVectorTo(centrObj2).Length
                                                Dist = Math.Round(Dist, 2)
                                                PlineArray(0, i2) = Dist
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                            If IsArray(PlineArray) = True Then
                                If PlineArray.GetUpperBound(1) > 1 Then
                                    Dim SortArrayPline As Boolean = MathFunction.FuncSortStrArray(PlineArray, 0)
                                End If
                            End If
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'заполняем массив
                    If IsArray(PlineArray) = True Then
                        Dim StartNumParcel As Integer = NumericUpDown4.Value
                        Dim StartNumKounter As Integer = NumericUpDown2.Value
                        For i As Integer = 0 To PlineArray.GetUpperBound(1)
                            Dim handleParcels As String = PlineArray(1, i)
                            Dim numberParcel As Integer = StartNumParcel
                            Dim numberCounters As Integer = StartNumKounter
                            'если участок существует
                            If dictionaryDataParcels.ContainsKey(numberParcel) = True Then
                                Dim strTempVal As String(,) = dictionaryDataParcels.Item(numberParcel)
                                strTempVal(2, 2) = strTempVal(2, 2) & " " & handleParcels 'добавляем к списку handle
                                strTempVal(2, 1) = strTempVal(2, 1) & " " & numberCounters 'добавляем к списку номер контура
                                dictionaryDataParcels.Item(numberParcel) = strTempVal
                                numberCounters += 1
                            Else 'новый участок
                                Dim arrayProjectParcels As String(,) = Nothing
                                ReDim Preserve arrayProjectParcels(2, DataGridView2.RowCount + 4)
                                arrayProjectParcels(1, 0) = "NumberParcel"
                                arrayProjectParcels(2, 0) = numberParcel
                                arrayProjectParcels(1, 1) = "NumberCounter"
                                If CheckBox7.Checked = True Then
                                    arrayProjectParcels(2, 1) = numberCounters
                                    numberCounters += 1
                                Else
                                    arrayProjectParcels(2, 1) = "0"
                                End If
                                arrayProjectParcels(1, 2) = "Handle"
                                arrayProjectParcels(2, 2) = handleParcels
                                arrayProjectParcels(1, 3) = "PS"
                                arrayProjectParcels(2, 3) = DataGridView2.Tag
                                Dim n As Integer = 4
                                For Each row As DataGridViewRow In DataGridView2.Rows
                                    Dim namefieldsFields As String = row.Cells(0).Tag
                                    Dim valueRows As String = row.Cells(1).Value
                                    If IsNothing(valueRows) = True Then valueRows = ""
                                    If IsNothing(namefieldsFields) = False Then
                                        If namefieldsFields.Trim.Length > 0 Then
                                            arrayProjectParcels(0, n) = DataGridView2.Tag
                                            arrayProjectParcels(1, n) = namefieldsFields
                                            arrayProjectParcels(2, n) = valueRows
                                            n += 1
                                        End If
                                    End If
                                Next
                                dictionaryDataParcels.Add(numberParcel, arrayProjectParcels)
                                If CheckBox7.Checked = False Then
                                    numberParcel += 1
                                End If
                            End If
                            StartNumKounter = numberCounters
                            StartNumParcel = numberParcel
                        Next i
                        NumericUpDown2.Value = StartNumKounter
                        NumericUpDown4.Value = StartNumParcel
                    End If
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Ошибка ввода данных. Макро прервано")
                Exit Sub
            End Try
            tr.Commit()
        End Using
        If dictionaryDataParcels.Count > 0 Then
            ComboBox12.Enabled = False
            ComboBox8.Enabled = False
        Else
            ComboBox12.Enabled = True
            ComboBox8.Enabled = True
        End If
        Button1.Enabled = True
        Me.Show()
    End Sub
    '==========================================================================================
    '==========================================================================================
    'ОТМЕНА
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Me.Close()
    End Sub
    '====================================================================================================================
    '=====================================================================================================================
    'ПРИМЕНИТЬ
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'окончательное формирование участков, многоконтурных участков, внутренних островков
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'нажата кнопка применить. Формируем участки и внутренние островки
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'ссылка на активный чертеж
        Dim ed As Editor = acDoc.Editor ' открываем базу данных чертежа
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'свойства линии
        Dim deskObject As String = ""
        Dim TypeUserLine As String = "Continuous"
        Dim LayerUserLine As String = "0"
        Dim WidthUserLine As Double = 30
        Dim ColorUserLine As Integer = 10
        Dim ScaleUserLine As Double = 1
        Dim CondUserLile As Double = 0
        Dim NameTableLinePS As String = "OBJECT"
        Dim idNameTableLinePS As ObjectId = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'свойства мультивыностки
        Dim NameStyleMleader As String = ComboBox11.Text
        Dim idMleaderStyle As ObjectId = Nothing
        Dim LayerMLeader As String = "0"
        Dim ColorMLeader As Integer = 7
        Dim NameTypeLineMLeader As String = "Continuous"
        Dim LineWeightMLeader As Double = -1
        Dim LineScaleMLeader As Double = -1
        Dim Size_Arrow As Double = 2
        Dim HText As Double = 2.5
        Dim TextStyleMleader As String = ""
        Dim idTextStyle As ObjectId = Nothing
        Dim NameTableMleaderPS As String = "OBJECT"
        Dim idTablePSMleader As ObjectId = Nothing
        Dim DeskMleader As String = ""
        Dim Note As String = "" 'примечание
        Dim numerator As String = ""
        Dim denumerator As String = ""
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'путь к шаблону
        Dim userCategoryObject As String = ComboBox8.Text.Trim
        Dim userObject As String = ComboBox12.Text.Trim
        Dim ArrayParcelProperties As String(,) = Nothing
        Dim ArrayMleaderProperties As String(,) = Nothing
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim mlstyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            For Each item As DBDictionaryEntry In mlstyles
                idMleaderStyle = item.Value
                Exit For
            Next
            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim acLineTypTbl As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim dictPropSetDefTemp = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Dim StText As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            idTextStyle = AcadFunction.FuncReturnIDTextStylebyName("CADASTRAL_OBJECT")
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Try
                If IO.File.Exists(PutchTemlateXml) = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'читаем все характеристики объекта и создаем неоходимые стили
                    Dim boolFindPropertiesObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(PutchTemlateXml, CategoryObject, userCategoryObject, "Полигон", userObject, ArrayParcelProperties)
                    If IsArray(ArrayParcelProperties) = True Then
                        For i As Integer = 0 To ArrayParcelProperties.GetUpperBound(1)
                            Dim nameFields As String = ArrayParcelProperties(0, i)
                            Dim value As String = ArrayParcelProperties(1, i)
                            If value.Trim.Length > 0 Then
                                If nameFields Like "Name" Then
                                    TypeUserLine = value.Trim
                                ElseIf nameFields Like "Layer" Then
                                    LayerUserLine = value.Trim
                                ElseIf nameFields Like "Description" Then
                                    deskObject = value.Trim
                                ElseIf nameFields Like "Color" Then
                                    ColorUserLine = Val(value.Trim)
                                    If ColorUserLine > 256 Then
                                        ColorUserLine = 256
                                    ElseIf ColorUserLine < 0 Then
                                        ColorUserLine = 0
                                    End If
                                ElseIf nameFields Like "LineWidth" Then
                                    WidthUserLine = Val(value)
                                ElseIf nameFields Like "LineScale" Then
                                    ScaleUserLine = Val(value)
                                    If ScaleUserLine <= 0 Then
                                        ScaleUserLine = 1
                                    End If
                                ElseIf nameFields Like "LineGWidth" Then
                                    CondUserLile = Val(value)
                                ElseIf nameFields Like "PS" Then
                                    NameTableLinePS = value.Trim
                                End If
                            End If
                        Next i
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'формируем тип линии
                    If acLineTypTbl.Has(TypeUserLine) = False Then
                        Dim arrayTypeLineProperties As String(,) = Nothing
                        Dim recTypeLine As LinetypeTableRecord = TOPOSymbols.FuncCreateLineTypeToXML(PutchTemlateXml, TypeUserLine)
                        If IsNothing(recTypeLine) = True Then
                            TypeUserLine = "Continuous"
                        Else
                            If recTypeLine.Id.IsNull = True Then
                                TypeUserLine = "Continuous"
                            End If
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'формируем слой
                    If acLyrTbl.Has(LayerUserLine) = False Then
                        Dim idLayer As ObjectId = TOPOSymbols.FuncCreateLayerToXML(PutchTemlateXml, LayerUserLine)
                        If IsNothing(idLayer) = True Then
                            LayerUserLine = "0"
                        Else
                            If idLayer.IsNull = True Then
                                LayerUserLine = "0"
                            End If
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'расширенные данные
                    idNameTableLinePS = FuncPropertySet.FuncCreateTablePSByXML(PutchTemlateXml, DataGridView2.Tag)
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'мультивыноска
                    If RadioButton1.Checked = False Then
                        'мультивыноска нужна
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'читаем все характеристики объекта и создаем неоходимые стили
                        Dim boolFindPropertiesMleader As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(PutchTemlateXml, CategoryObject, userCategoryObject, "Мультивыноска", NameStyleMleader, ArrayMleaderProperties)
                        If IsArray(ArrayMleaderProperties) = True Then
                            For i As Integer = 0 To ArrayMleaderProperties.GetUpperBound(1)
                                Dim nameFields As String = ArrayMleaderProperties(0, i)
                                Dim value As String = ArrayMleaderProperties(1, i)
                                If value.Trim.Length > 0 Then
                                    If nameFields Like "Name" Then
                                        NameStyleMleader = value.Trim
                                        Dim boolCopyMleader As Boolean = False
                                        Try
                                            idMleaderStyle = mlstyles.GetAt(NameStyleMleader)
                                            boolCopyMleader = False
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            boolCopyMleader = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemlateDwg, NameStyleMleader)
                                        End Try
                                        If boolCopyMleader = True Then
                                            Try
                                                idMleaderStyle = mlstyles.GetAt(NameStyleMleader)
                                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            End Try
                                        End If
                                    ElseIf nameFields Like "Layer" Then
                                        LayerMLeader = value.Trim
                                        If Mid(LayerMLeader, 1, 1) Like "_" Then
                                            LayerMLeader = LayerUserLine & LayerMLeader
                                        End If
                                    ElseIf nameFields Like "Color" Then
                                        ColorMLeader = Val(value.Trim)
                                        If ColorMLeader > 256 Then
                                            ColorMLeader = 256
                                        ElseIf ColorMLeader < 0 Then
                                            ColorMLeader = 0
                                        End If
                                    ElseIf nameFields Like "LineType" Then
                                        NameTypeLineMLeader = value
                                    ElseIf nameFields Like "LineWidth" Then
                                        LineWeightMLeader = Val(value)
                                    ElseIf nameFields Like "LineScale" Then
                                        LineScaleMLeader = Val(value)
                                        If LineScaleMLeader <= 0 Then
                                            LineScaleMLeader = 1
                                        End If
                                    ElseIf nameFields Like "ArrowSize" Then
                                        Size_Arrow = Val(value)
                                    ElseIf nameFields Like "TextStyle" Then
                                        TextStyleMleader = value
                                    ElseIf nameFields Like "TextHeight" Then
                                        HText = Val(value)
                                    ElseIf nameFields Like "PS" Then
                                        NameTableMleaderPS = value
                                    ElseIf nameFields Like "Note" Then
                                        Note = value
                                        If IsNothing(Note) = False Then
                                            If Note.Trim.Length > 0 Then
                                                Dim ArrayNote As String() = Note.Split(";")
                                                If IsArray(ArrayNote) = True Then
                                                    For j As Integer = 0 To ArrayNote.Length - 1
                                                        Dim str As String = ArrayNote(j)
                                                        Dim pos As Integer = str.IndexOf("=")
                                                        If pos > 0 Then
                                                            Dim NameField As String = Mid(str, 1, pos)
                                                            Dim valStr As String = Mid(str, pos + 2)
                                                            If j < ArrayNote.Length - 1 Then
                                                                For k As Integer = j + 1 To ArrayNote.Length - 1
                                                                    Dim TempStr As String = ArrayNote(k)
                                                                    If TempStr.IndexOf("=") = -1 Then
                                                                        valStr = valStr & ";" & TempStr
                                                                        j = k
                                                                    End If
                                                                Next k
                                                            End If
                                                            If NameField.Trim Like "numerator" Then
                                                                numerator = valStr
                                                            ElseIf NameField.Trim Like "denominator" Then
                                                                denumerator = valStr
                                                            End If
                                                        End If
                                                    Next j
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'формируем тип линии
                        If acLineTypTbl.Has(NameTypeLineMLeader) = False Then
                            Dim recTypeLine As LinetypeTableRecord = TOPOSymbols.FuncCreateLineTypeToXML(PutchTemlateXml, NameTypeLineMLeader)
                            If IsNothing(recTypeLine) = True Then
                                NameTypeLineMLeader = "Continuous"
                            Else
                                If recTypeLine.Id.IsNull = True Then
                                    NameTypeLineMLeader = "Continuous"
                                End If
                            End If
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'формируем слой
                        If acLyrTbl.Has(LayerMLeader) = False Then
                            Dim idLayer As ObjectId = TOPOSymbols.FuncCreateLayerToXML(PutchTemlateXml, LayerMLeader)
                            If IsNothing(idLayer) = True Then
                                LayerMLeader = "0"
                            Else
                                If idLayer.IsNull = True Then
                                    LayerMLeader = "0"
                                End If
                            End If
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'расширенные данные
                        idTablePSMleader = FuncPropertySet.FuncCreateTablePSByXML(PutchTemlateXml, NameTableMleaderPS)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'создаем текстовый стиль
                        If StText.Has(TextStyleMleader) = False Then
                            idTextStyle = TOPOSymbols.FuncCreateTextStyleToXML(PutchTemlateXml, TextStyleMleader)
                            If IsNothing(idTextStyle) = True Then
                                idTextStyle = db.Textstyle
                            End If
                            If idTextStyle.IsNull = True Then
                                idTextStyle = db.Textstyle
                            End If
                        End If
                    End If
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'преобразовываем объекты
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If dictionaryDataParcels.Count > 0 Then
                    Dim rng As Random = New Random
                    For Each dictParcel As KeyValuePair(Of Integer, String(,)) In dictionaryDataParcels
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'перебираем участки
                        Dim arrayPropertiesParcels As String(,) = dictParcel.Value
                        Dim prefParcel As String = ComboBox2.Text 'префикс участка
                        Dim oboznParcel As String = "" 'короткое обозначение
                        Dim cad_number As String = "" 'кадастровый номер исходного участка
                        Dim nameObject As String = "Объект"
                        Dim keyParameter As String = "Key"
                        Dim workType As String = "Work"
                        Dim globalID As Integer = rng.Next 'глобальный ID
                        Dim arrayHandles As String() = Nothing 'массив с контурами
                        Dim arrayNumberKounters As String() = Nothing 'массив с номерами контуров
                        Dim arrayRec As String(,) = Nothing 'результирующий массив для записи в PS
                        Dim countArrayRec As Integer = 0
                        Dim nameTablePS As String = ""
                        For j As Integer = 0 To arrayPropertiesParcels.GetUpperBound(1)
                            'обозначение участка (тег)
                            Dim fieldsName As String = arrayPropertiesParcels(1, j)
                            Dim value As String = arrayPropertiesParcels(2, j)
                            If fieldsName Like "NumberParcel" Then
                                oboznParcel = oboznParcel & value
                            ElseIf fieldsName Like "NumberCounter" Then
                                arrayNumberKounters = value.Trim.Split
                            ElseIf fieldsName Like "PS" Then
                                nameTablePS = value
                            ElseIf fieldsName Like "Handle" Then
                                arrayHandles = value.Trim.Split
                            ElseIf fieldsName Like "Cad_number" Then
                                cad_number = value.Trim
                                If cad_number.Length > 0 Then
                                    Dim strTemp As String = Mid(cad_number, cad_number.Length)
                                    If strTemp Like ":" Then
                                        cad_number = Mid(cad_number, 1, cad_number.Length - 1)
                                    End If
                                End If
                            ElseIf fieldsName Like "GLOBALID" Then
                                If IsArray(arrayHandles) = True Then
                                    If arrayHandles.Length > 1 Then
                                        arrayPropertiesParcels(2, j) = globalID
                                    End If
                                End If
                            ElseIf fieldsName Like "Name" Then
                                nameObject = value
                            ElseIf fieldsName Like "KeyParameter" Then
                                keyParameter = value
                            ElseIf fieldsName Like "TypeWork" Then
                                workType = value
                            End If
                        Next j
                        If prefParcel.IndexOf("00:00:0000000:0000:ЗУ") > -1 Then
                            Dim pos As Integer = cad_number.LastIndexOf(":")
                            If pos = cad_number.Length Then
                                oboznParcel = cad_number & "ЗУ" & oboznParcel
                            Else
                                oboznParcel = cad_number & ":ЗУ" & oboznParcel
                            End If
                        ElseIf prefParcel.IndexOf("0000:ЗУ") > -1 Then
                            Dim pos As Integer = cad_number.LastIndexOf(":")
                            oboznParcel = Mid(cad_number, pos + 1) & ":ЗУ" & oboznParcel
                        ElseIf prefParcel.IndexOf("00:00:0000000:0000/чзу") > -1 Then
                            oboznParcel = cad_number & "/чзу" & oboznParcel
                        ElseIf prefParcel.IndexOf("0000/чзу") > -1 Then
                            Dim pos As Integer = cad_number.LastIndexOf(":")
                            oboznParcel = Mid(cad_number, pos + 1) & "/чзу" & oboznParcel
                        Else
                            oboznParcel = prefParcel & oboznParcel
                        End If

                        If IsArray(arrayHandles) = True Then
                            For k As Integer = 0 To arrayHandles.Length - 1
                                Dim hgParcels As String = arrayHandles(k)
                                Dim LocalID As String = rng.Next
                                Dim OboznLocalParcel As String = oboznParcel
                                If IsArray(arrayNumberKounters) = True Then
                                    If arrayNumberKounters.Length > 1 Then
                                        If arrayNumberKounters.Length - 1 >= k Then
                                            Dim kounter As Integer = arrayNumberKounters(k)
                                            If kounter > 0 Then
                                                OboznLocalParcel = OboznLocalParcel & "(" & kounter & ")"
                                            End If
                                        End If
                                    End If
                                End If
                                Dim idEnt As ObjectId = AcadFunction.FuncReturnObjectID(hgParcels)
                                If IsNothing(idEnt) = False OrElse idEnt.IsNull = False Then
                                    '=====================================================================================================
                                    'получаем доступ к участку
                                    Dim objPLine As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(idEnt, OpenMode.ForWrite)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'удаляем совпадающие вершины
                                    Dim boolDblDelete As Boolean = AcadFunction.FuncPlineDelDubleVertex(objPLine)
                                    'заменяем дуговые элементы линейными
                                    Dim boolArcToLine As Boolean = AcadFunction.FuncReplaseArcToLineSegmentPline(objPLine)
                                    'вычисляем площадь участка
                                    Dim idParcel As String = rng.Next
                                    For j As Integer = 0 To arrayPropertiesParcels.GetUpperBound(1)
                                        'обозначение участка (тег)
                                        Dim fieldsName As String = arrayPropertiesParcels(1, j)
                                        Dim value As String = arrayPropertiesParcels(2, j)
                                        If fieldsName Like "ConditionalNumber" Then
                                            arrayPropertiesParcels(2, j) = OboznLocalParcel
                                        ElseIf fieldsName Like "LOCALID" Then
                                            arrayPropertiesParcels(2, j) = LocalID
                                        ElseIf fieldsName Like "Area" Then
                                            Dim colPoint3d As Point3dCollection = New Point3dCollection
                                            For k1 As Integer = 0 To objPLine.NumberOfVertices - 1
                                                Dim pt As Point3d = objPLine.GetPoint3dAt(k1)
                                                colPoint3d.Add(pt)
                                            Next k1
                                            Dim areaParcel As Double = MathFunction.FuncSQRTPolyline(colPoint3d)
                                            arrayPropertiesParcels(2, j) = areaParcel
                                        End If
                                    Next j
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'начинаем переименовывать линии
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'заносим характеристики
                                    objPLine.Layer = LayerUserLine
                                    objPLine.ColorIndex = ColorUserLine
                                    objPLine.LineWeight = WidthUserLine
                                    objPLine.Linetype = TypeUserLine
                                    objPLine.ConstantWidth = CondUserLile
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'PropertySet
                                    Dim ArrayTempData As String() = Nothing
                                    Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetTables(objPLine, ArrayTempData)
                                    Dim boolPropertySetAdd As Boolean = False
                                    If IsArray(ArrayTempData) = True Then
                                        For i1 As Integer = 0 To ArrayTempData.Length - 1
                                            Dim nameTableTemp As String = ArrayTempData(i1)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            If nameTableTemp Like NameTableLinePS Then
                                                'перезаписываем данные
                                                boolPropertySetAdd = FuncPropertySet.FuncWriteAllDataSetProperties(objPLine, arrayPropertiesParcels)
                                            End If
                                        Next i1
                                    End If
                                    If boolPropertySetAdd = False Then
                                        'присваиваем таблицу PS
                                        Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableLinePS)
                                        If IsNothing(idTablePs) = False Then
                                            If idTablePs.IsNull = False Then
                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(objPLine, idTablePs)
                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(objPLine, arrayPropertiesParcels)
                                            End If
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'вставляем мультивыноску с тестом
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    If ComboBox11.Text.Trim.Length > 0 Then
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'определение центра полигона 
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim CoordText As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(objPLine) 'находим центр полигона
                                        'формируем строку подписи
                                        '===============================================================================================================
                                        Dim ParcelNumerator As String = ""
                                        Dim ParcelDenumerator As String = ""
                                        If numerator.IndexOf("ConditionalNumber") > -1 Then
                                            ParcelNumerator = numerator.Replace("<ConditionalNumber>", OboznLocalParcel)
                                        End If
                                        If denumerator.IndexOf("ConditionalNumber") > -1 Then
                                            ParcelDenumerator = denumerator.Replace("<ConditionalNumber>", OboznLocalParcel)
                                        End If

                                        If numerator.IndexOf("Name") > -1 Then
                                            ParcelNumerator = numerator.Replace("<Name>", nameObject)
                                        End If
                                        If denumerator.IndexOf("Name") > -1 Then
                                            ParcelDenumerator = denumerator.Replace("<Name>", nameObject)
                                        End If

                                        If numerator.IndexOf("Description") > -1 Then
                                            ParcelNumerator = numerator.Replace("<Description>", nameObject)
                                        End If
                                        If denumerator.IndexOf("Description") > -1 Then
                                            ParcelDenumerator = denumerator.Replace("<Description>", nameObject)
                                        End If

                                        If numerator.IndexOf("KeyParameter") > -1 Then
                                            ParcelNumerator = numerator.Replace("<KeyParameter>", keyParameter)
                                        End If
                                        If denumerator.IndexOf("KeyParameter") > -1 Then
                                            ParcelDenumerator = denumerator.Replace("<KeyParameter>", keyParameter)
                                        End If

                                        If numerator.IndexOf("TypeWork") > -1 Then
                                            ParcelNumerator = numerator.Replace("<TypeWork>", workType)
                                        End If
                                        If denumerator.IndexOf("TypeWork") > -1 Then
                                            ParcelDenumerator = denumerator.Replace("<TypeWork>", workType)
                                        End If

                                        If numerator.IndexOf("Area") > -1 Then
                                            Dim StrId As Long = objPLine.Id.OldIdPtr
                                            Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%"
                                            ParcelNumerator = numerator.Replace("<Area>", AreaStr)
                                        End If
                                        If denumerator.IndexOf("Area") > -1 Then
                                            Dim StrId As Long = objPLine.Id.OldIdPtr
                                            Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%"
                                            ParcelDenumerator = denumerator.Replace("<Area>", AreaStr)
                                        End If

                                        If numerator.IndexOf("Lenght") > -1 Then
                                            Dim StrId As Long = objPLine.Id.OldIdPtr
                                            Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Length \f " & """" & "%lu2%pr0" & """" & ">%"
                                            ParcelNumerator = numerator.Replace("<Lenght>", AreaStr)
                                        End If
                                        If denumerator.IndexOf("Lenght") > -1 Then
                                            Dim StrId As Long = objPLine.Id.OldIdPtr
                                            Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Length \f " & """" & "%lu2%pr0" & """" & ">%"
                                            ParcelDenumerator = denumerator.Replace("<Lenght>", AreaStr)
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'формируем строку подписи
                                        Dim ContentMLeader As String = ParcelNumerator
                                        If denumerator.Trim.Length > 0 Then
                                            ContentMLeader = ContentMLeader & vbLf & ParcelDenumerator
                                        End If
                                        Dim MLeader1 As MLeader = Nothing
                                        '=========================================================================================================
                                        'формируем масштаб
                                        Dim M As Double = Val(ComboBox5.Text) / 1000  'масштаб
                                        'определяем точку вставки полки для мультивыноски
                                        Dim MaxPoint As Point3d = objPLine.GeometricExtents.MaxPoint
                                        If RadioButton2.Checked = True Then
                                            MLeader1 = AcadFunction.FuncDrawMleader2(CoordText, ContentMLeader, MaxPoint, idMleaderStyle, LayerMLeader, M, idTextStyle,,)
                                        Else
                                            MLeader1 = AcadFunction.FuncDrawMleader2(CoordText, ContentMLeader,, idMleaderStyle, LayerMLeader, M, idTextStyle,,)
                                        End If
                                        Dim ArrayOBJ1 As String(,) = Nothing
                                        Dim boolTableFindFieldsMLeader As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NameTableMleaderPS, ArrayOBJ1)
                                        If IsArray(ArrayOBJ1) = True Then
                                            For k1 As Integer = 0 To ArrayOBJ1.GetUpperBound(1)
                                                Dim NameField As String = ArrayOBJ1(1, k1)
                                                If NameField.Trim Like "LABEL" Then
                                                    ArrayOBJ1(2, k1) = ComboBox11.Text
                                                ElseIf NameField.Trim Like "Number" Then
                                                    ArrayOBJ1(2, k1) = "1"
                                                ElseIf NameField.Trim Like "IDRNG" Then
                                                    ArrayOBJ1(2, k1) = LocalID
                                                End If
                                            Next k1
                                        End If
                                        Dim idTableMLPs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableMleaderPS)
                                        If IsNothing(idTableMLPs) = False Then
                                            If idTableMLPs.IsNull = False Then
                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idTableMLPs)
                                                If boolCreateTbl = True Then
                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MLeader1, ArrayOBJ1)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next k
                        End If
                    Next
                End If
            Catch ex As System.NullReferenceException
            End Try
            tr.Commit()
        End Using
        Me.Hide()
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'прорисовка контура
    Private Sub PictureBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Me.Update()
        PictureBox1.Update()
        Me.Hide()
        Me.Show()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'прорисовка контура
    Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'отрисовка участков в графическом окне
        If dictionaryDataParcels.Count > 0 Then
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            ' открываем базу данных чертежа
            Dim ed As Editor = acDoc.Editor
            Dim db As Database = acDoc.Database
            Dim objPline As Polyline
            Dim pt3dmax As Point3d = New Point3d(0, 0, 0)
            Dim pt3dmin As Point3d = New Point3d(0, 0, 0)
            Dim Max_Pline As Point3d = New Point3d(0, 0, 0)
            Dim Min_Pline As Point3d = New Point3d(0, 0, 0)
            Dim TextFont As New System.Drawing.Font("Arial", 8, FontStyle.Regular)
            Dim myPen As New Pen(Drawing.Color.Black)

            If ComboBox8.Text.IndexOf("ервит") > -1 Then
                myPen = New Pen(Drawing.Color.Green)
            Else
                myPen = New Pen(Drawing.Color.Red)
            End If
            Dim g As Graphics
            g = e.Graphics
            g.Clear(Drawing.Color.AliceBlue)
            'вызываем транкзакцию
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Try
                    For Each dictParcel As KeyValuePair(Of Integer, String(,)) In dictionaryDataParcels
                        Dim arrayPropertiesParcels As String(,) = dictParcel.Value
                        Dim hgParcel As String() = arrayPropertiesParcels(2, 2).Split
                        Dim numKounters As String() = arrayPropertiesParcels(2, 1).Split
                        If IsArray(hgParcel) = True Then
                            For i1 As Integer = 0 To hgParcel.Length - 1
                                Dim HangleParcel As String = hgParcel(i1)
                                Dim NParcel As String = arrayPropertiesParcels(2, 0)
                                If IsArray(numKounters) = True Then
                                    If numKounters.Length > 1 Then
                                        If i1 <= numKounters.Length - 1 Then
                                            Dim numKounter As String = Val(numKounters(i1))
                                            If numKounter > 0 Then
                                                NParcel = NParcel & "(" & numKounter & ")"
                                            End If
                                        End If
                                    End If
                                End If
                                If IsNothing(HangleParcel) = False Then
                                    If HangleParcel.Trim.Length > 0 Then
                                        Dim m As Single = 0 'вычисляем масштабный фактор
                                        Dim d As Single = 0
                                        d = Math.Sqrt((Xmax - Xmin) ^ 2 + (Ymax - Ymin) ^ 2)
                                        m = Math.Round(600 / d, 4)
                                        'вычисляем смещение
                                        Dim dx1 As Single = Xmin * m
                                        Dim dx2 As Single = Xmax * m
                                        Dim c1 = -(600 - (dx2 - dx1)) / 2 + dx1
                                        Dim dy1 As Single = Ymin * m
                                        Dim dy2 As Single = Ymax * m
                                        Dim c2 = -(600 - (dy2 - dy1)) / 2 + dy1
                                        Dim ln As Long = Convert.ToInt64(HangleParcel, 16)
                                        Dim hn As Handle = New Handle(ln)
                                        Dim id As ObjectId = db.GetObjectId(False, hn, 0)
                                        Dim ent1 As Entity = tr.GetObject(id, OpenMode.ForRead)
                                        objPline = ent1
                                        'вставка обозначения в чертеж(подписи)
                                        Max_Pline = objPline.GeometricExtents.MaxPoint
                                        Min_Pline = objPline.GeometricExtents.MinPoint
                                        Dim XPline As Single = Math.Round((Max_Pline.X - Min_Pline.X) / 2, 2)
                                        Dim YPline As Single = Math.Round((Max_Pline.Y - Min_Pline.Y) / 2, 2)
                                        'смещение
                                        Dim Math1 As Integer = Val(ComboBox5.Text)
                                        Dim OfsetX As Double = 0 '(NumericUpDown1.Value * Math1) / 1000
                                        Dim OfsetY As Double = 0 '(NumericUpDown2.Value * Math1) / 1000
                                        'координата по оси Х
                                        Dim Xt As Single = (XPline + Min_Pline.X + OfsetX) * m - c1
                                        Xt = Math.Round(Xt, 2)
                                        'координата по оси Y
                                        Dim Yt As Single = (YPline + Min_Pline.Y + OfsetY) * m - c2
                                        Yt = Math.Round(600 - Yt, 2)
                                        'рисуем строку текста
                                        g.DrawString(NParcel, TextFont, Brushes.Black, Xt, Yt)
                                        'рисуем полигон
                                        Dim vn As Integer = objPline.NumberOfVertices - 2
                                        For j2 As Integer = 0 To vn
                                            Dim X1 As Single = Math.Round(objPline.GetPoint2dAt(j2).X * m - c1, 2)
                                            Dim Y1 As Single = Math.Round(objPline.GetPoint2dAt(j2).Y * m - c2 + 0, 2)
                                            Dim X2 As Single = Math.Round(objPline.GetPoint2dAt(j2 + 1).X * m - c1, 2)
                                            Dim Y2 As Single = Math.Round(objPline.GetPoint2dAt(j2 + 1).Y * m - c2 + 0, 2)
                                            Y1 = 600 - Y1
                                            Y2 = 600 - Y2
                                            g.DrawLine(myPen, X1, Y1, X2, Y2)
                                            If j2 = vn Then
                                                X1 = Math.Round(objPline.GetPoint2dAt(0).X * m - c1, 2)
                                                Y1 = Math.Round(objPline.GetPoint2dAt(0).Y * m - c2, 2)
                                                Y1 = 600 - Y1
                                                g.DrawLine(myPen, X2, Y2, X1, Y1)
                                            End If
                                        Next
                                    End If
                                End If
                            Next i1
                        End If
                    Next
                Catch ex As System.Exception
                    MsgBox(ex.ToString & vbLf & "Ошибка PictBox определения внешних границ выбранных контуров")
                End Try
            End Using
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'многоконтурный
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            NumericUpDown2.Enabled = True
            If NumericUpDown2.Value = 0 Then
                NumericUpDown2.Value = 1
            End If
            '===============================================================================
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "ConditionalNumber" Then
                            Dim str As String = row.Cells(1).Value
                            If IsNothing(str) = False Then
                                If str.Trim.Length > 0 Then
                                    Dim pos As Integer = str.IndexOf(")")
                                    If pos = str.Trim.Length Then
                                        Exit For
                                    Else
                                        str = str & "(" & NumericUpDown2.Value & ")"
                                        row.Cells(1).Value = str
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next
        Else
            NumericUpDown2.Enabled = False
            '===============================================================================
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "ConditionalNumber" Then
                            Dim str As String = row.Cells(1).Value
                            If IsNothing(str) = False Then
                                If str.Trim.Length > 0 Then
                                    Dim pos As Integer = str.IndexOf(")")
                                    If pos = str.Trim.Length - 1 Then
                                        Dim pos1 As Integer = str.LastIndexOf("(")
                                        str = Mid(str, 1, pos1)
                                        row.Cells(1).Value = str
                                        Exit For
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'категория земель
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '7. категории земель проектируемого участка
        Dim Kategory As String() = {"Земли сельскохозяйственного назначения", "Земли населенных пунктов", "Земли промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики, земли для обеспечения космической деятельности, земли обороны, безопасности и земли иного специального назначения", "Земли особо охраняемых территорий и объектов", "Земли лесного фонда", "Земли водного фонда", "Земли запаса", "Категория не установлена"}
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        selectObject.ListBox1.DataSource = Kategory
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Category" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'разрешенное использование
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '8. разрешенное использование
        Dim ArrayUtilization As String() = {"1 Сельскохозяйственное использование", "1.1	Растениеводство", "1.2 Выращивание зерновых и иных сельскохозяйственных культур", "1.3 Овощеводство", "1.4	Выращивание тонизирующих, лекарственных, цветочных культур",
            "1.5 Садоводство", "1.6 Выращивание льна и конопли", "1.7	Животноводство", "1.8 Скотоводство", "1.9	Звероводство", "1.1 Птицеводство", "1.11 Свиноводство", "1.12 Пчеловодство", "1.13 Рыбоводство", "1.14 Научное обеспечение сельского хозяйства",
            "1.15 Хранение и переработка сельскохозяйственной продукции", "1.16 Ведение личного подсобного хозяйства на полевых участках", "1.17 Питомники", "1.18 Обеспечение сельскохозяйственного производства", "1.19 Сенокошение", "1.2 Выпас сельскохозяйственных животных",
            "2 Жилая застройка", "2.1 Для индивидуального жилищного строительства", "2.1.1 Малоэтажная многоквартирная жилая застройка", "2.2 Для ведения личного подсобного хозяйства (приусадебный земельный участок)", "2.3 Блокированная жилая застройка", "2.4	Передвижное жилье",
            "2.5 Среднеэтажная жилая застройка", "2.6 Многоэтажная жилая застройка (высотная застройка)", "2.7 Обслуживание жилой застройки", "2.7.1	Хранение автотранспорта", "3 Общественное использование объектов капитального строительства", "3.1 Коммунальное обслуживание",
            "3.1.1 Предоставление коммунальных услуг", "3.1.2 Административные здания организаций, обеспечивающих предоставление коммунальных услуг", "3.2 Социальное обслуживание", "3.2.1 Дома социального обслуживания", "3.2.2 Оказание социальной помощи населению",
            "3.2.3 Оказание услуг связи", "3.2.4	Общежития", "3.3	Бытовое обслуживание", "3.4 Здравоохранение", "3.4.1 Амбулаторно-поликлиническое обслуживание", "3.4.2	Стационарное медицинское обслуживание", "3.4.3 Медицинские организации особого назначения", "3.5 Образование и просвещение",
            "3.5.1 Дошкольное, начальное и среднее общее образование", "3.5.2 Среднее и высшее профессиональное образование", "3.6 Культурное развитие", "3.6.1 Объекты культурно-досуговой деятельности", "3.6.2 Парки культуры и отдыха", "3.6.3 Цирки и зверинцы",
        "3.7 Религиозное использование", "3.7.1 Осуществление религиозных обрядов", "3.7.2 Религиозное управление и образование", "3.8 Общественное управление", "3.8.1 Государственное управление", "3.8.2 Представительская деятельность", "3.9 Обеспечение научной деятельности",
        "3.9.1 Обеспечение деятельности в области гидрометеорологии и смежных с ней областях", "3.9.2 Проведение научных исследований", "3.9.3 Проведение научных испытаний", "3.10 Ветеринарное обслуживание", "3.10.1	Амбулаторное ветеринарное обслуживание",
        "3.10.2 Приюты для животных", "4 Предпринимательство", "4.1 Деловое управление", "4.2 Объекты торговли (торговые центры, торгово-развлекательные центры (комплексы)", "4.3 Рынки", "4.4 Магазины", "4.5 Банковская и страховая деятельность", "4.6 Общественное питание",
        "4.7 Гостиничное обслуживание", "4.8	Развлечение", "4.8.1 Развлекательные мероприятия", "4.8.2	Проведение азартных игр", "4.8.3 Проведение азартных игр в игорных зонах", "4.9 Служебные гаражи", "4.9.1 Объекты дорожного сервиса", "4.9.1.1 Заправка транспортных средств",
        "4.9.1.2 Обеспечение дорожного отдыха", "4.9.1.3 Автомобильные мойки", "4.9.1.4 Ремонт автомобилей", "4.1 Выставочно-ярмарочная деятельность", "5 Отдых (рекреация)", "5.1 Спорт", "5.1.1 Обеспечение спортивно-зрелищных мероприятий", "5.1.2	Обеспечение занятий спортом в помещениях",
        "5.1.3 Площадки для занятий спортом", "5.1.4 Оборудованные площадки для занятий спортом", "5.1.5 Водный спорт", "5.1.6 Авиационный спорт", "5.1.7 Спортивные базы", "5.2 Природно-познавательный туризм", "5.2.1 Туристическое обслуживание", "5.3 Охота и рыбалка", "5.4 Причалы для маломерных судов",
"5.5 Поля для гольфа или конных прогулок", "6 Производственная деятельность", "6.1 Недропользование", "6.2	Тяжелая промышленность", "6.2.1 Автомобилестроительная промышленность", "6.3 Легкая промышленность", "6.3.1 Фармацевтическая промышленность", "6.4 Пищевая промышленность", "6.5 Нефтехимическая промышленность",
"6.6 Строительная промышленность", "6.7 Энергетика", "6.7.1 Атомная энергетика", "6.8 Связь", "6.9 Склад", "6.9.1 Складские площадки", "6.1 Обеспечение космической деятельности", "6.11 Целлюлозно-бумажная промышленность", "6.12 Научно-производственная деятельность", "7 Транспорт", "7.1 Железнодорожный транспорт",
"7.1.1 Железнодорожные пути", "7.1.2 Обслуживание железнодорожных перевозок", "7.2 Автомобильный транспорт", "7.2.1 Размещение автомобильных дорог", "7.2.2 Обслуживание перевозок пассажиров", "7.2.3 Стоянки транспорта общего пользования", "7.3 Водный транспорт", "7.4 Воздушный транспорт", "7.5	Трубопроводный транспорт",
"7.6 Внеуличный транспорт", "8 Обеспечение обороны и безопасности", "8.1 Обеспечение вооруженных сил", "8.2 Охрана Государственной границы Российской Федерации", "8.3 Обеспечение внутреннего правопорядка", "8.4 Обеспечение деятельности по исполнению наказаний", "9 Деятельность по особой охране и изучению природы",
"9.1 Охрана природных территорий", "9.2 Курортная деятельность", "9.2.1 Санаторная деятельность", "9.3 Историко-культурная деятельность", "10 Использование лесов", "10.1 Заготовка древесины", "10.2 Лесные плантации", "10.3 Заготовка лесных ресурсов", "10.4 Резервные леса", "11 Водные объекты",
"11.1 Общее пользование водными объектами", "11.2 Специальное пользование водными объектами", "11.3 Гидротехнические сооружения", "12 Земельные участки (территории) общего пользования", "12.0.1 Улично-дорожная сеть", "12.0.2 Благоустройство территории", "12.1	Ритуальная деятельность",
"12.2 Специальная деятельность", "12.3 Запас", "13 Земельные участки общего назначения", "13.1 Ведение огородничества", "13.2 Ведение садоводства"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        selectObject.ListBox1.DataSource = ArrayUtilization
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Utilization" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вид права
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        Dim ArrayVidPrava As String() = {"", "Частная собственность", "Государственная собственность", "Муниципальная собственность", "Собственность РФ", "Собственность субъекта РФ", "Неразграниченаая государственная собственность"}
        selectObject.ListBox1.DataSource = ArrayVidPrava
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Right" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вид кадастровых работ
    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '9. вид кадастровых работ
        Dim VidCadastralWork As String() = {"Образование земельного участка из земель, находящихся в государственной или муниципальной собственности",
                                            "Образование земельного участка путем раздела земельного участка",
                                            "Образование земельного участка путем раздела с сохранением земельного участка, раздел которого осуществлен, в измененных границах",
                                            "Образование земельного участка из состава земельного участка, представляющего собой единое землепользование",
                                            "Образование земельного участка путем объединения земельных участков",
                                            "Образование земельного участка при создании искусственного земельного участка",
                                            "Образование земельного участка путем перераспределения земельных участков",
                                            "Образование земельного участка путем перераспределения земельных участков или земель, находящихся в государственной или муниципальной собственности",
                                            "Уточнение местоположения границ и (или) площади земельного участка",
                                            "Уточнение описания местоположения границ и (или) площади земельного участка с целью исправления ошибки в описании местоположения его границ и (или) площади",
                                            "Образование части земельного участка",
                                            "Уточнение части земельного участка",
                                            "Установление сервитута",
                                            "Установление публичного сервитута"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = VidCadastralWork
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Method_create" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменение вида земельного участка
    Private Sub ComboBox8_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedValueChanged
        Dim NameObject As String = ComboBox8.Text
        If ComboBox11.Items.Count > 0 Then
            For i As Integer = 0 To ComboBox11.Items.Count - 1
                Dim nameMleader As String = ComboBox11.Items.Item(i).ToString
                If nameMleader Like NameObject Then
                    ComboBox11.Text = NameObject
                    Exit For
                End If
            Next i
        End If
    End Sub
    'Изменить тип объекта
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2. Читаем таблицу с объектов
        Dim tableName As String = ComboBox8.Text
        Dim arrayObject As String() = Nothing
        Dim boolFindObject As Boolean = TOPOSymbols.FuncReadTypeObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Полигон", arrayObject)
        If IsArray(arrayObject) = True Then
            ComboBox12.DataSource = arrayObject
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '3. Читаем характеристики объекта
        Dim deskObject As String = ComboBox12.Text
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolFindDataObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Полигон", deskObject, arrayPropertiesObject)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '4.читаем таблицу PS
        Dim userPSTable As String = "OBJECT"
        If IsArray(arrayPropertiesObject) = True Then
            For i As Integer = 0 To arrayPropertiesObject.GetUpperBound(1)
                If arrayPropertiesObject(0, i) Like "PS" Then
                    userPSTable = arrayPropertiesObject(1, i).Trim
                    Exit For
                End If
            Next i
        End If
        Dim arrayDataField As New Dictionary(Of Integer, String(,)) 'линейные
        Dim NameTablePS As String = TOPOSymbols.FuncReadDataTablePSToXML(PutchTemlateXml, userPSTable, arrayDataField)
        DataGridView2.Tag = userPSTable
        If userPSTable Like "PROJECT_PARCEL" Then
            DataGridView2.ContextMenuStrip = ContextMenuParcel
        ElseIf userPSTable Like "PROJECT_BULDING" Then
            DataGridView2.ContextMenuStrip = ContextMenuBuild
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем датагрид
        Dim firstOboznParcels As String = ComboBox2.Text
        If arrayDataField.Count > 0 Then
            DataGridView2.Rows.Clear()
            For i As Integer = 0 To arrayDataField.Count - 1
                DataGridView2.Rows.Add()
                If arrayDataField.ContainsKey(i) = True Then
                    Dim arrayField As String(,) = arrayDataField.Item(i)
                    If IsArray(arrayField) = True Then
                        For j As Integer = 0 To arrayField.GetUpperBound(1)
                            Dim nameF As String = arrayField(0, j)
                            Dim val As String = arrayField(1, j)
                            If nameF Like "Name" Then
                                DataGridView2.Rows(i).Cells(0).Tag = val
                            ElseIf nameF Like "Description" Then
                                DataGridView2.Rows(i).Cells(0).Value = val
                            ElseIf nameF Like "DefaultData" Then
                                DataGridView2.Rows(i).Cells(1).Value = val
                            End If
                            If val Like "ConditionalNumber" Then
                                DataGridView2.Rows(i).Cells(1).Value = firstOboznParcels
                            ElseIf val Like "LABEL" Then
                                DataGridView2.Rows(i).Cells(1).Value = deskObject
                            End If
                        Next j
                    End If
                End If
            Next i
            If DataGridView2.Rows(0).Cells(0).Value = Nothing Then
                DataGridView2.Rows.RemoveAt(0)
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'забираем все мультивыноски
        Dim arrayMleader As String() = Nothing
        Dim boolFindMLeader As Boolean = TOPOSymbols.FuncReadTypeObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Мультивыноска", arrayMleader)
        If IsArray(arrayMleader) = True Then
            ComboBox11.DataSource = arrayMleader
        Else
            RadioButton1.Checked = True
        End If
    End Sub

    'поменять создаваемый объект
    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '3. Читаем характеристики объекта
        Dim tableName As String = ComboBox8.Text
        Dim deskObject As String = ComboBox12.Text
        Dim arrayPropertiesObject As String(,) = Nothing
        Dim boolFindDataObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(PutchTemlateXml, CategoryObject, tableName, "Полигон", deskObject, arrayPropertiesObject)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '4.читаем таблицу PS
        Dim userPSTable As String = "OBJECT"
        If IsArray(arrayPropertiesObject) = True Then
            For i As Integer = 0 To arrayPropertiesObject.GetUpperBound(1)
                If arrayPropertiesObject(0, i) Like "PS" Then
                    userPSTable = arrayPropertiesObject(1, i).Trim
                    Exit For
                End If
            Next i
        End If
        'считываем поля таблицы PS
        'Dim arrayDataField As String(,) = Nothing
        Dim arrayDataField As New Dictionary(Of Integer, String(,)) 'линейные
        Dim NameTablePS As String = TOPOSymbols.FuncReadDataTablePSToXML(PutchTemlateXml, userPSTable, arrayDataField)
        DataGridView2.Tag = userPSTable
        If userPSTable Like "PROJECT_PARCEL" Then
            DataGridView2.ContextMenuStrip = ContextMenuParcel
        ElseIf userPSTable Like "PROJECT_BULDING" Then
            DataGridView2.ContextMenuStrip = ContextMenuBuild
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем датагрид
        Dim firstOboznParcels As String = ComboBox2.Text
        If arrayDataField.Count > 0 Then
            DataGridView2.Rows.Clear()
            For i As Integer = 0 To arrayDataField.Count - 1
                DataGridView2.Rows.Add()
                If arrayDataField.ContainsKey(i) = True Then
                    Dim arrayField As String(,) = arrayDataField.Item(i)
                    If IsArray(arrayField) = True Then
                        For j As Integer = 0 To arrayField.GetUpperBound(1)
                            Dim nameF As String = arrayField(0, j)
                            Dim val As String = arrayField(1, j)
                            If nameF Like "Name" Then
                                DataGridView2.Rows(i).Cells(0).Tag = val
                            ElseIf nameF Like "Description" Then
                                DataGridView2.Rows(i).Cells(0).Value = val
                            ElseIf nameF Like "Value" Then
                                DataGridView2.Rows(i).Cells(1).Value = val
                            End If
                            If val Like "ConditionalNumber" Then
                                DataGridView2.Rows(i).Cells(1).Value = firstOboznParcels
                            ElseIf val Like "LABEL" Then
                                DataGridView2.Rows(i).Cells(1).Value = deskObject
                            End If
                        Next j
                    End If
                End If
            Next i
            If DataGridView2.Rows(0).Cells(0).Value = Nothing Then
                DataGridView2.Rows.RemoveAt(0)
            End If
        End If
    End Sub

    Private Sub ОбоснованиеРазмераУчасткаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбоснованиеРазмераУчасткаToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'обоснование определение размера земельного участка
        Dim ArraySizeParcelDoc As String() = {"Приказ Минтранса РФ от 6 августа 2008 г. N 126 «Об утверждении Норм отвода земельных участков, необходимых для формирования полосы отвода железных дорог, а также норм расчета охранных зон железных дорог»",
                                              "Постановлением Правительства РФ от 2 сентября 2009 г. № 717 «О нормах отвода земель для размещения автомобильных дорог и (или) объектов дорожного сервиса»",
                                              "ГОСТ Р 55989-2014 «Магистральные газопроводы. Нормы проектирования на давление свыше 10 МПа. Основные требования», п.7.2.1 «Минимальная ширина полосы отвода земель для размещения МГ»",
                                              "СН 452-73 «Нормы отвода земель для магистральных трубопроводов»",
                                              "СН465-74 «Нормы отвода земель для электрических сетей напряжением 0.4-500 кВ»",
                                              "СН 461-74 «Нормы отвода земель для линий связи»"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = ArraySizeParcelDoc
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Rationale_size" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ОбоснованиеИзъятияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбоснованиеИзъятияToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'обоснование изъятия
        Dim ArrayObosnovanie As String() = {"Полное изъятие земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.56.3)",
                                            "Частичное изъятие земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.56.3)",
                                            "Резервирование земельного участка для государственных или муниципальных нужд (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.70.1)",
                                            "Установление публичного сервитута для прохода или проезда через земельный участок, в том числе в целях обеспечения свободного доступа граждан к водному объекту общего пользования и его береговой полосе (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление публичного сервитута для размещения на земельном участке межевых знаков, геодезических пунктов государственных геодезических сетей, гравиметрических пунктов, нивелирных пунктов и подъездов к ним (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление публичного сервитута для проведения дренажных и мелиоративных работ на земельном участке (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление публичного сервитута для забора (изъятия) водных ресурсов из водных объектов и водопоя (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление публичного сервитута для строительства, реконструкции, эксплуатации, капитального ремонта объектов электросетевого хозяйства, тепловых сетей, водопроводных сетей, сетей водоотведения, линий и сооружений связи, линейных объектов системы газоснабжения, нефтепроводов и нефтепродуктопроводов, их неотъемлемых технологических частей (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                            "Установление публичного сервитута для складирования строительных и иных материалов, размещение временных или вспомогательных сооружений (включая ограждения, бытовки, навесы) и (или) строительной техники, которые необходимы для обеспечения строительства, реконструкции, ремонта объектов транспортной инфраструктуры федерального, регионального или местного значения, на срок указанных строительства, реконструкции, ремонта (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                            "Установление публичного сервитута для устройства пересечений автомобильных дорог или железнодорожных путей с железнодорожными путями общего пользования на земельных участках, находящихся в государственной собственности, в границах полос отвода железных дорог, а также устройство пересечений автомобильных дорог или железнодорожных путей с автомобильными дорогами или примыканий автомобильных дорог к другим автомобильным дорогам на земельных участках, находящихся в государственной или муниципальной собственности, в границах полосы отвода автомобильной дороги (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.39.37)",
                                            "Установление публичного сервитута для размещения автомобильных дорог и железнодорожных путей в туннелях (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление публичного сервитута для реконструкции, капитального ремонта участков (частей) инженерных сооружений, являющихся линейными объектами (Земельный кодекс Российской Федерации от 25.10.2001 N 136-ФЗ, Ст.23)",
                                            "Установление сервитута для обеспечения прохода и проезда через соседний земельный участок, строительства, реконструкции и (или) эксплуатации линейных объектов, не препятствующих использованию земельного участка в соответствии с разрешенным использованием, а также других нужд собственника недвижимого имущества, которые не могут быть обеспечены без установления сервитута (Гражданский кодекс Российской Федерации от 30.11.1994 N 51-ФЗ, Ст.274)",
                                            "Нет"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = ArrayObosnovanie
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Confiscation" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ПрочитатьХарактеристикиИзPSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПрочитатьХарактеристикиИзPSToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'ссылка на активный чертеж
        Dim ed As Editor = acDoc.Editor ' открываем базу данных чертежа
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim selectObjResult As PromptEntityResult = ed.GetEntity(vbCrLf + "Выберите объект:")
            If selectObjResult.Status = PromptStatus.OK Then
                Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(selectObjResult.ObjectId, OpenMode.ForRead)
                'читаем XRecord
                Dim ArrayDataRec As String(,) = Nothing
                Dim boolOD As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(Ent, ArrayDataRec)
                If IsArray(ArrayDataRec) = True Then
                    For i As Integer = 0 To ArrayDataRec.GetUpperBound(1)
                        Dim NameField As String = ArrayDataRec(1, i)
                        Dim Zn As String = ArrayDataRec(2, i)
                        If IsNothing(Zn) = True Then Continue For
                        If Zn.Length = 0 Then Continue For
                        For Each row As DataGridViewRow In DataGridView2.Rows
                            Dim namefieldsFields As String = row.Cells(0).Tag
                            Dim valueRows As String = row.Cells(1).Value
                            If IsNothing(valueRows) = True OrElse valueRows = "" Then
                                If namefieldsFields Like NameField Then
                                    row.Cells(1).Value = Zn
                                End If
                            End If
                        Next
                    Next i
                End If
            End If
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub ПрочитатьХарактеристикиИзБлокаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПрочитатьХарактеристикиИзБлокаToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'ссылка на активный чертеж
        Dim ed As Editor = acDoc.Editor ' открываем базу данных чертежа
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim selectObjResult As PromptEntityResult = ed.GetEntity(vbCrLf + "Выберите объект:")
            If selectObjResult.Status = PromptStatus.OK Then
                Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(selectObjResult.ObjectId, OpenMode.ForRead)
                Dim ArrayDataRec As String(,) = Nothing
                If TypeOf Ent Is BlockReference Then
                    Dim blk As BlockReference = Ent
                    Dim boolReadDAta As Boolean = AcadFunction.FuncReadAllBlockAttr(blk, ArrayDataRec)
                    If IsArray(ArrayDataRec) Then
                        For i As Integer = 0 To ArrayDataRec.GetUpperBound(1)
                            Dim NameField As String = ArrayDataRec(0, i)
                            Dim Zn As String = ArrayDataRec(1, i)
                            If NameField Like "КН" Then
                                ArrayDataRec(0, i) = "Cad_number"
                            ElseIf NameField Like "Права" Then
                                ArrayDataRec(0, i) = "Right"
                            ElseIf NameField Like "Собственники" Then
                                ArrayDataRec(0, i) = "Patronymic"
                            ElseIf NameField Like "Арендаторы" Then
                                ArrayDataRec(0, i) = "Encumbrance"
                            ElseIf NameField Like "Адрес" Then
                                ArrayDataRec(0, i) = "Adress"
                            ElseIf NameField Like "Категория" Then
                                ArrayDataRec(0, i) = "Category"
                            ElseIf NameField Like "Использование" Then
                                ArrayDataRec(0, i) = "Utilization"
                            End If
                        Next i
                        For i As Integer = 0 To ArrayDataRec.GetUpperBound(1)
                            Dim NameField As String = ArrayDataRec(0, i)
                            Dim Zn As String = ArrayDataRec(1, i)
                            If IsNothing(Zn) = True Then Continue For
                            If Zn.Length = 0 Then Continue For
                            For Each row As DataGridViewRow In DataGridView2.Rows
                                Dim namefieldsFields As String = row.Cells(0).Tag
                                Dim valueRows As String = row.Cells(1).Value
                                If IsNothing(valueRows) = True OrElse valueRows = "" Then
                                    If namefieldsFields Like NameField Then
                                        row.Cells(1).Value = Zn
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            End If
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПрочитатьХарактеристикиИзРосреестраToolStripMenuItem.Click
        'Dim TextKNumber As String = "https://pkk.rosreestr.ru/api/features/1/59:1:4716054:825"
        Dim TextKNumber As String = ""
        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim namefieldsFields As String = row.Cells(0).Tag
            Dim valueRows As String = row.Cells(1).Value
            If IsNothing(valueRows) = False Then
                If namefieldsFields Like "Cad_number" Then
                    TextKNumber = row.Cells(1).Value
                End If
            End If
        Next
        Dim KNRezult As String = ""
        Dim StrArray As String() = TextKNumber.Split(":")
        If IsArray(StrArray) = True Then
            If StrArray.Length = 4 Then
                For i As Integer = 0 To 3
                    Dim str1 As Integer = CInt(StrArray(i))
                    If i = 0 Then
                        KNRezult = str1
                    Else
                        KNRezult = KNRezult & ":" & str1
                    End If
                Next
            End If
        End If
        If KNRezult.Trim.Length = 0 Then
            KNRezult = TextKNumber
        End If
        Dim url1 As String = "https://pkk.rosreestr.ru/api/features/1/" & KNRezult
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
        Dim wrGETURL1 As WebRequest = WebRequest.Create(url1)
        wrGETURL1.Credentials = System.Net.CredentialCache.DefaultCredentials
        '//Указываем сетевые учетные данные текущего контекста безопасности.
        wrGETURL1.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
        Dim response1 As WebResponse = wrGETURL1.GetResponse()
        Dim dataStream1 As Stream = response1.GetResponseStream()
        Dim reader1 As StreamReader = New StreamReader(dataStream1)
        Dim responseFromServer1 As String = reader1.ReadToEnd()
        Dim feature_oject As RosReestr_Parcels.Rootobject = New RosReestr_Parcels.Rootobject
        Dim boolReadurl2 As Boolean = FuncReadGSONLandCadaster.FuncReadJsonParcels(Of RosReestr_Parcels)(responseFromServer1, feature_oject)
        Dim ArrayDataRec As String(,) = Nothing
        Dim countArrayData As Integer = 0
        'делаем разбор найденного
        If IsNothing(feature_oject.feature) = False Then
            Dim addres As String = feature_oject.feature.attrs.address
            If IsNothing(addres) = False Then
                ReDim Preserve ArrayDataRec(2, countArrayData)
                ArrayDataRec(0, 0) = "PROJECT_PARCEL"
                ArrayDataRec(1, 0) = "Adress"
                ArrayDataRec(2, 0) = addres
                countArrayData += 1
            End If
            Dim KodCategory As String = feature_oject.feature.attrs.category_type
            Dim category_str As String = ""
            If IsNothing(KodCategory) = False Then
                category_str = FuncCadastr.FuncFindLandCategoryByCode(KodCategory)
                ReDim Preserve ArrayDataRec(2, countArrayData)
                ArrayDataRec(0, 0) = "PROJECT_PARCEL"
                ArrayDataRec(1, 0) = "Category"
                ArrayDataRec(2, 0) = category_str
                countArrayData += 1
            End If
            Dim UtilByDoc As String = feature_oject.feature.attrs.util_by_doc
            If IsNothing(UtilByDoc) = False Then
                ReDim Preserve ArrayDataRec(2, countArrayData)
                ArrayDataRec(0, 0) = "PROJECT_PARCEL"
                ArrayDataRec(1, 0) = "Utilization"
                ArrayDataRec(2, 0) = UtilByDoc
                countArrayData += 1
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор обозначение (префикс земельного участка)
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim oboznZu As String = ComboBox2.Text
        Dim numberZU As String = NumericUpDown4.Value.ToString
        Dim cad_number As String = "00:00:0000000:0000"
        If CheckBox7.Checked = True Then
            numberZU = numberZU & "(" & NumericUpDown2.Value.ToString & ")"
        End If

        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim namefieldsFields As String = row.Cells(0).Tag
            Dim valueRows As String = row.Cells(1).Value
            If IsNothing(valueRows) = False Then
                If namefieldsFields Like "Cad_number" Then
                    cad_number = row.Cells(1).Value
                    Exit For
                End If
            End If
        Next

        If oboznZu.IndexOf("00:00:0000000:0000:ЗУ") > -1 Then
            numberZU = cad_number & ":ЗУ" & numberZU
        ElseIf oboznZu.IndexOf("0000:ЗУ") > -1 Then
            Dim pos As Integer = cad_number.LastIndexOf(":")
            numberZU = Mid(cad_number, pos + 2) & ":ЗУ" & numberZU
        ElseIf oboznZu.IndexOf("00:00:0000000:0000/чзу") > -1 Then
            numberZU = cad_number & "/чзу" & numberZU
        Else
            numberZU = oboznZu & numberZU
        End If

        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim namefieldsFields As String = row.Cells(0).Tag
            Dim valueRows As String = row.Cells(1).Value
            If IsNothing(valueRows) = False Then
                If namefieldsFields Like "ConditionalNumber" Then
                    row.Cells(1).Value = numberZU
                    Exit For
                End If
            End If
        Next
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'обновление значений datagrid
    Private Sub DataGridView2_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellValueChanged
        Dim oboznZu As String = ComboBox2.Text
        Dim numberZU As String = NumericUpDown4.Value.ToString
        Dim cad_number As String = "00:00:0000000:0000"
        If CheckBox7.Checked = True Then
            numberZU = numberZU & "(" & NumericUpDown2.Value.ToString & ")"
        End If

        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim namefieldsFields As String = row.Cells(0).Tag
            Dim valueRows As String = row.Cells(1).Value
            If IsNothing(valueRows) = False Then
                If namefieldsFields Like "Cad_number" Then
                    cad_number = row.Cells(1).Value
                    Exit For
                End If
            End If
        Next

        If oboznZu.IndexOf("00:00:0000000:0000:ЗУ") > -1 Then
            Dim pos As Integer = cad_number.LastIndexOf(":")
            If pos = cad_number.Length Then
                numberZU = cad_number & "ЗУ" & numberZU
            Else
                numberZU = cad_number & ":ЗУ" & numberZU
            End If
        ElseIf oboznZu.IndexOf("0000:ЗУ") > -1 Then
            Dim pos As Integer = cad_number.LastIndexOf(":")
            numberZU = Mid(cad_number, pos + 2) & ":ЗУ" & numberZU
        ElseIf oboznZu.IndexOf("00:00:0000000:0000/чзу") > -1 Then
            numberZU = cad_number & "/чзу" & numberZU
        Else
            numberZU = oboznZu & numberZU
        End If

        For Each row As DataGridViewRow In DataGridView2.Rows
            Dim namefieldsFields As String = row.Cells(0).Tag
            Dim valueRows As String = row.Cells(1).Value
            If IsNothing(valueRows) = False Then
                If namefieldsFields Like "ConditionalNumber" Then
                    row.Cells(1).Value = numberZU
                    Exit For
                End If
            End If
        Next
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'здания
    Private Sub ТипСтроенияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТипСтроенияToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тив здания
        Dim ArrayObosnovanie As String() = {"Нежилое здание", "Жилой дом", "Многоквартирный дом"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = ArrayObosnovanie
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Type" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub


    Private Sub МатериалСтенToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МатериалСтенToolStripMenuItem.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тив здания
        Dim ArrayObosnovanie As String() = {"Стены", "Кирпичные", "Кирпичные облегченные", "Из природного камня", "Деревянные", "Рубленые", "Каркасно-засыпные", "Каркасно-обшивные",
                                           "Сборно-щитовые", "Дощатые", "Деревянный каркас без обшивки", "Смешанные", "Каменные и деревянные", "Каменные и бетонные", "Легкие из местных материалов",
                                           "Из прочих материалов", "Бетонные", "Монолитные", "Из мелких бетонных блоков", "Из легкобетонных панелей", "Железобетонные", "Крупнопанельные",
                                           "Каркасно-панельные", "Монолитные", "Крупноблочные", "Из унифицированных железобетонных элементов", "Из железобетонных сегментов", "Шлакобетонные", "Металлические"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = ArrayObosnovanie
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "MaterialWall" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ФункциональноеНазначениеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ФункциональноеНазначениеToolStripMenuItem.Click
        Dim ArrayObosnovanie As String() = {"Ф1 -здания, предназначенные для постоянного проживания и временного пребывания людей",
"Ф10.1 - здания дошкольных образовательных организаций, специализированных домов престарелых и инвалидов (неквартирные), спальные корпуса образовательных организаций с наличием интерната и детских организаций, здания медицинских организаций, предназначенные для оказания медицинской помощи в стационарных условиях (круглосуточно)",
"Ф10.2 - гостиницы, общежития(за исключением общежитий квартирного типа), спальные корпуса санаториев и домов отдыха общего типа, кемпингов",
"Ф1.3 - многоквартирные жилые дома, в том числе общежития квартирного типа",
"Ф10.4 - одноквартирные жилые дома, в том числе блокированные",
"Ф2 -здания зрелищных и культурно-просветительных учреждений",
"Ф20.1 - театры, кинотеатры, концертные залы, клубы, цирки, спортивные сооружения с трибунами, библиотеки и другие учреждения с расчетным числом посадочных мест для посетителей в закрытых помещениях",
"Ф20.2 - музеи, выставки, танцевальные залы и другие подобные учреждения в закрытых помещениях",
"Ф20.3 - здания учреждений, на открытом воздухе (Ф1.1)",
"Ф20.4 - здания учреждений, на открытом воздухе (Ф1.2)",
"Ф3 -здания организаций по обслуживанию населения",
"Ф30.1 - здания организаций торговли",
"Ф30.2 - здания организаций общественного питания",
"Ф30.3 - вокзалы",
"Ф3.4 - здания медицинских организаций, предназначенные для осуществления медицинской деятельности",
"Ф30.5 - помещения для посетителей организаций бытового и коммунального обслуживания с нерасчетным числом посадочных мест для посетителей",
"Ф30.6 - физкультурно - оздоровительные комплексы и спортивно-тренировочные учреждения с помещениями без трибун для зрителей, бытовые помещения, бани",
"Ф3.7 - объекты религиозного назначения",
"Ф4 -здания образовательных организаций, научных и проектных организаций, органов управления учреждений",
"Ф40.1 - здания общеобразовательных организаций, организаций дополнительного образования детей, профессиональных образовательных организаций",
"Ф40.2 - здания образовательных организаций высшего образования, организаций дополнительного профессионального образования",
"Ф40.3 - здания органов управления учреждений, проектно - конструкторских организаций, информационных и редакционно-издательских организаций, научных организаций, банков, контор, офисов",
"Ф4.4 - здания пожарных депо",
"Ф5 -здания производственного или складского назначения",
"Ф50.1 - производственные здания, сооружения, производственные и лабораторные помещения, мастерские, крематории",
"Ф5.2 - складские здания, сооружения, стоянки для автомобилей без технического обслуживания и ремонта, книгохранилища, архивы, складские помещения",
"Ф50.3 - здания сельскохозяйственного назначения"}
        Dim selectObject As SelectTemplateFiles = New SelectTemplateFiles()
        selectObject.ListBox1.DataSource = ArrayObosnovanie
        selectObject.ShowDialog()
        If selectObject.boolShow = True Then
            Dim selectObj = selectObject.ListBox1.SelectedItem
            Dim str As String = selectObj.ToString
            For Each row As DataGridViewRow In DataGridView2.Rows
                Dim namefieldsFields As String = row.Cells(0).Tag
                If IsNothing(namefieldsFields) = False Then
                    If namefieldsFields.Trim.Length > 0 Then
                        If namefieldsFields Like "Assignation" Then
                            row.Cells(1).Value = str
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub FormCreateParcels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub
End Class