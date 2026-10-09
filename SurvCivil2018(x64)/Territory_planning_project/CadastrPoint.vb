Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Windows
Public Class CadastrPoint
    '=============================================================================================
    'вставка отдельного блока
    <CommandMethod("InsertUserCadastralPoints")>
    Public Sub InsertUserCadastralPoints()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormLoad As FormCreatePoint = New FormCreatePoint
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        If MAnno = 10000 Then
            FormLoad.RadioButton1.Checked = True
            FormLoad.NumericUpDown1.Value = 10000
        ElseIf MAnno = 7000 Then
            FormLoad.RadioButton2.Checked = True
            FormLoad.NumericUpDown1.Value = 7000
        ElseIf MAnno = 5000 Then
            FormLoad.RadioButton3.Checked = True
            FormLoad.NumericUpDown1.Value = 5000
        ElseIf MAnno = 3000 Then
            FormLoad.RadioButton18.Checked = True
            FormLoad.NumericUpDown1.Value = 3000
        ElseIf MAnno = 2000 Then
            FormLoad.RadioButton4.Checked = True
            FormLoad.NumericUpDown1.Value = 2000
        ElseIf MAnno = 1000 Then
            FormLoad.RadioButton5.Checked = True
            FormLoad.NumericUpDown1.Value = 1000
        ElseIf MAnno = 500 Then
            FormLoad.RadioButton6.Checked = True
            FormLoad.NumericUpDown1.Value = 500
        ElseIf MAnno = 200 Then
            FormLoad.RadioButton17.Checked = True
            FormLoad.NumericUpDown1.Value = 200
        ElseIf MAnno = 100 Then
            FormLoad.RadioButton7.Checked = True
            FormLoad.NumericUpDown1.Value = 100
        ElseIf MAnno = 50 Then
            FormLoad.RadioButton14.Checked = True
            FormLoad.NumericUpDown1.Value = 50
        Else
            FormLoad.RadioButton8.Checked = True
            FormLoad.NumericUpDown1.Value = MAnno
        End If
        'вызываем диалоговое окно УСТАНОВКИ
        FormLoad.ShowDialog()
        If FormLoad.UstanovkiLoad = 0 Then
            Exit Sub
        End If
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
        Dim PutchTemplateDwg As String = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
        'ищем шаблон по ближе
        Try
            Dim TempDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                PutchTemplateDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
                If IO.File.Exists(PutchTemplateDir & FormLoad.ComboBox3.Text.Trim) = True Then
                    PutchTemplateDwg = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
                End If
            End If
        Catch ex As System.Security.SecurityException
        End Try
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        Dim idStyle As ObjectId = Nothing 'текстовые стиль
        'префикс номера
        Dim prefNumber As String = FormLoad.ComboBox4.Text
        'метод определения координат
        Dim MetodCoord As String = FormLoad.ComboBox5.Text
        '========================================================================================
        Dim M As Double = FormLoad.NumericUpDown1.Value 'масштаб
        Dim HText As Double = (M * FormLoad.NumericUpDown2.Value) / 1000 ' высота подписи
        Dim DText As Double = (M * FormLoad.NumericUpDown3.Value) / 1000 'расстояние смещения
        '==================================================================================================
        Dim NameObject As String = ""
        Dim LayerObject As String = ""
        Dim NamePsPoint As String = ""
        Dim ColorObject As String = 256
        Dim HPoint As Double = 1 'размер точки
        Dim DeskPoint As String = ""
        Dim attributeHeight As Double = 2
        Dim rotateAttrib As Double = 0
        Dim Delta As String = FormLoad.NumericUpDown6.Value
        Dim Zakrep As String = FormLoad.ComboBox2.Text
        Dim Dopusk As Double = FormLoad.NumericUpDown7.Value
        Dim idTablePs As ObjectId = Nothing
        Dim NamePSTable As String = ""
        Dim MashPoint As Scale3d = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayCounterNumber As String(,) = Nothing 'массив для нумерации всех контуров
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        Dim PointColection As ObjectIdCollection = New ObjectIdCollection()
        Dim idStyleMleader As ObjectId = Nothing
        Dim idMarkerStyle As ObjectId = Nothing
        Dim idLabelStyle As ObjectId = Nothing
        Dim dopuskXY As Double = FormLoad.NumericUpDown7.Value
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2.массив лоя мультивыносок
        Dim ArrayMLeader As String(,) = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем транзакцию
        Dim tr0 As Transaction = db.TransactionManager.StartTransaction()
        Using tr0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'создаем текстовый стиль для точек
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'создаем текстовый стиль для точек
            Dim st As TextStyleTable = CType(tr0.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
            If st.Has("CADASTRAL_POINT") = False Then
                Dim crTextstyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, "CADASTRAL_POINT", False)
                If crTextstyle = True Then
                    If st.Has("CADASTRAL_POINT") = True Then
                        idStyle = st.Item("CADASTRAL_POINT")
                    End If
                End If
                If IsNothing(idStyle) = True Then
                    idStyle = db.Textstyle
                End If
                If idStyle.IsNull = True Then
                    idStyle = db.Textstyle
                End If
            Else
                idStyle = st.Item("CADASTRAL_POINT")
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'таблицы xml
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ReDim ArrayTablesXML(1, 0)
            If FormLoad.RadioButton10.Checked = True Then
                If FormLoad.ComboBox1.Text.IndexOf("сервит") Then
                    ArrayTablesXML(0, 0) = "TEMPORARY_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок временного отвода"
                    Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
                Else
                    ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок постоянного отвода"
                    Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
                    Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
                End If
            ElseIf FormLoad.RadioButton1.Checked = True Then
                ArrayTablesXML(0, 0) = "PROJECT_BUILDING"
                ArrayTablesXML(1, 0) = "Здание или сооружение проектируемое"
            ElseIf FormLoad.RadioButton20.Checked = True Then
                ArrayTablesXML(0, 0) = "PROJECT_ZONA"
                ArrayTablesXML(1, 0) = "Зона с особыми условиями"
                Dim NameStyleMLeader As String() = {"ППТ-Территориальные зоны", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton21.Checked = True Then
                ArrayTablesXML(0, 0) = "URBAN_LINE"
                ArrayTablesXML(1, 0) = "Линии градостроительного регулирования"
                Dim NameStyleMLeader As String() = {"ППТ-Линии градостроительного регулирования", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton23.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_TREE"
                ArrayTablesXML(1, 0) = "Снос зеленых насаждений"
                Dim NameStyleMLeader As String() = {"ППТ-Граница контура", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            Else
                ArrayTablesXML(0, 0) = "PROJECT_LINE"
                ArrayTablesXML(1, 0) = "Сооружение линейного типа"
                Dim NameStyleMLeader As String() = {"ППТ-Инженерные коммуникации", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            End If
            NamePSTable = ArrayTablesXML(0, 0) 'имя таблицы дял записи OD
            '1. Считываем настройки с шаблона
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim ArrayPropertiesPoint As String(,) = Nothing
            Dim DeskObject As String = FormLoad.ComboBox1.Text
            'если шаблон найден
            If File.Exists(PutchTemplateDwg) = True Then
                'участки постоянного отвода
                If FormLoad.RadioButton10.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Земельные участки образуемые")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Земельные участки образуемые")
                    End If
                    'участки временного отвода
                ElseIf FormLoad.RadioButton22.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Здания (сооружения)")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Здания (сооружения)")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Здания (сооружения)")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Здания (сооружения)")
                    End If
                    'охранные зоны
                ElseIf FormLoad.RadioButton20.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Зоны с особыми условиями использования территорий")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Зоны с особыми условиями использования территорий")
                    End If
                    'линии градостроительного регулирования
                ElseIf FormLoad.RadioButton21.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Линии градостроительного регулирования")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Линии градостроительного регулирования")
                    End If
                Else
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Объекты капитального строительства (инженерные коммуникации)")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Объекты капитального строительства (инженерные коммуникации)")
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем разбор массива
            If IsArray(ArrayPropertiesPoint) = True Then
                For i As Integer = 0 To ArrayPropertiesPoint.GetUpperBound(1)
                    Dim NameInsObject As String = ArrayPropertiesPoint(1, i)
                    If ArrayPropertiesPoint(0, i) Like "Name" Then
                        If IsNothing(NameInsObject) = False Then
                            'загружаем блок Автокад
                            If FormLoad.RadioButton11.Checked = True Then
                                Dim acBlkTbl As BlockTable = tr0.GetObject(db.BlockTableId, OpenMode.ForRead)
                                If acBlkTbl.Has(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        NameObject = NameInsObject
                                    End If
                                Else
                                    NameObject = NameInsObject
                                End If
                                'загружаем мультивыноску
                            ElseIf FormLoad.RadioButton13.Checked = True Then
                                Dim mlStyles As DBDictionary = tr0.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                                If mlStyles.Contains(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        idStyleMleader = mlStyles.GetAt(NameInsObject)
                                    End If
                                Else
                                    idStyleMleader = mlStyles.GetAt(NameInsObject)
                                End If
                                'загружаем метку
                            ElseIf FormLoad.RadioButton16.Checked = True Then
                                Try
                                    idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                Catch ex2 As System.ArgumentException
                                End Try
                                If idStyle.IsNull = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyNoteLabelStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    Try
                                        idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                    Catch ex2 As System.ArgumentException
                                    End Try
                                End If
                                'точка COGO
                            ElseIf FormLoad.RadioButton12.Checked = True Then
                                'копируем ключ-описатель
                                Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db).SearchOrder
                                Dim boolKey As Boolean = False
                                If pdKeySetIds.Count > 0 Then
                                    For Each idKeySet As ObjectId In pdKeySetIds
                                        Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = tr0.GetObject(idKeySet, OpenMode.ForRead)
                                        Dim idDesk As ObjectId = Nothing
                                        Try
                                            idDesk = DeskKeySet.Item(NameInsObject)
                                        Catch ex As System.ArgumentException
                                        End Try
                                        If IsNothing(NameInsObject) = False Then
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
                                    Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplateDwg, NameInsObject)
                                End If
                                NameObject = NameInsObject
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "PointGroup" Then
                        'копируем группу точек
                        Dim idGroupPoint As ObjectId = Nothing
                        Try
                            idGroupPoint = CivilDoc.PointGroups.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idGroupPoint) = False Then
                            If idGroupPoint.IsNull = True Then
                                Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                            End If
                        Else
                            Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "MarkerStyle" Then
                        'копируем стиль маркера
                        Try
                            idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idMarkerStyle) = False Then
                            If idMarkerStyle.IsNull = True Then
                                Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                Try
                                    idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                        Else
                            Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                            Try
                                idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                            Catch ex As System.ArgumentException
                            End Try
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Layer" Then
                        Dim LayerinsBlk As String = ArrayPropertiesPoint(1, i)
                        If IsNothing(LayerinsBlk) = False Then
                            Dim acLyrTbl As LayerTable = tr0.GetObject(db.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(LayerinsBlk) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, LayerinsBlk)
                                If boolCopyLayer = True Then
                                    LayerObject = LayerinsBlk
                                End If
                            Else
                                LayerObject = LayerinsBlk
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Description" Then
                        DeskPoint = ArrayPropertiesPoint(1, i)
                    ElseIf ArrayPropertiesPoint(0, i) Like "PS" Then
                        NamePsPoint = ArrayPropertiesPoint(1, i)
                        If IsNothing(NamePsPoint) = False Then
                            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                            If dictPropSetDefTemp.Has(NamePsPoint, tr0) = False Then
                                Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NamePsPoint)
                            End If
                        End If
                        idTablePs = FuncPropertySet.FuncReturnIDDictSetProperties(NamePsPoint)
                    ElseIf ArrayPropertiesPoint(0, i) Like "Color" Then
                        Dim IndexColor As String = ArrayPropertiesPoint(1, i)
                        If IndexColor Like "ПоСлою" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ПоБлоку" Then
                            ColorObject = 0
                        ElseIf IndexColor Like "ByLayer" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ByBlock" Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) < 1 Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) > 255 Then
                            ColorObject = 256
                        Else
                            ColorObject = Val(IndexColor)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Scale" Then
                        HPoint = Val(ArrayPropertiesPoint(1, i))
                        If HPoint <= 0 Then
                            HPoint = 1
                        End If
                        MashPoint = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
                    ElseIf ArrayPropertiesPoint(0, i) Like "AttributeHeight" Then
                        attributeHeight = Val(ArrayPropertiesPoint(1, i))
                        If attributeHeight <= 0 Then
                            attributeHeight = 2
                        End If
                    End If
                Next i
                tr0.Commit()
            End If
        End Using
        Dim NumbPoint As Integer = 1
        Dim prPointRes1 As PromptEntityResult = Nothing
        Dim idParentObj As String = ""
        Dim ArrayData As String(,) = Nothing
        Dim LocalIDObject As String = ""
        Dim MPoint As Double = 1
        Dim Category As String = ""
        Dim clockPline As Boolean = True
        Do
            Try
                prPointRes1 = ed.GetEntity(vbLf & "Выберите точку ближайшую к узлу полилинии:")
                If prPointRes1.Status = PromptStatus.OK Then
                    Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                    If IsDBNull(entObjID) = False Then
                        Dim tr As Transaction = db.TransactionManager.StartTransaction()
                        Using tr
                            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                            If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = entObj
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'новый контур спрашиваем заново установки
                                If Not (Pline.Id.ToString Like idParentObj) Then
                                    Dim idRngOstrov As String = FuncPropertySet.FuncReadValueSetProperties(Pline, "IDKOUNTER", "CADASTR_OSTROV")
                                    If idRngOstrov.Trim.Length > 2 Then
                                        Dim ParententObj As Autodesk.AutoCAD.DatabaseServices.Entity = FuncCadastr.FuncFindObjectByLocalID(idRngOstrov)
                                        If IsNothing(ParententObj) = False Then
                                            If ParententObj.Id.IsNull = False Then
                                                Category = FuncPropertySet.FuncReadValueSetProperties(ParententObj, "Category")
                                            End If
                                        End If
                                    Else
                                        Category = FuncPropertySet.FuncReadValueSetProperties(Pline, "Category")
                                    End If
                                    LocalIDObject = FuncPropertySet.FuncReadValueSetProperties(Pline, "LOCALID")
                                    '=================================================================================================
                                    'запоминаем id выбранной полилинии
                                    idParentObj = Pline.Id.ToString
                                    '=================================================================================================
                                    'берем цвет и слой выбранной полилинии
                                    If ColorObject = 0 Then
                                        ColorObject = Pline.ColorIndex
                                    End If
                                    If LayerObject.Trim.Length = 0 Then
                                        LayerObject = Pline.Layer
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    '=========================================================================================
                                    'начальный номер точки
                                    If FormLoad.RadioButton15.Checked = True Then
                                        If FormLoad.RadioButton12.Checked = False Then
                                            Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(vbLf & "Введите номер точки: ") 'запрашиваем номер
                                            NumbPoint = pStrRes.Value
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    If FormLoad.CheckBox5.Checked = False Then
                                        Dim pos1 As Integer = -1
                                        If Category.IndexOf("ельскохоз") > -1 Then
                                            Delta = 2.5
                                        ElseIf Category.IndexOf("промышл") > -1 Then
                                            Delta = 0.5
                                        ElseIf Category.IndexOf("водн") > -1 Then
                                            Delta = 2.5
                                        ElseIf Category.IndexOf("запас") > -1 Then
                                            Delta = 2.5
                                        ElseIf Category.IndexOf("особо") > -1 Then
                                            Delta = 2.5
                                        ElseIf Category.IndexOf("лесн") > -1 Then
                                            Delta = 5
                                        ElseIf Category.IndexOf("насел") > -1 Then
                                            Delta = 0.1
                                        End If
                                    End If
                                    If Pline.NumberOfVertices > 2 Then
                                        Dim area As Double = MathFunction.GetArea(Pline)
                                        If area > 0 Then
                                            clockPline = False
                                        End If
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'начинаем нумеровать полигон
                                '==================================================================
                                Dim Vert0 As Point2d = New Point2d(0, 0)
                                Dim Vert1 As Point2d = New Point2d(0, 0)
                                Dim Vert2 As Point2d = New Point2d(0, 0)
                                Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                Dim pointPline As Point2d = Nothing
                                Dim Dist As Double = 1000000
                                Dim NumberVertexSelect As Integer = 0
                                For i As Integer = 0 To Pline.NumberOfVertices - 1
                                    Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                    Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                    If L < Dist Then
                                        Dist = L
                                        pointPline = tempPoint
                                        NumberVertexSelect = i
                                    End If
                                Next i
                                If NumberVertexSelect = 0 Then
                                    Vert0 = Pline.GetPoint2dAt(Pline.NumberOfVertices - 1)
                                    Vert1 = Pline.GetPoint2dAt(0)
                                    Vert2 = Pline.GetPoint2dAt(1)
                                ElseIf NumberVertexSelect = Pline.NumberOfVertices - 1 Then
                                    Vert0 = Pline.GetPoint2dAt(Pline.NumberOfVertices - 2)
                                    Vert1 = Pline.GetPoint2dAt(Pline.NumberOfVertices - 1)
                                    Vert2 = Pline.GetPoint2dAt(0)
                                Else
                                    Vert0 = Pline.GetPoint2dAt(NumberVertexSelect - 1)
                                    Vert1 = Pline.GetPoint2dAt(NumberVertexSelect)
                                    Vert2 = Pline.GetPoint2dAt(NumberVertexSelect + 1)
                                End If
                                Dim insAttr As Double() = Nothing 'массив с хорактеристиками
                                Dim boolInsAttr As Boolean = FuncCadastr.FuncReplaseInsertCoordAttr2(Vert0, Vert1, Vert2, insAttr, DText, clockPline)
                                Dim point3d2 As Point3d = New Point3d(Vert1.X, Vert1.Y, 0)
                                Dim NamePointZ As String = prefNumber & NumbPoint
                                '=================================================================================
                                'формируем массив для записи
                                Dim ArrayOBJ1 As String(,) = Nothing
                                ReDim Preserve ArrayOBJ1(2, 5)
                                ArrayOBJ1(0, 0) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 0) = "LABEL"
                                ArrayOBJ1(2, 0) = DeskPoint

                                ArrayOBJ1(0, 1) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 1) = "Number"
                                ArrayOBJ1(2, 1) = NamePointZ

                                ArrayOBJ1(0, 2) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 2) = "IDRNG"
                                ArrayOBJ1(2, 2) = LocalIDObject

                                ArrayOBJ1(0, 3) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 3) = "Delta"
                                ArrayOBJ1(2, 3) = Delta

                                ArrayOBJ1(0, 4) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 4) = "FastDesk"
                                ArrayOBJ1(2, 4) = Zakrep

                                ArrayOBJ1(0, 5) = "CADASTRAL_POINT"
                                ArrayOBJ1(1, 5) = "MethodCoordinates"
                                ArrayOBJ1(2, 5) = MetodCoord
                                '=====================================================================================
                                If FormLoad.RadioButton11.Checked = True Then
                                    'вставляем блок Автокад
                                    Dim blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(point3d1, NameObject, LayerObject, ColorObject, (HPoint * M) / 1000)
                                    '(Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
                                    If IsNothing(blk) = False Then
                                        Dim arrayDataAttrib As String(,) = {}
                                        ReDim arrayDataAttrib(6, 3)
                                        arrayDataAttrib(0, 0) = "NUMBER"
                                        arrayDataAttrib(1, 0) = NamePointZ
                                        arrayDataAttrib(3, 0) = (attributeHeight * M) / 1000
                                        arrayDataAttrib(4, 0) = rotateAttrib

                                        arrayDataAttrib(0, 1) = "DELTA"
                                        arrayDataAttrib(1, 1) = Delta

                                        arrayDataAttrib(0, 2) = "FASTDESK"
                                        arrayDataAttrib(1, 2) = Zakrep

                                        arrayDataAttrib(0, 3) = "METHODCOORDINATES"
                                        arrayDataAttrib(1, 3) = MetodCoord
                                        Dim boolWriteAttrai As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, arrayDataAttrib)
                                        'Dim InsAttribRez As ObjectId = FuncCadastr.FuncInsertCadastrPoint(point3d1, insAttr, NamePointZ, Delta, Zakrep, HText, LayerObject, NameObject, MashPoint, 0, idStyle, MetodCoord)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        blk.ColorIndex = ColorObject
                                        If FormLoad.CheckBox4.Checked = True Then
                                            PointColection.Add(blk.Id)
                                        End If
                                        '=======================================================================================================
                                        If IsNothing(idTablePs) = False Then
                                            If idTablePs.IsNull = False Then
                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(blk, idTablePs)
                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blk, ArrayOBJ1)
                                            End If
                                        End If
                                    End If
                                ElseIf FormLoad.RadioButton12.Checked = True Then
                                    'вставляем точку COGO
                                    Dim InsAttribRez As ObjectId = FuncSurvey.FuncCreateCoGoPoint(point3d2, NameObject)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем расширенные данные в точку
                                    If IsNothing(InsAttribRez) = False Then
                                        If InsAttribRez.IsNull = False Then
                                            Dim CoGoObj As Autodesk.Civil.DatabaseServices.CogoPoint = tr.GetObject(InsAttribRez, OpenMode.ForWrite)
                                            CoGoObj.ColorIndex = ColorObject
                                            If FormLoad.CheckBox4.Checked = True Then
                                                PointColection.Add(CoGoObj.Id)
                                            End If
                                            If prefNumber.Trim.Length > 0 Then
                                                Dim boolWriteUDP As Boolean = FuncSurvey.FuncWriteUDPValue(CoGoObj, "CADASTRAL_POINT", "Pref_number", prefNumber)
                                            End If
                                            '=======================================================================================================
                                            If IsNothing(idTablePs) = False Then
                                                If idTablePs.IsNull = False Then
                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoObj, idTablePs)
                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoObj, ArrayOBJ1)
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf FormLoad.RadioButton13.Checked = True Then
                                    'вставляем мультивыноску
                                    If FormLoad.RadioButton15.Checked = True Then
                                        Dim offsetx As Double = insAttr(0) - point3d1.X
                                        Dim offsety As Double = insAttr(1) - point3d1.Y
                                        Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d2.Convert2d(New Plane), NamePointZ, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем расширенные данные в точку
                                        If IsNothing(InsAttribRez) = False Then
                                            If InsAttribRez.IsNull = False Then
                                                Dim MleaderObj As MLeader = tr.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                MleaderObj.ColorIndex = ColorObject
                                                MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                If FormLoad.CheckBox4.Checked = True Then
                                                    PointColection.Add(MleaderObj.Id)
                                                End If
                                                '=======================================================================================================
                                                If IsNothing(idTablePs) = False Then
                                                    If idTablePs.IsNull = False Then
                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MleaderObj, idTablePs)
                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MleaderObj, ArrayOBJ1)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        Dim x As String = FuncFormatZN.FuncFormatCoordinate(point3d1.Y, 2)
                                        Dim y As String = FuncFormatZN.FuncFormatCoordinate(point3d1.X, 2)
                                        Dim TextMleader As String = "X=" & x & vbLf & "Y=" & y
                                        Dim offsetx As Double = insAttr(0) - point3d1.X
                                        Dim offsety As Double = insAttr(1) - point3d1.Y
                                        Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d2.Convert2d(New Plane), TextMleader, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                        If IsNothing(InsAttribRez) = False Then
                                            If InsAttribRez.IsNull = False Then
                                                Dim MleaderObj As MLeader = tr.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                MleaderObj.ColorIndex = ColorObject
                                                MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                If FormLoad.CheckBox4.Checked = True Then
                                                    PointColection.Add(InsAttribRez)
                                                End If
                                            End If
                                        End If
                                    End If
                                ElseIf FormLoad.RadioButton16.Checked = True Then
                                    'вставляем метку Civil3D
                                    If FormLoad.RadioButton15.Checked = True Then

                                    Else
                                        Dim IDlb As ObjectId = Autodesk.Civil.DatabaseServices.NoteLabel.Create(point3d2, idLabelStyle, idMarkerStyle)
                                        If FormLoad.CheckBox4.Checked = True Then
                                            PointColection.Add(IDlb)
                                        End If
                                        Dim lb As Autodesk.Civil.DatabaseServices.NoteLabel = tr.GetObject(IDlb, OpenMode.ForWrite)
                                        lb.Layer = LayerObject
                                        lb.ColorIndex = ColorObject
                                        lb.LabelLocation = New Point3d(insAttr(0), insAttr(1), 0)
                                    End If
                                End If
                                NumbPoint += 1
                            End If
                            tr.Commit()
                        End Using
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Loop Until prPointRes1.Status <> PromptStatus.OK
    End Sub
    '==============================================================================================
    'нумеровать полигоны
    <CommandMethod("AutoNumberUserPolygons")>
    Public Sub AutoNumberUserPolygons()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormLoad As FormCreatePoint = New FormCreatePoint
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        If MAnno = 10000 Then
            FormLoad.RadioButton1.Checked = True
            FormLoad.NumericUpDown1.Value = 10000
        ElseIf MAnno = 7000 Then
            FormLoad.RadioButton2.Checked = True
            FormLoad.NumericUpDown1.Value = 7000
        ElseIf MAnno = 5000 Then
            FormLoad.RadioButton3.Checked = True
            FormLoad.NumericUpDown1.Value = 5000
        ElseIf MAnno = 3000 Then
            FormLoad.RadioButton18.Checked = True
            FormLoad.NumericUpDown1.Value = 3000
        ElseIf MAnno = 2000 Then
            FormLoad.RadioButton4.Checked = True
            FormLoad.NumericUpDown1.Value = 2000
        ElseIf MAnno = 1000 Then
            FormLoad.RadioButton5.Checked = True
            FormLoad.NumericUpDown1.Value = 1000
        ElseIf MAnno = 500 Then
            FormLoad.RadioButton6.Checked = True
            FormLoad.NumericUpDown1.Value = 500
        ElseIf MAnno = 200 Then
            FormLoad.RadioButton17.Checked = True
            FormLoad.NumericUpDown1.Value = 200
        ElseIf MAnno = 100 Then
            FormLoad.RadioButton7.Checked = True
            FormLoad.NumericUpDown1.Value = 100
        ElseIf MAnno = 50 Then
            FormLoad.RadioButton14.Checked = True
            FormLoad.NumericUpDown1.Value = 50
        Else
            FormLoad.RadioButton8.Checked = True
            FormLoad.NumericUpDown1.Value = MAnno
        End If
        'вызываем диалоговое окно УСТАНОВКИ
        FormLoad.ShowDialog()
        If FormLoad.UstanovkiLoad = 0 Then
            Exit Sub
        End If
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
        Dim PutchTemplateDwg As String = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
        'ищем шаблон по ближе
        Try
            Dim TempDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                PutchTemplateDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
                If IO.File.Exists(PutchTemplateDir & FormLoad.ComboBox3.Text.Trim) = True Then
                    PutchTemplateDwg = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
                End If
            End If
        Catch ex As System.Security.SecurityException
        End Try
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        Dim idStyle As ObjectId = Nothing 'текстовые стиль
        'префикс номера
        Dim prefNumber As String = FormLoad.ComboBox4.Text
        'метод определения координат
        Dim MetodCoord As String = FormLoad.ComboBox5.Text
        '========================================================================================
        Dim M As Double = FormLoad.NumericUpDown1.Value 'масштаб
        Dim HText As Double = (M * FormLoad.NumericUpDown2.Value) / 1000 ' высота подписи
        Dim DText As Double = (M * FormLoad.NumericUpDown3.Value) / 1000 'расстояние смещения
        '==================================================================================================
        Dim NameObject As String = "" 'имя блока
        Dim LayerObject As String = "" 'слой блока
        Dim NamePsPoint As String = "" 'таблица PS
        Dim ColorObject As String = 256 'цвет блока
        Dim HPoint As Double = 1 'размер точки
        Dim attributeHeight As Double = 2
        Dim rotateAttrib As Double = 0
        Dim DeskPoint As String = ""
        Dim Delta As String = FormLoad.NumericUpDown6.Value
        Dim Zakrep As String = FormLoad.ComboBox2.Text
        Dim Dopusk As Double = FormLoad.NumericUpDown7.Value
        Dim idTablePs As ObjectId = Nothing
        Dim NamePSTable As String = ""
        Dim MashPoint As Scale3d = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayCounterNumber As String(,) = Nothing 'массив для нумерации всех контуров
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        Dim PointColection As ObjectIdCollection = New ObjectIdCollection()
        Dim idStyleMleader As ObjectId = Nothing
        Dim idMarkerStyle As ObjectId = Nothing
        Dim idLabelStyle As ObjectId = Nothing
        Dim dopuskXY As Double = FormLoad.NumericUpDown7.Value
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2.массив лоя мультивыносок
        Dim ArrayMLeader As String(,) = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем транзакцию
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'создаем текстовый стиль для точек
            Dim st As TextStyleTable = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
            If st.Has("CADASTRAL_POINT") = False Then
                Dim crTextstyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, "CADASTRAL_POINT", False)
                If crTextstyle = True Then
                    If st.Has("CADASTRAL_POINT") = True Then
                        idStyle = st.Item("CADASTRAL_POINT")
                    End If
                End If
                If IsNothing(idStyle) = True Then
                    idStyle = db.Textstyle
                End If
                If idStyle.IsNull = True Then
                    idStyle = db.Textstyle
                End If
            Else
                idStyle = st.Item("CADASTRAL_POINT")
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'таблицы xml
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If FormLoad.RadioButton10.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                If FormLoad.ComboBox1.Text.IndexOf("сервит") > 0 Then
                    ArrayTablesXML(0, 0) = "TEMPORARY_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок временного отвода"
                Else
                    ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок постоянного отвода"
                End If
                Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton9.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_BUILDING"
                ArrayTablesXML(1, 0) = "Здание или сооружение проектируемое"
            ElseIf FormLoad.RadioButton20.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_ZONA"
                ArrayTablesXML(1, 0) = "Зона с особыми условиями"
                Dim NameStyleMLeader As String() = {"ППТ-Территориальные зоны", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton21.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "URBAN_LINE"
                ArrayTablesXML(1, 0) = "Линии градостроительного регулирования"
                Dim NameStyleMLeader As String() = {"ППТ-Линии градостроительного регулирования", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton23.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_TREE"
                ArrayTablesXML(1, 0) = "Снос зеленых насаждений"
                Dim NameStyleMLeader As String() = {"ППТ-Граница контура", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            Else
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_LINE"
                ArrayTablesXML(1, 0) = "Сооружение линейного типа"
                Dim NameStyleMLeader As String() = {"ППТ-Инженерные коммуникации", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            End If
            NamePSTable = ArrayTablesXML(0, 0) 'имя таблицы дял записи OD
            '1. Считываем настройки с шаблона
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim ArrayPropertiesPoint As String(,) = Nothing
            Dim DeskObject As String = FormLoad.ComboBox1.Text
            'если шаблон найден
            If File.Exists(PutchTemplateDwg) = True Then
                'участки постоянного отвода
                If FormLoad.RadioButton10.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Земельные участки образуемые")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Земельные участки образуемые")
                    End If
                    'здания (сооружения)
                ElseIf FormLoad.RadioButton9.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Здания (строения) образуемые")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Здания (строения) образуемые")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Здания (строения) образуемые")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Здания (строения) образуемые")
                    End If
                    'охранные зоны
                ElseIf FormLoad.RadioButton20.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Зоны с особыми условиями использования территорий")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Зоны с особыми условиями использования территорий")
                    End If
                    'линии градостроительного регулирования
                ElseIf FormLoad.RadioButton21.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Линии градостроительного регулирования")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Линии градостроительного регулирования")
                    End If
                ElseIf FormLoad.RadioButton23.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Снос зеленых насаждений")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Снос зеленых насаждений")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Снос зеленых насаждений")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Снос зеленых насаждений")
                    End If
                Else
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Объекты капитального строительства (инженерные коммуникации)")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Объекты капитального строительства (инженерные коммуникации)")
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем разбор массива
            If IsArray(ArrayPropertiesPoint) = True Then
                For i As Integer = 0 To ArrayPropertiesPoint.GetUpperBound(1)
                    Dim NameInsObject As String = ArrayPropertiesPoint(1, i)
                    If ArrayPropertiesPoint(0, i) Like "Name" Then
                        If IsNothing(NameInsObject) = False Then
                            'загружаем блок Автокад
                            If FormLoad.RadioButton11.Checked = True Then
                                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                                If acBlkTbl.Has(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        NameObject = NameInsObject
                                    End If
                                Else
                                    NameObject = NameInsObject
                                End If
                                'загружаем мультивыноску
                            ElseIf FormLoad.RadioButton13.Checked = True Then
                                Dim mlStyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                                If mlStyles.Contains(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        idStyleMleader = mlStyles.GetAt(NameInsObject)
                                    End If
                                Else
                                    idStyleMleader = mlStyles.GetAt(NameInsObject)
                                End If
                                'загружаем метку
                            ElseIf FormLoad.RadioButton16.Checked = True Then
                                Try
                                    idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                Catch ex2 As System.ArgumentException
                                End Try
                                If idStyle.IsNull = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyNoteLabelStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    Try
                                        idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                    Catch ex2 As System.ArgumentException
                                    End Try
                                End If
                                'копируем ключ-описатель
                            ElseIf FormLoad.RadioButton12.Checked = True Then
                                Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db).SearchOrder
                                Dim boolKey As Boolean = False
                                If pdKeySetIds.Count > 0 Then
                                    For Each idKeySet As ObjectId In pdKeySetIds
                                        Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = tr.GetObject(idKeySet, OpenMode.ForRead)
                                        Dim idDesk As ObjectId = Nothing
                                        Try
                                            idDesk = DeskKeySet.Item(NameInsObject)
                                        Catch ex As System.ArgumentException
                                        End Try
                                        If IsNothing(NameInsObject) = False Then
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
                                    Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplateDwg, NameInsObject)
                                End If
                                NameObject = NameInsObject
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "PointGroup" Then
                        'копируем группу точек
                        Dim idGroupPoint As ObjectId = Nothing
                        Try
                            idGroupPoint = CivilDoc.PointGroups.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idGroupPoint) = False Then
                            If idGroupPoint.IsNull = True Then
                                Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                            End If
                        Else
                            Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "MarkerStyle" Then
                        'копируем стиль маркера
                        Try
                            idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idMarkerStyle) = False Then
                            If idMarkerStyle.IsNull = True Then
                                Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                Try
                                    idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                        Else
                            Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                            Try
                                idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                            Catch ex As System.ArgumentException
                            End Try
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Layer" Then
                        Dim LayerinsBlk As String = ArrayPropertiesPoint(1, i)
                        If IsNothing(LayerinsBlk) = False Then
                            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(LayerinsBlk) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, LayerinsBlk)
                                If boolCopyLayer = True Then
                                    LayerObject = LayerinsBlk
                                End If
                            Else
                                LayerObject = LayerinsBlk
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Description" Then
                        DeskPoint = ArrayPropertiesPoint(1, i)
                    ElseIf ArrayPropertiesPoint(0, i) Like "PS" Then
                        NamePsPoint = ArrayPropertiesPoint(1, i)
                        If IsNothing(NamePsPoint) = False Then
                            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                            If dictPropSetDefTemp.Has(NamePsPoint, tr) = False Then
                                Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NamePsPoint)
                            End If
                        End If
                        idTablePs = FuncPropertySet.FuncReturnIDDictSetProperties(NamePsPoint)
                    ElseIf ArrayPropertiesPoint(0, i) Like "Color" Then
                        Dim IndexColor As String = ArrayPropertiesPoint(1, i)
                        If IndexColor Like "ПоСлою" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ПоБлоку" Then
                            ColorObject = 0
                        ElseIf IndexColor Like "ByLayer" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ByBlock" Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) < 1 Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) > 255 Then
                            ColorObject = 256
                        Else
                            ColorObject = Val(IndexColor)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Scale" Then
                        HPoint = Val(ArrayPropertiesPoint(1, i))
                        If HPoint <= 0 Then
                            HPoint = 1
                        End If
                        MashPoint = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
                    ElseIf ArrayPropertiesPoint(0, i) Like "AttributeHeight" Then
                        attributeHeight = Val(ArrayPropertiesPoint(1, i))
                        If attributeHeight <= 0 Then
                            attributeHeight = 2
                        End If
                    End If
                Next i
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'фильтруем объекты
                Dim sourceBtr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                'разрешена проверка дубликатов
                If FormLoad.CheckBox4.Checked = True Then
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
                                If FormLoad.RadioButton11.Checked = True Then 'Блок Автокад
                                    Dim blkPoint As BlockReference = ent
                                    Dim namePointBlk As String = blkPoint.Name
                                    If namePointBlk Like NameObject Then
                                        PointColection.Add(id)
                                    End If
                                End If
                            ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                If FormLoad.RadioButton12.Checked = True Then 'точка COGO
                                    Dim cgPoint As Autodesk.Civil.DatabaseServices.CogoPoint = ent
                                    If cgPoint.RawDescription Like NameObject Then
                                        PointColection.Add(id)
                                    End If
                                End If
                            ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.MLeader Then
                                If FormLoad.RadioButton13.Checked = True Then 'Мультивыноска
                                    Dim MLeaderPoint As MLeader = ent
                                    Dim idStyleName As ObjectId = MLeaderPoint.MLeaderStyle
                                    Dim StyleNameMLeader As MLeaderStyle = tr.GetObject(idStyleMleader, OpenMode.ForRead)
                                    If StyleNameMLeader.Name Like NameObject Then
                                        If MLeaderPoint.PlotStyleName Like NameObject Then
                                            PointColection.Add(id)
                                        End If
                                    End If
                                End If
                            ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.NoteLabel Then
                                If FormLoad.RadioButton16.Checked = True Then 'метка Civil3d
                                    Dim NotePoint As Autodesk.Civil.DatabaseServices.NoteLabel = ent
                                    If NotePoint.StyleName Like NameObject Then
                                        PointColection.Add(id)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                tr.Commit()
            End If
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим наибольший номер
        Dim LargNumber As Integer = 1
        If PointColection.Count > 0 Then
            LargNumber = FuncCadastr.FuncFindLargNumber(PointColection, "")
        End If
        Dim NumbPoint As Integer = LargNumber
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальный номер точки
        If FormLoad.RadioButton15.Checked = True Then
            If FormLoad.RadioButton12.Checked = False Then
                Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер точки: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowNone = True
                pStrOpts.AllowZero = False
                pStrOpts.DefaultValue = LargNumber
                Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts) 'запрашиваем номер
                NumbPoint = pStrRes.Value
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr1 As Transaction = db.TransactionManager.StartTransaction()
            Using tr1
                Dim acBlkTbl As BlockTable = tr1.GetObject(db.BlockTableId, OpenMode.ForRead)
                If acBlkTbl.Has(NameObject) = True Then
                    Try
                        'выбираем примитивы
                        Dim CountParcels As Integer = 0
                        Dim SelectParcel As PromptSelectionResult = ed.GetSelection()
                        If SelectParcel.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                            Dim acSSet1 As SelectionSet = SelectParcel.Value
                            For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                Dim entPLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(IDObj, OpenMode.ForRead, False)
                                If TypeOf entPLine Is Polyline Then
                                    '========================================================================================
                                    'читаем расширенные данные
                                    Dim GLOBID As String = "" 'глобальный id
                                    Dim LocalID As String = "" 'локальный id
                                    Dim Kategory As String = "" 'категория земель
                                    Dim ConditionNumber As String = "" 'номер участка
                                    Dim ArrayData As String(,) = Nothing
                                    Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(entPLine, ArrayData, ArrayTablesXML)
                                    If IsArray(ArrayData) = True Then
                                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                            Dim NTbl As String = ArrayData(0, i)
                                            Dim NField As String = ArrayData(1, i)
                                            Dim Zn As String = ArrayData(2, i)
                                            If NField Like "GLOBALID" Then
                                                GLOBID = Zn
                                            ElseIf NField Like "Category" Then
                                                Kategory = Zn
                                            ElseIf NField Like "LOCALID" Then
                                                LocalID = Zn
                                            ElseIf NField Like "ConditionalNumber" Then
                                                ConditionNumber = Zn
                                            End If
                                        Next i
                                    Else
                                        Continue For
                                    End If
                                    '===============================================================================================
                                    'проверка (проверка на дубликаты по GlobalID)
                                    If GLOBID.Trim.Length > 2 Then
                                        If IsArray(ArrayPoligon) = True Then
                                            For i As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                                                Dim valGlobalId As String = ArrayPoligon(2, i)
                                                If IsNothing(valGlobalId) = True Then valGlobalId = ""
                                                If valGlobalId.Trim Like GLOBID.Trim Then
                                                    GoTo line1
                                                End If
                                            Next i
                                        End If
                                    End If
                                    If NamePSTable Like "PROJECT_PARCEL" Then
                                        '======================================================================================
                                        'категория земель
                                        If FormLoad.CheckBox5.Checked = False Then
                                            Dim pos1 As Integer = -1
                                            If Kategory.IndexOf("ельскохоз") > -1 Then
                                                Delta = 2.5
                                            ElseIf Kategory.IndexOf("промышл") > -1 Then
                                                Delta = 0.5
                                            ElseIf Kategory.IndexOf("водн") > -1 Then
                                                Delta = 2.5
                                            ElseIf Kategory.IndexOf("запас") > -1 Then
                                                Delta = 2.5
                                            ElseIf Kategory.IndexOf("особо") > -1 Then
                                                Delta = 2.5
                                            ElseIf Kategory.IndexOf("лесн") > -1 Then
                                                Delta = 5
                                            ElseIf Kategory.IndexOf("насел") > -1 Then
                                                Delta = 0.1
                                            End If
                                        End If
                                    End If
                                    '=========================================================================================================
                                    '================================================================================================================
                                    'ищем обозначение участка по его локальному ID
                                    If LocalID.Trim.Length > 0 Then
                                        '===============================================================================
                                        'отделяем обозначение участка от его оставшейся части
                                        ConditionNumber = FuncCadastr.FuncFindNumberCadastralObject(ConditionNumber)
                                        'получаем отдельно номер участка для сортировки
                                        Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionNumber)
                                        'получаем обозначение зу без номера контура
                                        Dim FullCondParcel As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionNumber)
                                        If Val(numbParcel) = 0 Then numbParcel = "0"
                                        '=====================================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем все в массив
                                        ReDim Preserve ArrayPoligon(4, CountParcels)
                                        ArrayPoligon(0, CountParcels) = numbParcel 'номер участка
                                        ArrayPoligon(1, CountParcels) = entPLine.Handle.ToString 'хангл полилинии
                                        ArrayPoligon(2, CountParcels) = GLOBID 'глобальный ID
                                        ArrayPoligon(3, CountParcels) = LocalID 'локальный ID
                                        ArrayPoligon(4, CountParcels) = Delta 'погрешность
                                        CountParcels += 1
                                    End If
                                End If
line1:
                            Next
                        End If
                        '==================================================================================================
                        If IsArray(ArrayPoligon) = False Then
                            MsgBox("Массив с участками пуст!!!")
                            Exit Sub
                        Else
                            If ArrayPoligon.GetUpperBound(1) > 0 Then
                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayPoligon, 0)
                            End If
                        End If
                        '====================================================================================================
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'массив с участками
                        For i As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                            'получаем доступ к полилинии для извлечения координат
                            Dim HandlePLine As String = ArrayPoligon(1, i)
                            Dim idPline As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                            Dim PlineObj As Autodesk.AutoCAD.DatabaseServices.Polyline = tr1.GetObject(idPline, OpenMode.ForRead)
                            Dim idLoc As String = ArrayPoligon(3, i) 'локальный ID
                            Dim idGlob As String = ArrayPoligon(2, i) 'глобальный ID
                            '=============================================================================================================
                            'ищем контура
                            Dim ArrayCounter As String(,) = Nothing
                            If idGlob.Trim.Length > 2 Then
                                Dim ArrayCounterTemp As String(,) = Nothing
                                Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(idGlob, ArrayCounterTemp, NamePSTable)
                                'если участок многоконтурный тогда выбираем все контура
                                If IsArray(ArrayCounterTemp) = True Then
                                    If ArrayCounterTemp.GetUpperBound(1) > 0 Then
                                        'это многоконтурный участок ищем обозначения всех контуров
                                        For j As Integer = 0 To ArrayCounterTemp.GetUpperBound(1)
                                            Dim ConditionCounterParcel As String = "0"
                                            Dim LocalTempCounter As String = ArrayCounterTemp(1, j)
                                            Dim hgCounter As String = ArrayCounterTemp(0, j)
                                            Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                                            If IsNothing(idCounter) = False Then
                                                If idCounter.IsNull = False Then
                                                    Dim EntCounter As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(idCounter, OpenMode.ForRead)
                                                    If TypeOf EntCounter Is Polyline Then
                                                        Dim PlineCounter As Polyline = EntCounter
                                                        'ищем расширенные данные
                                                        ConditionCounterParcel = FuncPropertySet.FuncReadValueSetProperties(PlineCounter, "ConditionalNumber", NamePSTable)
                                                    End If
                                                End If
                                            End If
                                            ConditionCounterParcel = FuncCadastr.FuncFindNumberCadastralObject(ConditionCounterParcel)
                                            Dim numbCounter As Integer = FuncCadastr.FuncFindNumberCounterCadastralObject(ConditionCounterParcel)
                                            ReDim Preserve ArrayCounter(2, j)
                                            ArrayCounter(0, j) = numbCounter 'номер контура
                                            ArrayCounter(1, j) = LocalTempCounter 'локальный id
                                            ArrayCounter(2, j) = hgCounter 'handle контура
                                        Next j
                                        'сортируем контура
                                        If ArrayCounter.GetUpperBound(1) > 0 Then
                                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCounter, 0)
                                        End If
                                    Else
                                        'участок у которого есть GLOBALID но найден только один контур
                                        ReDim Preserve ArrayCounter(3, 0)
                                        ArrayCounter(0, 0) = 0
                                        ArrayCounter(1, 0) = idLoc 'локальный ID
                                        ArrayCounter(2, 0) = HandlePLine 'handle контура
                                    End If
                                Else
                                    'участок у которого есть GLOBALID но контуров не найдено
                                    ReDim Preserve ArrayCounter(3, 0)
                                    ArrayCounter(0, 0) = 0
                                    ArrayCounter(1, 0) = idLoc 'локальный ID
                                    ArrayCounter(2, 0) = HandlePLine 'handle контура
                                End If
                            Else
                                ReDim Preserve ArrayCounter(3, 0)
                                ArrayCounter(0, 0) = 0
                                ArrayCounter(1, 0) = idLoc 'локальный ID
                                ArrayCounter(2, 0) = HandlePLine 'handle контура
                            End If
                            '==================================================================================================
                            If IsArray(ArrayCounter) = True Then
                                If ArrayCounter.GetUpperBound(1) > 0 Then
                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCounter, 0)
                                End If
                            End If
                            '================================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'начинаем доставать контуры из массива
                            For j As Integer = 0 To ArrayCounter.GetUpperBound(1)
                                Dim idLocCounter As String = ArrayCounter(1, j)
                                Dim HandleCounter As String = ArrayCounter(2, j)
                                Dim idPlineCounter As ObjectId = AcadFunction.FuncReturnObjectID(HandleCounter)
                                Dim ArrayOstrov1 As String(,) = Nothing 'массив с островками
                                If IsNothing(idPlineCounter) = False Then
                                    If idPlineCounter.IsNull = False Then
                                        Dim PlineObjCounter As Autodesk.AutoCAD.DatabaseServices.Polyline = tr1.GetObject(idPlineCounter, OpenMode.ForWrite)
                                        If ColorObject = 0 Then ColorObject = PlineObjCounter.ColorIndex
                                        If LayerObject = "" Then LayerObject = PlineObjCounter.Layer
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'удаляем совпадающие вершины
                                        Dim boolDblDelete As Boolean = MathFunction.FuncKadastrPlineDelDubleVertex(PlineObjCounter, dopuskXY)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'делаем кругление вершин до см
                                        If FormLoad.CheckBox1.Checked = True Then
                                            Dim boolRound As Boolean = MathFunction.FuncKadastrPlineRoundVertex(PlineObjCounter)
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разворачиваем по часовой стрелке
                                        Dim boolPr As Boolean = False
                                        If PlineObj.Closed = True Then
                                            boolPr = True
                                        Else
                                            Dim ptSt As Point2d = PlineObjCounter.GetPoint2dAt(0)
                                            Dim ptEn As Point2d = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 1)
                                            Dim dist As Double = ptSt.GetVectorTo(ptEn).Length
                                            If dist <= Dopusk Then
                                                PlineObjCounter.RemoveVertexAt(PlineObjCounter.NumberOfVertices - 1)
                                                PlineObjCounter.Closed = True
                                                boolPr = True
                                            End If
                                        End If

                                        Dim clockPline As Boolean = True
                                        If PlineObjCounter.NumberOfVertices > 2 Then
                                            Dim area As Double = MathFunction.GetArea(PlineObjCounter)
                                            If area > 0 Then
                                                clockPline = False
                                            End If
                                            If FormLoad.CheckBox2.Checked = True Then
                                                If area > 0 Then
                                                    Dim boolClockVert As Boolean = MathFunction.FuncKadastrReversPline(PlineObjCounter)
                                                    If boolClockVert = True Then
                                                        clockPline = True
                                                    End If
                                                End If
                                            End If
                                        End If

                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'определяем северозападную вершину
                                        If FormLoad.CheckBox3.Checked = True And boolPr = True Then
                                            Dim boolSZAngle As Boolean = MathFunction.FuncKadastrSVAngle(PlineObjCounter)
                                        End If
                                        boolPr = False
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLocCounter, ArrayOstrov1)
                                        If IsArray(ArrayOstrov1) = True Then
                                            If ArrayOstrov1.GetUpperBound(1) > 0 Then
                                                'сортируем массив
                                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayOstrov1, 0)
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'нумеруем основной полигон
                                        '==================================================================
                                        Dim Vert0 As Point2d = New Point2d(0, 0)
                                        Dim Vert1 As Point2d = New Point2d(0, 0)
                                        Dim Vert2 As Point2d = New Point2d(0, 0)
                                        '=====================================================================================================
                                        'начинаем перебирать вершины полилинии
                                        For i2 As Integer = 0 To PlineObjCounter.NumberOfVertices - 1
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'делаем проверку дубликатов
                                            If FormLoad.CheckBox4.Checked = True Then
                                                Dim idFindPoint As ObjectId = Nothing
                                                'находим ранее созданную точку в этом узле
                                                idFindPoint = FuncCadastr.FuncFindGeoPoint(PointColection, PlineObjCounter.GetPoint2dAt(i2), dopuskXY)
                                                If IsNothing(idFindPoint) = False Then
                                                    If idFindPoint.IsNull = False Then
                                                        Continue For
                                                    End If
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'вычисляем позицию вставки метки точки
                                            If i2 = 0 Then
                                                Vert0 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 1)
                                                Vert1 = PlineObjCounter.GetPoint2dAt(0)
                                                Vert2 = PlineObjCounter.GetPoint2dAt(1)
                                            ElseIf i2 = PlineObjCounter.NumberOfVertices - 1 Then
                                                Vert0 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 2)
                                                Vert1 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 1)
                                                Vert2 = PlineObjCounter.GetPoint2dAt(0)
                                            Else
                                                Vert0 = PlineObjCounter.GetPoint2dAt(i2 - 1)
                                                Vert1 = PlineObjCounter.GetPoint2dAt(i2)
                                                Vert2 = PlineObjCounter.GetPoint2dAt(i2 + 1)
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'массив с хорактеристиками
                                            Dim insAttr As Double() = Nothing
                                            Dim boolInsAttr As Boolean = FuncCadastr.FuncReplaseInsertCoordAttr2(Vert0, Vert1, Vert2, insAttr, DText, clockPline)
                                            Dim point3d1 As Point3d = New Point3d(Vert1.X, Vert1.Y, 0)
                                            Dim NamePointZ As String = prefNumber & NumbPoint
                                            '=================================================================================
                                            'формируем массив для записи
                                            Dim ArrayOBJ1 As String(,) = Nothing
                                            ReDim Preserve ArrayOBJ1(2, 5)
                                            ArrayOBJ1(0, 0) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 0) = "LABEL"
                                            ArrayOBJ1(2, 0) = DeskPoint

                                            ArrayOBJ1(0, 1) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 1) = "Number"
                                            ArrayOBJ1(2, 1) = NamePointZ

                                            ArrayOBJ1(0, 2) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 2) = "IDRNG"
                                            ArrayOBJ1(2, 2) = idLocCounter

                                            ArrayOBJ1(0, 3) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 3) = "Delta"
                                            ArrayOBJ1(2, 3) = Delta

                                            ArrayOBJ1(0, 4) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 4) = "FastDesk"
                                            ArrayOBJ1(2, 4) = Zakrep

                                            ArrayOBJ1(0, 5) = "CADASTRAL_POINT"
                                            ArrayOBJ1(1, 5) = "MethodCoordinates"
                                            ArrayOBJ1(2, 5) = MetodCoord
                                            '=====================================================================================
                                            If FormLoad.RadioButton11.Checked = True Then
                                                'вставляем блок Автокад
                                                '(Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
                                                'Dim blk As BlockReference = AcadFunction.FuncInsertBlk(point3d1, NameObject, LayerObject, ColorObject, (HPoint * M) / 1000)
                                                'If IsNothing(blk) = False Then
                                                'Dim arrayDataAttrib As String(,) = {}
                                                'ReDim arrayDataAttrib(6, 3)
                                                'arrayDataAttrib(0, 0) = "NUMBER"
                                                'arrayDataAttrib(1, 0) = NamePointZ
                                                'arrayDataAttrib(3, 0) = (attributeHeight * M) / 1000
                                                'arrayDataAttrib(4, 0) = rotateAttrib

                                                'arrayDataAttrib(0, 1) = "DELTA"
                                                'arrayDataAttrib(1, 1) = Delta

                                                'arrayDataAttrib(0, 2) = "FASTDESK"
                                                'arrayDataAttrib(1, 2) = Zakrep

                                                'arrayDataAttrib(0, 3) = "METHODCOORDINATES"
                                                'arrayDataAttrib(1, 3) = MetodCoord
                                                'Dim boolWriteAttrai As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, arrayDataAttrib)
                                                Dim idInsertBlk As BlockReference = FuncCadastr.FuncInsertCadastrPoint(point3d1, insAttr, NamePointZ, Delta, Zakrep, HText, LayerObject, NameObject, MashPoint, 0, idStyle, MetodCoord)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'blk.ColorIndex = ColorObject
                                                If FormLoad.CheckBox4.Checked = True Then
                                                    PointColection.Add(idInsertBlk.Id)
                                                End If
                                                '=======================================================================================================
                                                If IsNothing(idInsertBlk) = False Then
                                                    If IsNothing(idTablePs) = False Then
                                                        If idTablePs.IsNull = False Then
                                                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(idInsertBlk, idTablePs)
                                                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(idInsertBlk, ArrayOBJ1)
                                                        End If
                                                    End If
                                                End If
                                                'End If
                                            ElseIf FormLoad.RadioButton12.Checked = True Then
                                                'вставляем точку COGO
                                                Dim InsAttribRez As ObjectId = FuncSurvey.FuncCreateCoGoPoint(point3d1, NameObject)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'записываем расширенные данные в точку
                                                If IsNothing(InsAttribRez) = False Then
                                                    If InsAttribRez.IsNull = False Then
                                                        Dim CoGoObj As Autodesk.Civil.DatabaseServices.CogoPoint = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                        CoGoObj.ColorIndex = ColorObject
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(CoGoObj.Id)
                                                        End If
                                                        If prefNumber.Trim.Length > 0 Then
                                                            Dim boolWriteUDP As Boolean = FuncSurvey.FuncWriteUDPValue(CoGoObj, "CADASTRAL_POINT", "Pref_number", prefNumber)
                                                        End If
                                                        '=======================================================================================================
                                                        If IsNothing(idTablePs) = False Then
                                                            If idTablePs.IsNull = False Then
                                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoObj, idTablePs)
                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoObj, ArrayOBJ1)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf FormLoad.RadioButton13.Checked = True Then
                                                'вставляем мультивыноску
                                                If FormLoad.RadioButton15.Checked = True Then
                                                    Dim offsetx As Double = insAttr(0) - point3d1.X
                                                    Dim offsety As Double = insAttr(1) - point3d1.Y
                                                    Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), NamePointZ, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные в точку
                                                    If IsNothing(InsAttribRez) = False Then
                                                        If InsAttribRez.IsNull = False Then
                                                            Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                            MleaderObj.ColorIndex = ColorObject
                                                            MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                            If FormLoad.CheckBox4.Checked = True Then
                                                                PointColection.Add(MleaderObj.Id)
                                                            End If
                                                            '=======================================================================================================
                                                            If IsNothing(idTablePs) = False Then
                                                                If idTablePs.IsNull = False Then
                                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MleaderObj, idTablePs)
                                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MleaderObj, ArrayOBJ1)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    Dim x As String = FuncFormatZN.FuncFormatCoordinate(point3d1.Y, 2)
                                                    Dim y As String = FuncFormatZN.FuncFormatCoordinate(point3d1.X, 2)
                                                    Dim TextMleader As String = "X=" & x & vbLf & "Y=" & y
                                                    Dim offsetx As Double = insAttr(0) - point3d1.X
                                                    Dim offsety As Double = insAttr(1) - point3d1.Y
                                                    Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), TextMleader, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                    If IsNothing(InsAttribRez) = False Then
                                                        If InsAttribRez.IsNull = False Then
                                                            Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                            MleaderObj.ColorIndex = ColorObject
                                                            MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                            If FormLoad.CheckBox4.Checked = True Then
                                                                PointColection.Add(InsAttribRez)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf FormLoad.RadioButton16.Checked = True Then
                                                'вставляем метку Civil3D
                                                If FormLoad.RadioButton15.Checked = True Then

                                                Else
                                                    Dim IDlb As ObjectId = Autodesk.Civil.DatabaseServices.NoteLabel.Create(point3d1, idLabelStyle, idMarkerStyle)
                                                    If FormLoad.CheckBox4.Checked = True Then
                                                        PointColection.Add(IDlb)
                                                    End If
                                                    Dim lb As Autodesk.Civil.DatabaseServices.NoteLabel = tr1.GetObject(IDlb, OpenMode.ForWrite)
                                                    lb.Layer = LayerObject
                                                    lb.ColorIndex = ColorObject
                                                    lb.LabelLocation = New Point3d(insAttr(0), insAttr(1), 0)
                                                End If
                                            End If
                                            NumbPoint += 1
                                        Next i2
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'нумеруем островки
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                If IsArray(ArrayOstrov1) = True Then
                                    For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                        Dim idLocOstrov As String = ArrayOstrov1(3, k)
                                        Dim hgOstrov As String = ArrayOstrov1(1, k)
                                        Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                        If IsNothing(idOstrov) = False Then
                                            If idOstrov.IsNull = False Then
                                                Dim PlineObjOstrov As Polyline = tr1.GetObject(idOstrov, OpenMode.ForWrite)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'удаляем совпадающие вершины
                                                Dim boolDblDelete As Boolean = MathFunction.FuncKadastrPlineDelDubleVertex(PlineObjOstrov, dopuskXY)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'делаем кругление вершин до см
                                                If FormLoad.CheckBox1.Checked = True Then
                                                    Dim boolRound As Boolean = MathFunction.FuncKadastrPlineRoundVertex(PlineObjOstrov)
                                                End If
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'разворачиваем по часовой стрелке
                                                Dim boolPr As Boolean = False
                                                If PlineObj.Closed = True Then
                                                    boolPr = True
                                                Else
                                                    Dim ptSt As Point2d = PlineObjOstrov.GetPoint2dAt(0)
                                                    Dim ptEn As Point2d = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 1)
                                                    Dim dist As Double = ptSt.GetVectorTo(ptEn).Length
                                                    If dist <= Dopusk Then
                                                        PlineObjOstrov.RemoveVertexAt(PlineObjOstrov.NumberOfVertices - 1)
                                                        PlineObjOstrov.Closed = True
                                                        boolPr = True
                                                    End If
                                                End If

                                                Dim clockPline As Boolean = True
                                                If PlineObjOstrov.NumberOfVertices > 2 Then
                                                    Dim area As Double = MathFunction.GetArea(PlineObjOstrov)
                                                    If area > 0 Then
                                                        clockPline = False
                                                    End If
                                                    If FormLoad.CheckBox2.Checked = True Then
                                                        If area > 0 Then
                                                            Dim boolClockVert As Boolean = MathFunction.FuncKadastrReversPline(PlineObjOstrov)
                                                            If boolClockVert = True Then
                                                                clockPline = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'определяем северозападную вершину
                                                If FormLoad.CheckBox3.Checked = True Then
                                                    Dim boolSZAngle As Boolean = MathFunction.FuncKadastrSVAngle(PlineObjOstrov)
                                                End If
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                Dim Vert0 As Point2d = New Point2d(0, 0)
                                                Dim Vert1 As Point2d = New Point2d(0, 0)
                                                Dim Vert2 As Point2d = New Point2d(0, 0)
                                                For i2 As Integer = 0 To PlineObjOstrov.NumberOfVertices - 1
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'делаем проверку дубликатов
                                                    If FormLoad.CheckBox4.Checked = True Then
                                                        Dim idFindPoint As ObjectId = Nothing
                                                        'находим номер точки
                                                        idFindPoint = FuncCadastr.FuncFindGeoPoint(PointColection, PlineObjOstrov.GetPoint2dAt(i2), dopuskXY)
                                                        If IsNothing(idFindPoint) = False Then
                                                            If idFindPoint.IsNull = False Then
                                                                Continue For
                                                            End If
                                                        End If
                                                    End If
                                                    If i2 = 0 Then
                                                        Vert0 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 1)
                                                        Vert1 = PlineObjOstrov.GetPoint2dAt(0)
                                                        Vert2 = PlineObjOstrov.GetPoint2dAt(1)
                                                    ElseIf i2 = PlineObjOstrov.NumberOfVertices - 1 Then
                                                        Vert0 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 2)
                                                        Vert1 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 1)
                                                        Vert2 = PlineObjOstrov.GetPoint2dAt(0)
                                                    Else
                                                        Vert0 = PlineObjOstrov.GetPoint2dAt(i2 - 1)
                                                        Vert1 = PlineObjOstrov.GetPoint2dAt(i2)
                                                        Vert2 = PlineObjOstrov.GetPoint2dAt(i2 + 1)
                                                    End If
                                                    Dim insAttr As Double() = Nothing 'массив с хорактеристиками
                                                    Dim boolInsAttr As Boolean = FuncCadastr.FuncReplaseInsertCoordAttr2(Vert0, Vert1, Vert2, insAttr, DText, clockPline)
                                                    Dim point3d1 As Point3d = New Point3d(Vert1.X, Vert1.Y, 0)
                                                    Dim NamePointZ As String = prefNumber & NumbPoint
                                                    '=================================================================================
                                                    'формируем массив для записи
                                                    Dim ArrayOBJ1 As String(,) = Nothing
                                                    ReDim Preserve ArrayOBJ1(2, 5)
                                                    ArrayOBJ1(0, 0) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 0) = "LABEL"
                                                    ArrayOBJ1(2, 0) = DeskPoint

                                                    ArrayOBJ1(0, 1) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 1) = "Number"
                                                    ArrayOBJ1(2, 1) = NamePointZ

                                                    ArrayOBJ1(0, 2) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 2) = "IDRNG"
                                                    ArrayOBJ1(2, 2) = idLocCounter

                                                    ArrayOBJ1(0, 3) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 3) = "Delta"
                                                    ArrayOBJ1(2, 3) = Delta

                                                    ArrayOBJ1(0, 4) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 4) = "FastDesk"
                                                    ArrayOBJ1(2, 4) = Zakrep

                                                    ArrayOBJ1(0, 5) = "CADASTRAL_POINT"
                                                    ArrayOBJ1(1, 5) = "MethodCoordinates"
                                                    ArrayOBJ1(2, 5) = MetodCoord
                                                    '=====================================================================================
                                                    If FormLoad.RadioButton11.Checked = True Then
                                                        'вставляем блок Автокад
                                                        'Dim blk As BlockReference = AcadFunction.FuncInsertBlk(point3d1, NameObject, LayerObject, ColorObject, (HPoint * M) / 1000)
                                                        'If IsNothing(blk) = False Then
                                                        '(Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
                                                        'Dim arrayDataAttrib As String(,) = {}
                                                        'ReDim arrayDataAttrib(6, 3)
                                                        'arrayDataAttrib(0, 0) = "NUMBER"
                                                        'arrayDataAttrib(1, 0) = NamePointZ
                                                        'arrayDataAttrib(3, 0) = (attributeHeight * M) / 1000
                                                        'arrayDataAttrib(4, 0) = rotateAttrib

                                                        'arrayDataAttrib(0, 1) = "DELTA"
                                                        'arrayDataAttrib(1, 1) = Delta

                                                        'arrayDataAttrib(0, 2) = "FASTDESK"
                                                        'arrayDataAttrib(1, 2) = Zakrep

                                                        'arrayDataAttrib(0, 3) = "METHODCOORDINATES"
                                                        'arrayDataAttrib(1, 3) = MetodCoord
                                                        'Dim boolWriteAttrai As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, arrayDataAttrib)
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем данные в точку
                                                        Dim idInsertBlk As BlockReference = FuncCadastr.FuncInsertCadastrPoint(point3d1, insAttr, NamePointZ, Delta, Zakrep, HText, LayerObject, NameObject, MashPoint, 0, idStyle, MetodCoord)
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'blk.ColorIndex = ColorObject
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(idInsertBlk.Id)
                                                        End If
                                                        '=======================================================================================================
                                                        If IsNothing(idInsertBlk) = False Then
                                                            If IsNothing(idTablePs) = False Then
                                                                If idTablePs.IsNull = False Then
                                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(idInsertBlk, idTablePs)
                                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(idInsertBlk, ArrayOBJ1)
                                                                End If
                                                            End If
                                                        End If
                                                        'End If
                                                    ElseIf FormLoad.RadioButton12.Checked = True Then
                                                        'вставляем точку COGO
                                                        Dim InsAttribRez As ObjectId = FuncSurvey.FuncCreateCoGoPoint(point3d1, NameObject)
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем расширенные данные в точку
                                                        If IsNothing(InsAttribRez) = False Then
                                                            If InsAttribRez.IsNull = False Then
                                                                Dim CoGoObj As Autodesk.Civil.DatabaseServices.CogoPoint = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                                CoGoObj.ColorIndex = ColorObject
                                                                If FormLoad.CheckBox4.Checked = True Then
                                                                    PointColection.Add(CoGoObj.Id)
                                                                End If
                                                                If prefNumber.Trim.Length > 0 Then
                                                                    Dim boolWriteUDP As Boolean = FuncSurvey.FuncWriteUDPValue(CoGoObj, "CADASTRAL_POINT", "Pref_number", prefNumber)
                                                                End If
                                                                '=======================================================================================================
                                                                If IsNothing(idTablePs) = False Then
                                                                    If idTablePs.IsNull = False Then
                                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoObj, idTablePs)
                                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoObj, ArrayOBJ1)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    ElseIf FormLoad.RadioButton13.Checked = True Then
                                                        'вставляем мультивыноску
                                                        If FormLoad.RadioButton15.Checked = True Then
                                                            Dim offsetx As Double = insAttr(0) - point3d1.X
                                                            Dim offsety As Double = insAttr(1) - point3d1.Y
                                                            Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), NamePointZ, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                            'записываем расширенные данные в точку
                                                            If IsNothing(InsAttribRez) = False Then
                                                                If InsAttribRez.IsNull = False Then
                                                                    Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                                    MleaderObj.ColorIndex = ColorObject
                                                                    MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                                    If FormLoad.CheckBox4.Checked = True Then
                                                                        PointColection.Add(MleaderObj.Id)
                                                                    End If
                                                                    '=======================================================================================================
                                                                    If IsNothing(idTablePs) = False Then
                                                                        If idTablePs.IsNull = False Then
                                                                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MleaderObj, idTablePs)
                                                                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MleaderObj, ArrayOBJ1)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            Dim x As String = FuncFormatZN.FuncFormatCoordinate(point3d1.Y, 2)
                                                            Dim y As String = FuncFormatZN.FuncFormatCoordinate(point3d1.X, 2)
                                                            Dim TextMleader As String = "X=" & x & vbLf & "Y=" & y
                                                            Dim offsetx As Double = insAttr(0) - point3d1.X
                                                            Dim offsety As Double = insAttr(1) - point3d1.Y
                                                            Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), TextMleader, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                            If IsNothing(InsAttribRez) = False Then
                                                                If InsAttribRez.IsNull = False Then
                                                                    Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                                    MleaderObj.ColorIndex = ColorObject
                                                                    MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                                    If FormLoad.CheckBox4.Checked = True Then
                                                                        PointColection.Add(InsAttribRez)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    ElseIf FormLoad.RadioButton16.Checked = True Then
                                                        'вставляем метку Civil3D
                                                        If FormLoad.RadioButton15.Checked = True Then

                                                        Else
                                                            Dim IDlb As ObjectId = Autodesk.Civil.DatabaseServices.NoteLabel.Create(point3d1, idLabelStyle, idMarkerStyle)
                                                            If FormLoad.CheckBox4.Checked = True Then
                                                                PointColection.Add(IDlb)
                                                            End If
                                                            Dim lb As Autodesk.Civil.DatabaseServices.NoteLabel = tr1.GetObject(IDlb, OpenMode.ForWrite)
                                                            lb.Layer = LayerObject
                                                            lb.ColorIndex = ColorObject
                                                            lb.LabelLocation = New Point3d(insAttr(0), insAttr(1), 0)
                                                        End If
                                                    End If
                                                    NumbPoint += 1
                                                Next i2
                                            End If
                                        End If
                                    Next k
                                End If
                            Next j
                        Next i
                    Catch
                    End Try
                End If
                tr1.Commit()
            End Using
        End Using
    End Sub
    '==============================================================================================
    'перенумеровать полигоны
    <CommandMethod("ReNumberUserPolygons")>
    Public Sub ReNumberUserPolygons()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormLoad As FormCreatePoint = New FormCreatePoint
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        If MAnno = 10000 Then
            FormLoad.RadioButton1.Checked = True
            FormLoad.NumericUpDown1.Value = 10000
        ElseIf MAnno = 7000 Then
            FormLoad.RadioButton2.Checked = True
            FormLoad.NumericUpDown1.Value = 7000
        ElseIf MAnno = 5000 Then
            FormLoad.RadioButton3.Checked = True
            FormLoad.NumericUpDown1.Value = 5000
        ElseIf MAnno = 3000 Then
            FormLoad.RadioButton18.Checked = True
            FormLoad.NumericUpDown1.Value = 3000
        ElseIf MAnno = 2000 Then
            FormLoad.RadioButton4.Checked = True
            FormLoad.NumericUpDown1.Value = 2000
        ElseIf MAnno = 1000 Then
            FormLoad.RadioButton5.Checked = True
            FormLoad.NumericUpDown1.Value = 1000
        ElseIf MAnno = 500 Then
            FormLoad.RadioButton6.Checked = True
            FormLoad.NumericUpDown1.Value = 500
        ElseIf MAnno = 200 Then
            FormLoad.RadioButton17.Checked = True
            FormLoad.NumericUpDown1.Value = 200
        ElseIf MAnno = 100 Then
            FormLoad.RadioButton7.Checked = True
            FormLoad.NumericUpDown1.Value = 100
        ElseIf MAnno = 50 Then
            FormLoad.RadioButton14.Checked = True
            FormLoad.NumericUpDown1.Value = 50
        Else
            FormLoad.RadioButton8.Checked = True
            FormLoad.NumericUpDown1.Value = MAnno
        End If
        'вызываем диалоговое окно УСТАНОВКИ
        FormLoad.ShowDialog()
        If FormLoad.UstanovkiLoad = 0 Then
            Exit Sub
        End If
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
        Dim PutchTemplateDwg As String = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
        'ищем шаблон по ближе
        Try
            Dim TempDir As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                PutchTemplateDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
                If IO.File.Exists(PutchTemplateDir & FormLoad.ComboBox3.Text.Trim) = True Then
                    PutchTemplateDwg = PutchTemplateDir & FormLoad.ComboBox3.Text.Trim
                End If
            End If
        Catch ex As System.Security.SecurityException
        End Try
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        Dim idStyle As ObjectId = Nothing 'текстовые стиль
        'префикс номера
        Dim prefNumber As String = FormLoad.ComboBox4.Text
        'метод определения координат
        Dim MetodCoord As String = FormLoad.ComboBox5.Text
        '========================================================================================
        Dim M As Double = FormLoad.NumericUpDown1.Value 'масштаб
        Dim HText As Double = (M * FormLoad.NumericUpDown2.Value) / 1000 ' высота подписи
        Dim DText As Double = (M * FormLoad.NumericUpDown3.Value) / 1000 'расстояние смещения
        '==================================================================================================
        Dim NameObject As String = ""
        Dim LayerObject As String = ""
        Dim NamePsPoint As String = ""
        Dim ColorObject As String = 256
        Dim HPoint As Double = 1 'размер точки
        Dim DeskPoint As String = ""
        Dim attributeHeight As Double = 2
        Dim rotateAttrib As Double = 0
        Dim Delta As String = FormLoad.NumericUpDown6.Value
        Dim Zakrep As String = FormLoad.ComboBox2.Text
        Dim Dopusk As Double = FormLoad.NumericUpDown7.Value
        Dim idTablePs As ObjectId = Nothing
        Dim NamePSTable As String = ""
        Dim MashPoint As Scale3d = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayCounterNumber As String(,) = Nothing 'массив для нумерации всех контуров
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        Dim PointColection As ObjectIdCollection = New ObjectIdCollection()
        Dim idStyleMleader As ObjectId = Nothing
        Dim idMarkerStyle As ObjectId = Nothing
        Dim idLabelStyle As ObjectId = Nothing
        Dim ArrayOldPoint As String(,) = Nothing
        Dim CountArrayOldPoint As Integer = 0
        Dim dopuskXY As Double = FormLoad.NumericUpDown7.Value
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '2.массив лоя мультивыносок
        Dim ArrayMLeader As String(,) = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем транзакцию
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'создаем текстовый стиль для точек
            Dim st As TextStyleTable = CType(tr.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
            If st.Has("CADASTRAL_POINT") = False Then
                Dim crTextstyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, "CADASTRAL_POINT", False)
                If crTextstyle = True Then
                    If st.Has("CADASTRAL_POINT") = True Then
                        idStyle = st.Item("CADASTRAL_POINT")
                    End If
                End If
                If IsNothing(idStyle) = True Then
                    idStyle = db.Textstyle
                End If
                If idStyle.IsNull = True Then
                    idStyle = db.Textstyle
                End If
            Else
                idStyle = st.Item("CADASTRAL_POINT")
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'таблицы xml
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ReDim ArrayTablesXML(1, 0)
            If FormLoad.RadioButton10.Checked = True Then
                If FormLoad.ComboBox1.Text.IndexOf("сервит") > -1 Then
                    ArrayTablesXML(0, 0) = "TEMPORARY_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок временного отвода"
                    Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
                Else
                    ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
                    ArrayTablesXML(1, 0) = "Участок постоянного отвода"
                    Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
                    Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
                End If
            ElseIf FormLoad.RadioButton1.Checked = True Then
                ArrayTablesXML(0, 0) = "PROJECT_BUILDING"
                ArrayTablesXML(1, 0) = "Здание или сооружение проектируемое"
            ElseIf FormLoad.RadioButton20.Checked = True Then
                ArrayTablesXML(0, 0) = "PROJECT_ZONA"
                ArrayTablesXML(1, 0) = "Зона с особыми условиями"
                Dim NameStyleMLeader As String() = {"ППТ-Территориальные зоны", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton21.Checked = True Then
                ArrayTablesXML(0, 0) = "URBAN_LINE"
                ArrayTablesXML(1, 0) = "Линии градостроительного регулирования"
                Dim NameStyleMLeader As String() = {"ППТ-Линии градостроительного регулирования", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            ElseIf FormLoad.RadioButton23.Checked = True Then
                ReDim ArrayTablesXML(1, 0)
                ArrayTablesXML(0, 0) = "PROJECT_TREE"
                ArrayTablesXML(1, 0) = "Снос зеленых насаждений"
                Dim NameStyleMLeader As String() = {"ППТ-Граница контура", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            Else
                ArrayTablesXML(0, 0) = "PROJECT_LINE"
                ArrayTablesXML(1, 0) = "Сооружение линейного типа"
                Dim NameStyleMLeader As String() = {"ППТ-Инженерные коммуникации", "CADASTRAL_OBJECT"}
                Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            End If
            NamePSTable = ArrayTablesXML(0, 0) 'имя таблицы дял записи OD
            '1. Считываем настройки с шаблона
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim ArrayPropertiesPoint As String(,) = Nothing
            Dim DeskObject As String = FormLoad.ComboBox1.Text
            'если шаблон найден
            If File.Exists(PutchTemplateDwg) = True Then
                'участки постоянного отвода
                If FormLoad.RadioButton10.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Земельные участки образуемые")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Земельные участки образуемые")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Земельные участки образуемые")
                    End If
                    'участки временного отвода
                ElseIf FormLoad.RadioButton22.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Здания (сооружения)")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Здания (сооружения)")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Здания (сооружения)")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Здания (сооружения)")
                    End If
                    'охранные зоны
                ElseIf FormLoad.RadioButton20.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Зоны с особыми условиями использования территорий")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Зоны с особыми условиями использования территорий")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Зоны с особыми условиями использования территорий")
                    End If
                    'линии градостроительного регулирования
                ElseIf FormLoad.RadioButton21.Checked = True Then
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Линии градостроительного регулирования")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Линии градостроительного регулирования")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Линии градостроительного регулирования")
                    End If
                Else
                    If FormLoad.RadioButton11.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Блок", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton12.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "ТочкаCOGO", "Объекты капитального строительства (инженерные коммуникации)")
                    ElseIf FormLoad.RadioButton13.Checked = True Then
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Мультивыноска", "Объекты капитального строительства (инженерные коммуникации)")
                    Else
                        Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskObject, ArrayPropertiesPoint, "Метка", "Объекты капитального строительства (инженерные коммуникации)")
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем разбор массива
            If IsArray(ArrayPropertiesPoint) = True Then
                For i As Integer = 0 To ArrayPropertiesPoint.GetUpperBound(1)
                    Dim NameInsObject As String = ArrayPropertiesPoint(1, i)
                    If ArrayPropertiesPoint(0, i) Like "Name" Then
                        If IsNothing(NameInsObject) = False Then
                            'загружаем блок Автокад
                            If FormLoad.RadioButton11.Checked = True Then
                                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                                If acBlkTbl.Has(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        NameObject = NameInsObject
                                    End If
                                Else
                                    NameObject = NameInsObject
                                End If
                                'загружаем мультивыноску
                            ElseIf FormLoad.RadioButton13.Checked = True Then
                                Dim mlStyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                                If mlStyles.Contains(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        idStyleMleader = mlStyles.GetAt(NameInsObject)
                                    End If
                                Else
                                    idStyleMleader = mlStyles.GetAt(NameInsObject)
                                End If
                                'загружаем метку
                            ElseIf FormLoad.RadioButton16.Checked = True Then
                                Try
                                    idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                Catch ex2 As System.ArgumentException
                                End Try
                                If idStyle.IsNull = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCopyNoteLabelStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    Try
                                        idLabelStyle = CivilDoc.Styles.LabelStyles.GeneralNoteLabelStyles.Item(NameInsObject)
                                    Catch ex2 As System.ArgumentException
                                    End Try
                                End If
                                'копируем ключ-описатель
                            ElseIf FormLoad.RadioButton12.Checked = True Then
                                Dim pdKeySetIds As ObjectIdCollection = Autodesk.Civil.DatabaseServices.PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db).SearchOrder
                                Dim boolKey As Boolean = False
                                If pdKeySetIds.Count > 0 Then
                                    For Each idKeySet As ObjectId In pdKeySetIds
                                        Dim DeskKeySet As Autodesk.Civil.DatabaseServices.PointDescriptionKeySet = tr.GetObject(idKeySet, OpenMode.ForRead)
                                        Dim idDesk As ObjectId = Nothing
                                        Try
                                            idDesk = DeskKeySet.Item(NameInsObject)
                                        Catch ex As System.ArgumentException
                                        End Try
                                        If IsNothing(NameInsObject) = False Then
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
                                    Dim boolCreatePG As Boolean = AcadFunction.FuncCopyDeskKeyDestDraw(PutchTemplateDwg, NameInsObject)
                                End If
                                NameObject = NameInsObject
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "PointGroup" Then
                        'копируем группу точек
                        Dim idGroupPoint As ObjectId = Nothing
                        Try
                            idGroupPoint = CivilDoc.PointGroups.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idGroupPoint) = False Then
                            If idGroupPoint.IsNull = True Then
                                Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                            End If
                        Else
                            Dim boolCreatePG As Boolean = AcadFunction.FuncCopyGroupPointsDestDraw(PutchTemplateDwg, NameInsObject)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "MarkerStyle" Then
                        'копируем стиль маркера
                        Try
                            idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                        Catch ex As System.ArgumentException
                        End Try
                        If IsNothing(idMarkerStyle) = False Then
                            If idMarkerStyle.IsNull = True Then
                                Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                Try
                                    idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                        Else
                            Dim boolCreateLabelLineStyle As Boolean = AcadFunction.FuncCopyMarkerStyleDestDraw(PutchTemplateDwg, NameInsObject)
                            Try
                                idMarkerStyle = CivilDoc.Styles.MarkerStyles.Item(NameInsObject)
                            Catch ex As System.ArgumentException
                            End Try
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Layer" Then
                        Dim LayerinsBlk As String = ArrayPropertiesPoint(1, i)
                        If IsNothing(LayerinsBlk) = False Then
                            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                            If acLyrTbl.Has(LayerinsBlk) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, LayerinsBlk)
                                If boolCopyLayer = True Then
                                    LayerObject = LayerinsBlk
                                End If
                            Else
                                LayerObject = LayerinsBlk
                            End If
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Description" Then
                        DeskPoint = ArrayPropertiesPoint(1, i)
                    ElseIf ArrayPropertiesPoint(0, i) Like "PS" Then
                        NamePsPoint = ArrayPropertiesPoint(1, i)
                        If IsNothing(NamePsPoint) = False Then
                            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                            If dictPropSetDefTemp.Has(NamePsPoint, tr) = False Then
                                Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NamePsPoint)
                            End If
                        End If
                        idTablePs = FuncPropertySet.FuncReturnIDDictSetProperties(NamePsPoint)
                    ElseIf ArrayPropertiesPoint(0, i) Like "Color" Then
                        Dim IndexColor As String = ArrayPropertiesPoint(1, i)
                        If IndexColor Like "ПоСлою" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ПоБлоку" Then
                            ColorObject = 0
                        ElseIf IndexColor Like "ByLayer" Then
                            ColorObject = 256
                        ElseIf IndexColor Like "ByBlock" Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) < 1 Then
                            ColorObject = 0
                        ElseIf Val(IndexColor) > 255 Then
                            ColorObject = 256
                        Else
                            ColorObject = Val(IndexColor)
                        End If
                    ElseIf ArrayPropertiesPoint(0, i) Like "Scale" Then
                        HPoint = Val(ArrayPropertiesPoint(1, i))
                        If HPoint <= 0 Then
                            HPoint = 1
                        End If
                        MashPoint = New Scale3d((HPoint * M) / 1000, (HPoint * M) / 1000, (HPoint * M) / 1000)
                    ElseIf ArrayPropertiesPoint(0, i) Like "AttributeHeight" Then
                        attributeHeight = Val(ArrayPropertiesPoint(1, i))
                        If attributeHeight <= 0 Then
                            attributeHeight = 2
                        End If
                    End If
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'фильтруем блоки вершин 
            If FormLoad.RadioButton15.Checked = True Then 'выводим номера точек
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
                            If FormLoad.RadioButton11.Checked = True Then 'Блок Автокад
                                Dim blkPoint As BlockReference = ent
                                Dim namePointBlk As String = blkPoint.Name
                                If namePointBlk Like NameObject Then
                                    ReDim Preserve ArrayOldPoint(4, CountArrayOldPoint)
                                    ArrayOldPoint(0, CountArrayOldPoint) = ""
                                    ArrayOldPoint(1, CountArrayOldPoint) = blkPoint.Position.X
                                    ArrayOldPoint(2, CountArrayOldPoint) = blkPoint.Position.Y
                                    ArrayOldPoint(3, CountArrayOldPoint) = blkPoint.Handle.ToString
                                    ArrayOldPoint(4, CountArrayOldPoint) = "0"
                                    CountArrayOldPoint += 1
                                End If
                            End If
                        ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                            If FormLoad.RadioButton12.Checked = True Then 'точка COGO
                                Dim cgPoint As Autodesk.Civil.DatabaseServices.CogoPoint = ent
                                If cgPoint.RawDescription Like NameObject Then
                                    ReDim Preserve ArrayOldPoint(4, CountArrayOldPoint)
                                    ArrayOldPoint(0, CountArrayOldPoint) = cgPoint.PointNumber
                                    ArrayOldPoint(1, CountArrayOldPoint) = cgPoint.Easting
                                    ArrayOldPoint(2, CountArrayOldPoint) = cgPoint.Northing
                                    ArrayOldPoint(3, CountArrayOldPoint) = cgPoint.Handle.ToString
                                    ArrayOldPoint(4, CountArrayOldPoint) = "0"
                                    CountArrayOldPoint += 1
                                End If
                            End If
                        ElseIf TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.MLeader Then
                            If FormLoad.RadioButton13.Checked = True Then 'Мультивыноска
                                Dim MLeaderPoint As MLeader = ent
                                Dim idStyleName As ObjectId = MLeaderPoint.MLeaderStyle
                                Dim StyleNameMLeader As MLeaderStyle = tr.GetObject(idStyleMleader, OpenMode.ForRead)
                                If StyleNameMLeader.Name Like NameObject Then
                                    ArrayOldPoint(0, CountArrayOldPoint) = ""
                                    ArrayOldPoint(1, CountArrayOldPoint) = MLeaderPoint.GetFirstVertex(0).X
                                    ArrayOldPoint(2, CountArrayOldPoint) = MLeaderPoint.GetFirstVertex(0).Y
                                    ArrayOldPoint(3, CountArrayOldPoint) = MLeaderPoint.Handle.ToString
                                    ArrayOldPoint(4, CountArrayOldPoint) = "0"
                                    CountArrayOldPoint += 1
                                End If
                            End If
                        ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.NoteLabel Then
                            If FormLoad.RadioButton16.Checked = True Then 'метка Civil3d
                                Dim NotePoint As Autodesk.Civil.DatabaseServices.NoteLabel = ent
                                If NotePoint.StyleName Like NameObject Then
                                    ArrayOldPoint(0, CountArrayOldPoint) = ""
                                    ArrayOldPoint(1, CountArrayOldPoint) = NotePoint.LabelLocation.X
                                    ArrayOldPoint(2, CountArrayOldPoint) = NotePoint.LabelLocation.Y
                                    ArrayOldPoint(3, CountArrayOldPoint) = NotePoint.Handle.ToString
                                    ArrayOldPoint(4, CountArrayOldPoint) = "0"
                                    CountArrayOldPoint += 1
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            tr.Commit()
        End Using
        If IsArray(ArrayOldPoint) = False Then
            MsgBox("Фильтр не нашел ни одной существующей точки!!! Макро прервано.")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim NumbPoint As Integer = 1
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальный номер точки
        If FormLoad.RadioButton15.Checked = True Then
            If FormLoad.RadioButton12.Checked = False Then
                Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер точки: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowNone = True
                pStrOpts.AllowZero = False
                pStrOpts.DefaultValue = NumbPoint
                Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts) 'запрашиваем номер
                NumbPoint = pStrRes.Value
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr1 As Transaction = db.TransactionManager.StartTransaction()
            Using tr1
                Try
                    'выбираем примитивы
                    Dim CountParcels As Integer = 0
                    Dim SelectParcel As PromptSelectionResult = ed.GetSelection()
                    If SelectParcel.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                        Dim acSSet1 As SelectionSet = SelectParcel.Value
                        For Each IDObj As ObjectId In acSSet1.GetObjectIds
                            Dim entPLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(IDObj, OpenMode.ForRead, False)
                            If TypeOf entPLine Is Polyline Then
                                '========================================================================================
                                'читаем расширенные данные
                                Dim GLOBID As String = "" 'глобальный id
                                Dim LocalID As String = "" 'локальный id
                                Dim Kategory As String = "" 'категория земель
                                Dim ConditionNumber As String = "" 'номер участка
                                Dim ArrayData As String(,) = Nothing
                                Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(entPLine, ArrayData, ArrayTablesXML)
                                If IsArray(ArrayData) = True Then
                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Dim NTbl As String = ArrayData(0, i)
                                        Dim NField As String = ArrayData(1, i)
                                        Dim Zn As String = ArrayData(2, i)
                                        If NField Like "GLOBALID" Then
                                            GLOBID = Zn
                                        ElseIf NField Like "Category" Then
                                            Kategory = Zn
                                        ElseIf NField Like "LOCALID" Then
                                            LocalID = Zn
                                        ElseIf NField Like "ConditionalNumber" Then
                                            ConditionNumber = Zn
                                        End If
                                    Next i
                                Else
                                    Continue For
                                End If
                                '===============================================================================================
                                'проверка (проверка на дубликаты по Glibalid)
                                If GLOBID.Trim.Length > 2 Then
                                    If IsArray(ArrayPoligon) = True Then
                                        For i As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                                            Dim valGlobalId As String = ArrayPoligon(2, i)
                                            If IsNothing(valGlobalId) = True Then valGlobalId = ""
                                            If valGlobalId.Trim Like GLOBID.Trim Then
                                                GoTo line1
                                            End If
                                        Next i
                                    End If
                                End If

                                If NamePSTable Like "PROJECT_PARCEL" Or NamePSTable Like "TEMPORARY_PARCEL" Then
                                    '======================================================================================
                                    'категория земель
                                    If FormLoad.CheckBox5.Checked = False Then
                                        Dim pos1 As Integer = -1
                                        If Kategory.IndexOf("ельскохоз") > -1 Then
                                            Delta = 2.5
                                        ElseIf Kategory.IndexOf("промышл") > -1 Then
                                            Delta = 0.5
                                        ElseIf Kategory.IndexOf("водн") > -1 Then
                                            Delta = 2.5
                                        ElseIf Kategory.IndexOf("запас") > -1 Then
                                            Delta = 2.5
                                        ElseIf Kategory.IndexOf("особо") > -1 Then
                                            Delta = 2.5
                                        ElseIf Kategory.IndexOf("лесн") > -1 Then
                                            Delta = 5
                                        ElseIf Kategory.IndexOf("насел") > -1 Then
                                            Delta = 0.1
                                        End If
                                    End If
                                End If
                                '=========================================================================================================
                                '================================================================================================================
                                'ищем обозначение участка по его локальному ID
                                If LocalID.Trim.Length > 0 Then
                                    '===============================================================================
                                    'отделяем обозначение участка от его оставшейся части
                                    ConditionNumber = FuncCadastr.FuncFindNumberCadastralObject(ConditionNumber)
                                    'получаем отдельно номер участка для сортировки
                                    Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionNumber)
                                    'получаем обозначение зу без номера контура
                                    Dim FullCondParcel As String = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionNumber)
                                    If Val(numbParcel) = 0 Then numbParcel = "0"
                                    '=====================================================================================================
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем все в массив
                                    ReDim Preserve ArrayPoligon(4, CountParcels)
                                    ArrayPoligon(0, CountParcels) = numbParcel 'номер участка
                                    ArrayPoligon(1, CountParcels) = entPLine.Handle.ToString 'хангл полилинии
                                    ArrayPoligon(2, CountParcels) = GLOBID 'глобальный ID
                                    ArrayPoligon(3, CountParcels) = LocalID 'локальный ID
                                    ArrayPoligon(4, CountParcels) = Delta 'погрешность
                                    CountParcels += 1
                                End If
                            End If
line1:
                        Next
                    End If
                    '==================================================================================================
                    If IsArray(ArrayPoligon) = False Then
                        MsgBox("Массив с участками пуст!!!")
                        Exit Sub
                    Else
                        If ArrayPoligon.GetUpperBound(1) > 0 Then
                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayPoligon, 0)
                        End If
                    End If
                    '====================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'массив с участками
                    For i As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                        'получаем доступ к полилинии для извлечения координат
                        Dim HandlePLine As String = ArrayPoligon(1, i)
                        Dim idPline As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                        Dim PlineObj As Autodesk.AutoCAD.DatabaseServices.Polyline = tr1.GetObject(idPline, OpenMode.ForRead)
                        Dim idLoc As String = ArrayPoligon(3, i) 'локальный ID
                        Dim idGlob As String = ArrayPoligon(2, i) 'глобальный ID
                        '=============================================================================================================
                        'ищем контура
                        Dim ArrayCounter As String(,) = Nothing
                        If idGlob.Trim.Length > 2 Then
                            Dim ArrayCounterTemp As String(,) = Nothing
                            Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(idGlob, ArrayCounterTemp, NamePSTable)
                            'если участок многоконтурный тогда выбираем все контура
                            If IsArray(ArrayCounterTemp) = True Then
                                If ArrayCounterTemp.GetUpperBound(1) > 0 Then
                                    'это многоконтурный участок ищем обозначения всех контуров
                                    For j As Integer = 0 To ArrayCounterTemp.GetUpperBound(1)
                                        Dim ConditionCounterParcel As String = "0"
                                        Dim LocalTempCounter As String = ArrayCounterTemp(1, j)
                                        Dim hgCounter As String = ArrayCounterTemp(0, j)
                                        Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                                        If IsNothing(idCounter) = False Then
                                            If idCounter.IsNull = False Then
                                                Dim EntCounter As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(idCounter, OpenMode.ForRead)
                                                If TypeOf EntCounter Is Polyline Then
                                                    Dim PlineCounter As Polyline = EntCounter
                                                    'ищем расширенные данные
                                                    ConditionCounterParcel = FuncPropertySet.FuncReadValueSetProperties(PlineCounter, "ConditionalNumber", NamePSTable)
                                                End If
                                            End If
                                        End If
                                        ConditionCounterParcel = FuncCadastr.FuncFindNumberCadastralObject(ConditionCounterParcel)
                                        Dim numbCounter As Integer = FuncCadastr.FuncFindNumberCounterCadastralObject(ConditionCounterParcel)
                                        ReDim Preserve ArrayCounter(2, j)
                                        ArrayCounter(0, j) = numbCounter 'номер контура
                                        ArrayCounter(1, j) = LocalTempCounter 'локальный id
                                        ArrayCounter(2, j) = hgCounter 'handle контура
                                    Next j
                                    'сортируем контура
                                    If ArrayCounter.GetUpperBound(1) > 0 Then
                                        Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCounter, 0)
                                    End If
                                Else
                                    'участок у которого есть GLOBALID но найден только один контур
                                    ReDim Preserve ArrayCounter(3, 0)
                                    ArrayCounter(0, 0) = 0
                                    ArrayCounter(1, 0) = idLoc 'локальный ID
                                    ArrayCounter(2, 0) = HandlePLine 'handle контура
                                End If
                            Else
                                'участок у которого есть GLOBALID но контуров не найдено
                                ReDim Preserve ArrayCounter(3, 0)
                                ArrayCounter(0, 0) = 0
                                ArrayCounter(1, 0) = idLoc 'локальный ID
                                ArrayCounter(2, 0) = HandlePLine 'handle контура
                            End If
                        Else
                            ReDim Preserve ArrayCounter(3, 0)
                            ArrayCounter(0, 0) = 0
                            ArrayCounter(1, 0) = idLoc 'локальный ID
                            ArrayCounter(2, 0) = HandlePLine 'handle контура
                        End If
                        '==================================================================================================
                        If IsArray(ArrayCounter) = True Then
                            If ArrayCounter.GetUpperBound(1) > 0 Then
                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCounter, 0)
                            End If
                        End If
                        '================================================================================================
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'начинаем доставать контуры из массива
                        For j As Integer = 0 To ArrayCounter.GetUpperBound(1)
                            Dim idLocCounter As String = ArrayCounter(1, j)
                            Dim HandleCounter As String = ArrayCounter(2, j)
                            Dim idPlineCounter As ObjectId = AcadFunction.FuncReturnObjectID(HandleCounter)
                            Dim ArrayOstrov1 As String(,) = Nothing 'массив с островками
                            If IsNothing(idPlineCounter) = False Then
                                If idPlineCounter.IsNull = False Then
                                    Dim PlineObjCounter As Autodesk.AutoCAD.DatabaseServices.Polyline = tr1.GetObject(idPlineCounter, OpenMode.ForWrite)
                                    If ColorObject = 0 Then ColorObject = PlineObjCounter.ColorIndex
                                    If LayerObject = "" Then LayerObject = PlineObjCounter.Layer
                                    Dim clockPline As Boolean = True
                                    If PlineObjCounter.NumberOfVertices > 2 Then
                                        Dim area As Double = MathFunction.GetArea(PlineObjCounter)
                                        If area > 0 Then
                                            clockPline = False
                                        End If
                                        If FormLoad.CheckBox2.Checked = True Then
                                            If area > 0 Then
                                                Dim boolClockVert As Boolean = MathFunction.FuncKadastrReversPline(PlineObjCounter)
                                                If boolClockVert = True Then
                                                    clockPline = True
                                                End If
                                            End If
                                        End If
                                    End If

                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLocCounter, ArrayOstrov1)
                                    If IsArray(ArrayOstrov1) = True Then
                                        If ArrayOstrov1.GetUpperBound(1) > 0 Then
                                            'сортируем массив
                                            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayOstrov1, 0)
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'нумеруем основной полигон
                                    '==================================================================
                                    Dim Vert0 As Point2d = New Point2d(0, 0)
                                    Dim Vert1 As Point2d = New Point2d(0, 0)
                                    Dim Vert2 As Point2d = New Point2d(0, 0)
                                    '=====================================================================================================
                                    'начинаем перебирать вершины полилинии
                                    For i2 As Integer = 0 To PlineObjCounter.NumberOfVertices - 1
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim CurrentCoordVert As Point2d = PlineObjCounter.GetPoint2dAt(i2)
                                        Dim NamePointZ As String = prefNumber & NumbPoint
                                        'делаем проверку дубликатов
                                        Dim StrHGPoint As String = ""
                                        Dim boolSt As Boolean = False
                                        For k1 As Integer = 0 To ArrayOldPoint.GetUpperBound(1)
                                            Dim XYTemp As Point2d = New Point2d(ArrayOldPoint(1, k1), ArrayOldPoint(2, k1))
                                            Dim LPoint As Double = CurrentCoordVert.GetVectorTo(XYTemp).Length
                                            If LPoint <= dopuskXY Then
                                                If ArrayOldPoint(4, k1) Like "0" Then
                                                    StrHGPoint = ArrayOldPoint(3, k1)
                                                    ArrayOldPoint(4, k1) = "1"
                                                    ArrayOldPoint(3, k1) = 0
                                                Else
                                                    boolSt = True
                                                End If
                                                Exit For
                                            End If
                                        Next k1
                                        'это точка уже была пронумерована предыдущим участком
                                        If boolSt = True Then
                                            Continue For
                                        End If
                                        '=================================================================================
                                        'формируем массив для записи
                                        Dim ArrayOBJ1 As String(,) = Nothing
                                        ReDim Preserve ArrayOBJ1(2, 5)
                                        ArrayOBJ1(0, 0) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 0) = "LABEL"
                                        ArrayOBJ1(2, 0) = DeskPoint

                                        ArrayOBJ1(0, 1) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 1) = "Number"
                                        ArrayOBJ1(2, 1) = NamePointZ

                                        ArrayOBJ1(0, 2) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 2) = "IDRNG"
                                        ArrayOBJ1(2, 2) = idLocCounter

                                        ArrayOBJ1(0, 3) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 3) = "Delta"
                                        ArrayOBJ1(2, 3) = Delta

                                        ArrayOBJ1(0, 4) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 4) = "FastDesk"
                                        ArrayOBJ1(2, 4) = Zakrep

                                        ArrayOBJ1(0, 5) = "CADASTRAL_POINT"
                                        ArrayOBJ1(1, 5) = "MethodCoordinates"
                                        ArrayOBJ1(2, 5) = MetodCoord
                                        If StrHGPoint.Trim.Length > 0 Then
                                            If FormLoad.RadioButton15.Checked = True Then
                                                '==============================================================================================================
                                                'меняем номер точки и больше ни чего
                                                Dim idPoint As ObjectId = AcadFunction.FuncReturnObjectID(StrHGPoint)
                                                If IsNothing(idPoint) = True Then Continue For
                                                If idPoint.IsNull = True Then Continue For
                                                Dim entPoint As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(idPoint, OpenMode.ForWrite)
                                                If TypeOf entPoint Is BlockReference Then
                                                    Dim blk As BlockReference = entPoint
                                                    Dim ArrayData As String(,) = Nothing
                                                    ReDim ArrayData(1, 0)
                                                    ArrayData(0, 0) = "NUMBER"
                                                    ArrayData(1, 0) = NamePointZ
                                                    Dim boolInsBlk As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayData)
                                                    If boolInsBlk = False Then
                                                        ArrayData(0, 0) = "NUM_GEOPOINT"
                                                        ArrayData(1, 0) = NamePointZ
                                                        Dim boolInsBlk2 As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayData)
                                                    End If
                                                ElseIf TypeOf entPoint Is MLeader Then
                                                    Dim acMleader As MLeader = entPoint
                                                    acMleader.MText.Contents = NamePointZ
                                                ElseIf TypeOf entPoint Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                    Dim cgPoint As Autodesk.Civil.DatabaseServices.CogoPoint = entPoint
                                                    cgPoint.Northing = CurrentCoordVert.Y
                                                    cgPoint.Easting = CurrentCoordVert.X
                                                End If
                                                If IsNothing(idTablePs) = False Then
                                                    If idTablePs.IsNull = False Then
                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(entPoint, ArrayOBJ1)
                                                    End If
                                                End If
                                                NumbPoint += 1
                                                Continue For
                                            Else
                                                Continue For
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вычисляем позицию вставки метки точки
                                        If i2 = 0 Then
                                            Vert0 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 1)
                                            Vert1 = PlineObjCounter.GetPoint2dAt(0)
                                            Vert2 = PlineObjCounter.GetPoint2dAt(1)
                                        ElseIf i2 = PlineObjCounter.NumberOfVertices - 1 Then
                                            Vert0 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 2)
                                            Vert1 = PlineObjCounter.GetPoint2dAt(PlineObjCounter.NumberOfVertices - 1)
                                            Vert2 = PlineObjCounter.GetPoint2dAt(0)
                                        Else
                                            Vert0 = PlineObjCounter.GetPoint2dAt(i2 - 1)
                                            Vert1 = PlineObjCounter.GetPoint2dAt(i2)
                                            Vert2 = PlineObjCounter.GetPoint2dAt(i2 + 1)
                                        End If
                                        Dim insAttr As Double() = Nothing 'массив с хорактеристиками
                                        Dim boolInsAttr As Boolean = FuncCadastr.FuncReplaseInsertCoordAttr2(Vert0, Vert1, Vert2, insAttr, DText, clockPline)
                                        Dim point3d1 As Point3d = New Point3d(Vert1.X, Vert1.Y, 0)
                                        '=====================================================================================
                                        If FormLoad.RadioButton11.Checked = True Then
                                            'вставляем блок Автокад
                                            'вставляем блок Автокад
                                            Dim blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(point3d1, NameObject, LayerObject, ColorObject, (HPoint * M) / 1000)
                                            '(Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
                                            If IsNothing(blk) = False Then
                                                Dim arrayDataAttrib As String(,) = {}
                                                ReDim arrayDataAttrib(6, 3)
                                                arrayDataAttrib(0, 0) = "NUMBER"
                                                arrayDataAttrib(1, 0) = NamePointZ
                                                arrayDataAttrib(3, 0) = (attributeHeight * M) / 1000
                                                arrayDataAttrib(4, 0) = rotateAttrib

                                                arrayDataAttrib(0, 1) = "DELTA"
                                                arrayDataAttrib(1, 1) = Delta

                                                arrayDataAttrib(0, 2) = "FASTDESK"
                                                arrayDataAttrib(1, 2) = Zakrep

                                                arrayDataAttrib(0, 3) = "METHODCOORDINATES"
                                                arrayDataAttrib(1, 3) = MetodCoord
                                                Dim boolWriteAttrai As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, arrayDataAttrib)
                                                'Dim InsAttribRez As ObjectId = FuncCadastr.FuncInsertCadastrPoint(point3d1, insAttr, NamePointZ, Delta, Zakrep, HText, LayerObject, NameObject, MashPoint, 0, idStyle, MetodCoord)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                blk.ColorIndex = ColorObject
                                                If FormLoad.CheckBox4.Checked = True Then
                                                    PointColection.Add(blk.Id)
                                                End If
                                                '=======================================================================================================
                                                If IsNothing(idTablePs) = False Then
                                                    If idTablePs.IsNull = False Then
                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(blk, idTablePs)
                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blk, ArrayOBJ1)
                                                    End If
                                                End If
                                            End If
                                        ElseIf FormLoad.RadioButton12.Checked = True Then
                                            'вставляем точку COGO
                                            Dim InsAttribRez As ObjectId = FuncSurvey.FuncCreateCoGoPoint(point3d1, NameObject)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'записываем расширенные данные в точку
                                            If IsNothing(InsAttribRez) = False Then
                                                If InsAttribRez.IsNull = False Then
                                                    Dim CoGoObj As Autodesk.Civil.DatabaseServices.CogoPoint = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                    CoGoObj.ColorIndex = ColorObject
                                                    If FormLoad.CheckBox4.Checked = True Then
                                                        PointColection.Add(CoGoObj.Id)
                                                    End If
                                                    If prefNumber.Trim.Length > 0 Then
                                                        Dim boolWriteUDP As Boolean = FuncSurvey.FuncWriteUDPValue(CoGoObj, "CADASTRAL_POINT", "Pref_number", prefNumber)
                                                    End If
                                                    '=======================================================================================================
                                                    If IsNothing(idTablePs) = False Then
                                                        If idTablePs.IsNull = False Then
                                                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoObj, idTablePs)
                                                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoObj, ArrayOBJ1)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf FormLoad.RadioButton13.Checked = True Then
                                            'вставляем мультивыноску
                                            If FormLoad.RadioButton15.Checked = True Then
                                                Dim offsetx As Double = insAttr(0) - point3d1.X
                                                Dim offsety As Double = insAttr(1) - point3d1.Y
                                                Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), NamePointZ, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'записываем расширенные данные в точку
                                                If IsNothing(InsAttribRez) = False Then
                                                    If InsAttribRez.IsNull = False Then
                                                        Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                        MleaderObj.ColorIndex = ColorObject
                                                        MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(MleaderObj.Id)
                                                        End If
                                                        '=======================================================================================================
                                                        If IsNothing(idTablePs) = False Then
                                                            If idTablePs.IsNull = False Then
                                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MleaderObj, idTablePs)
                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MleaderObj, ArrayOBJ1)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                Dim x As String = FuncFormatZN.FuncFormatCoordinate(point3d1.Y, 2)
                                                Dim y As String = FuncFormatZN.FuncFormatCoordinate(point3d1.X, 2)
                                                Dim TextMleader As String = "X=" & x & vbLf & "Y=" & y
                                                Dim offsetx As Double = insAttr(0) - point3d1.X
                                                Dim offsety As Double = insAttr(1) - point3d1.Y
                                                Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), TextMleader, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                If IsNothing(InsAttribRez) = False Then
                                                    If InsAttribRez.IsNull = False Then
                                                        Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                        MleaderObj.ColorIndex = ColorObject
                                                        MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(InsAttribRez)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf FormLoad.RadioButton16.Checked = True Then
                                            'вставляем метку Civil3D
                                            If FormLoad.RadioButton15.Checked = True Then

                                            Else
                                                Dim IDlb As ObjectId = Autodesk.Civil.DatabaseServices.NoteLabel.Create(point3d1, idLabelStyle, idMarkerStyle)
                                                If FormLoad.CheckBox4.Checked = True Then
                                                    PointColection.Add(IDlb)
                                                End If
                                                Dim lb As Autodesk.Civil.DatabaseServices.NoteLabel = tr1.GetObject(IDlb, OpenMode.ForWrite)
                                                lb.Layer = LayerObject
                                                lb.ColorIndex = ColorObject
                                                lb.LabelLocation = New Point3d(insAttr(0), insAttr(1), 0)
                                            End If
                                        End If
                                        NumbPoint += 1
                                    Next i2
                                End If
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'нумеруем островки
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            If IsArray(ArrayOstrov1) = True Then
                                For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                    Dim idLocOstrov As String = ArrayOstrov1(3, k)
                                    Dim hgOstrov As String = ArrayOstrov1(1, k)
                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                    If IsNothing(idOstrov) = False Then
                                        If idOstrov.IsNull = False Then
                                            Dim PlineObjOstrov As Polyline = tr1.GetObject(idOstrov, OpenMode.ForWrite)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            Dim clockPline As Boolean = True
                                            If PlineObjOstrov.NumberOfVertices > 2 Then
                                                Dim area As Double = MathFunction.GetArea(PlineObjOstrov)
                                                If area > 0 Then
                                                    clockPline = False
                                                End If
                                                If FormLoad.CheckBox2.Checked = True Then
                                                    If area > 0 Then
                                                        Dim boolClockVert As Boolean = MathFunction.FuncKadastrReversPline(PlineObjOstrov)
                                                        If boolClockVert = True Then
                                                            clockPline = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            Dim Vert0 As Point2d = New Point2d(0, 0)
                                            Dim Vert1 As Point2d = New Point2d(0, 0)
                                            Dim Vert2 As Point2d = New Point2d(0, 0)
                                            For i2 As Integer = 0 To PlineObjOstrov.NumberOfVertices - 1
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                Dim CurrentCoordVert As Point2d = PlineObjOstrov.GetPoint2dAt(i2)
                                                Dim NamePointZ As String = prefNumber & NumbPoint
                                                'делаем проверку дубликатов
                                                Dim StrHGPoint As String = ""
                                                Dim boolSt As Boolean = False
                                                For k1 As Integer = 0 To ArrayOldPoint.GetUpperBound(1)
                                                    Dim XYTemp As Point2d = New Point2d(ArrayOldPoint(1, k1), ArrayOldPoint(2, k1))
                                                    Dim LPoint As Double = CurrentCoordVert.GetVectorTo(XYTemp).Length
                                                    If LPoint <= dopuskXY Then
                                                        If ArrayOldPoint(4, k1) Like "0" Then
                                                            StrHGPoint = ArrayOldPoint(3, k1)
                                                            ArrayOldPoint(4, k1) = "1"
                                                            ArrayOldPoint(3, k1) = 0
                                                            Exit For
                                                            'Else
                                                            '    boolSt = True
                                                            '    Continue For
                                                        End If
                                                    End If
                                                Next k1
                                                'If boolSt = True Then
                                                '    If StrHGPoint.Trim.Length = 0 Then
                                                '        Continue For
                                                '    End If
                                                'End If
                                                '=================================================================================
                                                'формируем массив для записи
                                                Dim ArrayOBJ1 As String(,) = Nothing
                                                ReDim Preserve ArrayOBJ1(2, 5)
                                                ArrayOBJ1(0, 0) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 0) = "LABEL"
                                                ArrayOBJ1(2, 0) = DeskPoint

                                                ArrayOBJ1(0, 1) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 1) = "Number"
                                                ArrayOBJ1(2, 1) = NamePointZ

                                                ArrayOBJ1(0, 2) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 2) = "IDRNG"
                                                ArrayOBJ1(2, 2) = idLocCounter

                                                ArrayOBJ1(0, 3) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 3) = "Delta"
                                                ArrayOBJ1(2, 3) = Delta

                                                ArrayOBJ1(0, 4) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 4) = "FastDesk"
                                                ArrayOBJ1(2, 4) = Zakrep

                                                ArrayOBJ1(0, 5) = "CADASTRAL_POINT"
                                                ArrayOBJ1(1, 5) = "MethodCoordinates"
                                                ArrayOBJ1(2, 5) = MetodCoord

                                                If StrHGPoint.Trim.Length > 1 Then
                                                    If FormLoad.RadioButton15.Checked = True Then
                                                        '==============================================================================================================
                                                        'меняем номер точки и больше ни чего
                                                        Dim idPoint As ObjectId = AcadFunction.FuncReturnObjectID(StrHGPoint)
                                                        Dim entPoint As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(idPoint, OpenMode.ForWrite)
                                                        If TypeOf entPoint Is BlockReference Then
                                                            Dim blk As BlockReference = entPoint
                                                            Dim ArrayData As String(,) = Nothing
                                                            ReDim ArrayData(1, 0)
                                                            ArrayData(0, 0) = "NUMBER"
                                                            ArrayData(1, 0) = NamePointZ
                                                            Dim boolInsBlk As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayData)
                                                            If boolInsBlk = False Then
                                                                ArrayData(0, 0) = "NUM_GEOPOINT"
                                                                ArrayData(1, 0) = NamePointZ
                                                                Dim boolInsBlk2 As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, ArrayData)
                                                            End If
                                                        ElseIf TypeOf entPoint Is MLeader Then
                                                            Dim acMleader As MLeader = entPoint
                                                            acMleader.MText.Contents = NamePointZ
                                                        ElseIf TypeOf entPoint Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                                            Dim cgPoint As Autodesk.Civil.DatabaseServices.CogoPoint = entPoint
                                                            cgPoint.Northing = CurrentCoordVert.Y
                                                            cgPoint.Easting = CurrentCoordVert.X
                                                        End If
                                                        If IsNothing(idTablePs) = False Then
                                                            If idTablePs.IsNull = False Then
                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(entPoint, ArrayOBJ1)
                                                            End If
                                                        End If
                                                        NumbPoint += 1
                                                        Continue For
                                                    Else
                                                        Continue For
                                                    End If
                                                End If
                                                'точка не найдена вставляем новую точку
                                                If i2 = 0 Then
                                                    Vert0 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 1)
                                                    Vert1 = PlineObjOstrov.GetPoint2dAt(0)
                                                    Vert2 = PlineObjOstrov.GetPoint2dAt(1)
                                                ElseIf i2 = PlineObjOstrov.NumberOfVertices - 1 Then
                                                    Vert0 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 2)
                                                    Vert1 = PlineObjOstrov.GetPoint2dAt(PlineObjOstrov.NumberOfVertices - 1)
                                                    Vert2 = PlineObjOstrov.GetPoint2dAt(0)
                                                Else
                                                    Vert0 = PlineObjOstrov.GetPoint2dAt(i2 - 1)
                                                    Vert1 = PlineObjOstrov.GetPoint2dAt(i2)
                                                    Vert2 = PlineObjOstrov.GetPoint2dAt(i2 + 1)
                                                End If
                                                Dim insAttr As Double() = Nothing 'массив с хорактеристиками
                                                Dim boolInsAttr As Boolean = FuncCadastr.FuncReplaseInsertCoordAttr2(Vert0, Vert1, Vert2, insAttr, DText, clockPline)
                                                Dim point3d1 As Point3d = New Point3d(Vert1.X, Vert1.Y, 0)
                                                '=====================================================================================
                                                If FormLoad.RadioButton11.Checked = True Then
                                                    'вставляем блок Автокад
                                                    Dim blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(point3d1, NameObject, LayerObject, ColorObject, (HPoint * M) / 1000)
                                                    If IsNothing(blk) = False Then
                                                        '(Массив 0-имя атрибута, 1-содержание атрибута, 3-высота атрибута, 4-угол наклона атрибута,5-текстовый стиль,6-цвет)
                                                        Dim arrayDataAttrib As String(,) = {}
                                                        ReDim arrayDataAttrib(6, 3)
                                                        arrayDataAttrib(0, 0) = "NUMBER"
                                                        arrayDataAttrib(1, 0) = NamePointZ
                                                        arrayDataAttrib(3, 0) = (attributeHeight * M) / 1000
                                                        arrayDataAttrib(4, 0) = rotateAttrib

                                                        arrayDataAttrib(0, 1) = "DELTA"
                                                        arrayDataAttrib(1, 1) = Delta

                                                        arrayDataAttrib(0, 2) = "FASTDESK"
                                                        arrayDataAttrib(1, 2) = Zakrep

                                                        arrayDataAttrib(0, 3) = "METHODCOORDINATES"
                                                        arrayDataAttrib(1, 3) = MetodCoord
                                                        Dim boolWriteAttrai As Boolean = AcadFunction.FuncWriteAllBlockAttr(blk, arrayDataAttrib)
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем данные в точку
                                                        blk.ColorIndex = ColorObject
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(blk.Id)
                                                        End If
                                                        '=======================================================================================================
                                                        If IsNothing(idTablePs) = False Then
                                                            If idTablePs.IsNull = False Then
                                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(blk, idTablePs)
                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(blk, ArrayOBJ1)
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf FormLoad.RadioButton12.Checked = True Then
                                                    'вставляем точку COGO
                                                    Dim InsAttribRez As ObjectId = FuncSurvey.FuncCreateCoGoPoint(point3d1, NameObject)
                                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                    'записываем расширенные данные в точку
                                                    If IsNothing(InsAttribRez) = False Then
                                                        If InsAttribRez.IsNull = False Then
                                                            Dim CoGoObj As Autodesk.Civil.DatabaseServices.CogoPoint = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                            CoGoObj.ColorIndex = ColorObject
                                                            If FormLoad.CheckBox4.Checked = True Then
                                                                PointColection.Add(CoGoObj.Id)
                                                            End If
                                                            If prefNumber.Trim.Length > 0 Then
                                                                Dim boolWriteUDP As Boolean = FuncSurvey.FuncWriteUDPValue(CoGoObj, "CADASTRAL_POINT", "Pref_number", prefNumber)
                                                            End If
                                                            '=======================================================================================================
                                                            If IsNothing(idTablePs) = False Then
                                                                If idTablePs.IsNull = False Then
                                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(CoGoObj, idTablePs)
                                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(CoGoObj, ArrayOBJ1)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf FormLoad.RadioButton13.Checked = True Then
                                                    'вставляем мультивыноску
                                                    If FormLoad.RadioButton15.Checked = True Then
                                                        Dim offsetx As Double = insAttr(0) - point3d1.X
                                                        Dim offsety As Double = insAttr(1) - point3d1.Y
                                                        Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), NamePointZ, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'записываем расширенные данные в точку
                                                        If IsNothing(InsAttribRez) = False Then
                                                            If InsAttribRez.IsNull = False Then
                                                                Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                                MleaderObj.ColorIndex = ColorObject
                                                                MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                                If FormLoad.CheckBox4.Checked = True Then
                                                                    PointColection.Add(MleaderObj.Id)
                                                                End If
                                                                '=======================================================================================================
                                                                If IsNothing(idTablePs) = False Then
                                                                    If idTablePs.IsNull = False Then
                                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MleaderObj, idTablePs)
                                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MleaderObj, ArrayOBJ1)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        Dim x As String = FuncFormatZN.FuncFormatCoordinate(point3d1.Y, 2)
                                                        Dim y As String = FuncFormatZN.FuncFormatCoordinate(point3d1.X, 2)
                                                        Dim TextMleader As String = "X=" & x & vbLf & "Y=" & y
                                                        Dim offsetx As Double = insAttr(0) - point3d1.X
                                                        Dim offsety As Double = insAttr(1) - point3d1.Y
                                                        Dim InsAttribRez As ObjectId = AcadFunction.FuncDrawMleader(point3d1.Convert2d(New Plane), TextMleader, offsetx, offsety, LayerObject, HText, idStyleMleader)
                                                        If IsNothing(InsAttribRez) = False Then
                                                            If InsAttribRez.IsNull = False Then
                                                                Dim MleaderObj As MLeader = tr1.GetObject(InsAttribRez, OpenMode.ForWrite)
                                                                MleaderObj.ColorIndex = ColorObject
                                                                MleaderObj.ArrowSize = (MleaderObj.ArrowSize * M) / 1000
                                                                If FormLoad.CheckBox4.Checked = True Then
                                                                    PointColection.Add(InsAttribRez)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                ElseIf FormLoad.RadioButton16.Checked = True Then
                                                    'вставляем метку Civil3D
                                                    If FormLoad.RadioButton15.Checked = True Then

                                                    Else
                                                        Dim IDlb As ObjectId = Autodesk.Civil.DatabaseServices.NoteLabel.Create(point3d1, idLabelStyle, idMarkerStyle)
                                                        If FormLoad.CheckBox4.Checked = True Then
                                                            PointColection.Add(IDlb)
                                                        End If
                                                        Dim lb As Autodesk.Civil.DatabaseServices.NoteLabel = tr1.GetObject(IDlb, OpenMode.ForWrite)
                                                        lb.Layer = LayerObject
                                                        lb.ColorIndex = ColorObject
                                                        lb.LabelLocation = New Point3d(insAttr(0), insAttr(1), 0)
                                                    End If
                                                End If
                                                NumbPoint += 1
                                            Next i2
                                        End If
                                    End If
                                Next k
                            End If
                        Next j
                    Next i
                    Dim mgBox As MsgBoxResult = MsgBox("Макро выполнено. Удалить лишние точки?", MsgBoxStyle.OkCancel)
                    If mgBox = MsgBoxResult.Ok Then
                        For i As Integer = 0 To ArrayOldPoint.GetUpperBound(1)
                            Dim hgBlk As String = ArrayOldPoint(3, i)
                            If hgBlk.Trim.Length > 1 Then
                                Dim idPoint As ObjectId = AcadFunction.FuncReturnObjectID(hgBlk)
                                Dim entPoint As Autodesk.AutoCAD.DatabaseServices.Entity = tr1.GetObject(idPoint, OpenMode.ForWrite)
                                entPoint.Erase()  ' <-- Этого не хватало для удаления
                            End If
                        Next i
                    End If
                Catch
                End Try
                tr1.Commit()
            End Using
        End Using
    End Sub
    '==============================================================================================
    'статус точки
    <CommandMethod("StatusCadastralPoint")>
    Public Sub StatusCadastralPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
            Dim formStatPoint As FormStatusPoint = New FormStatusPoint
            formStatPoint.ShowDialog()
        Catch ex As System.Exception
        End Try
    End Sub
    '==============================================================================================
    'найти наибольший номер точки
    <CommandMethod("oldFindGeopointToLargNumber")>
    Public Sub oldFindGeopointToLargNumber()
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
        Try
            Dim arrayNameBlk As String() = {"PROJECT_POINT", "TEMP_POINT"}
            Dim arrayTag As String() = {"Имя_точки", "NUM_GEOPOINT", "NUMBER"}
            Dim blkRez As BlockReference = FuncCadastr.FuncReturnBlkToLargNumber(arrayTag, arrayNameBlk)
            If IsNothing(blkRez) = False Then
                If blkRez.ObjectId.IsNull = False Then
                    Dim minPoint As Point2d = blkRez.Bounds.Value.MinPoint.Convert2d(New Plane)
                    Dim maxPoint As Point2d = blkRez.Bounds.Value.MaxPoint.Convert2d(New Plane)
                    Dim arrayID As ObjectId() = {blkRez.ObjectId}
                    If IsArray(arrayID) = True Then
                        ed.SetImpliedSelection(arrayID)
                    End If
                    Dim sentrePoint As Point3d = blkRez.Position
                    Dim minPoint1 As Point3d = New Point3d(sentrePoint.X - 10, sentrePoint.Y - 10, 0)
                    Dim maxPoint1 As Point3d = New Point3d(sentrePoint.X + 10, sentrePoint.Y + 10, 0)
                    SubZoom.Zoom(minPoint1, maxPoint1, New Point3d(), 1)
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Sub
    '=============================================================================================
    'Удалить точки
    <CommandMethod("oldDeleteCadastralPoint")>
    Public Sub oldDeleteCadastralPoint()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        Dim countDeletePoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Удалить точки?"
                pIntOpts.Keywords.Add("Все")
                pIntOpts.Keywords.Add("Новые")
                pIntOpts.Keywords.Add("Существующие")
                pIntOpts.Keywords.Add("сТарые")
                pIntOpts.Keywords.Add("Полигон")
                pIntOpts.Keywords.Add("Диапазон")
                pIntOpts.Keywords.Default = "Полигон"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                Dim arrayNameBlk As String() = {"PROJECT_POINT", "TEMP_POINT"}
                Select Case pIntRes.StringResult
                    Case "Все"
                        Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                        If pointIDColl.Count > 0 Then
                            countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(pointIDColl, 0)
                        End If
                    Case "Новые"
                        Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                        If pointIDColl.Count > 0 Then
                            countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(pointIDColl, 1)
                        End If
                    Case "Существующие"
                        Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                        If pointIDColl.Count > 0 Then
                            countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(pointIDColl, 2)
                        End If
                    Case "сТарые"
                        Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                        If pointIDColl.Count > 0 Then
                            countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(pointIDColl, 3)
                        End If
                    Case "Полигон"
                        'выбираем примитивы
                        Dim acSSPrompt As PromptSelectionResult
                        acSSPrompt = ed.GetSelection()
                        Dim acSSet1 As SelectionSet
                        If acSSPrompt.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                            acSSet1 = acSSPrompt.Value
                            For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                Dim ent As Entity = tr.GetObject(IDObj, OpenMode.ForRead, False)
                                If TypeOf ent Is Polyline Then
                                    Dim Pline As Polyline = ent
                                    Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                                    Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                                    For i As Integer = 0 To Pline.NumberOfVertices - 1
                                        Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                        Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                        If IsNothing(idGeoPoint) = False Then
                                            newPointIDColl.Add(idGeoPoint)
                                        End If
                                    Next i
                                    If newPointIDColl.Count > 0 Then
                                        countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(newPointIDColl, 0)
                                    End If
                                End If
                            Next
                        End If
                    Case "Диапазон"
                        Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                        Dim pointStartPline As Point2d = Nothing
                        Dim numberVertexStart As Integer = -1
                        Dim pointEndPline As Point2d = Nothing
                        Dim numberVertexEnd As Integer = -1
                        Dim prPointRes1 As PromptEntityResult = Nothing
                        prPointRes1 = ed.GetEntity(vbLf & "Выберите первую точку ближайшую к узлу полилинии:")
                        If prPointRes1.Status = PromptStatus.OK Then
                            Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                            If IsDBNull(entObjID) = False Then
                                Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Pline = entObj
                                    Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                    Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                    Dim Dist As Double = 1000000
                                    For i As Integer = 0 To Pline.NumberOfVertices - 1
                                        Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                        Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                        If L < Dist Then
                                            Dist = L
                                            pointStartPline = tempPoint
                                            numberVertexStart = i
                                        End If
                                    Next i
                                End If
                            End If
                        End If

                        prPointRes1 = ed.GetEntity(vbLf & "Выберите последнюю точку ближайшую к узлу полилинии:")
                        If prPointRes1.Status = PromptStatus.OK Then
                            Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                            If IsDBNull(entObjID) = False Then
                                Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    If IsNothing(Pline) = False Then
                                        If Pline.Id <> entObjID Then
                                            MsgBox("Изменить статус точкам можно только если они находятся на одной и тойже полилинии. Макро превано!!!!")
                                            Exit Sub
                                        End If
                                    Else
                                        Pline = entObj
                                    End If
                                    Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                    Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                    Dim Dist As Double = 1000000
                                    For i As Integer = 0 To Pline.NumberOfVertices - 1
                                        Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                        Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                        If L < Dist Then
                                            Dist = L
                                            pointEndPline = tempPoint
                                            numberVertexEnd = i
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                        If IsNothing(Pline) = False Then
                            If numberVertexStart = -1 Then
                                numberVertexStart = 0
                            End If
                            If numberVertexEnd = -1 Then
                                numberVertexEnd = Pline.NumberOfVertices - 1
                            End If
                        End If
                        Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                        If idPointColl.Count > 0 Then
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'меняем статус точки
                            Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                            For i As Integer = numberVertexStart To numberVertexEnd
                                Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                'ищем геоточку в узле
                                Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                If IsNothing(idGeoPoint) = False Then
                                    If idGeoPoint.IsNull = False Then
                                        'меняем статус
                                        newPointIDColl.Add(idGeoPoint)
                                    End If
                                End If
                            Next i
                            If newPointIDColl.Count > 0 Then
                                countDeletePoint = FuncCadastr.FuncDeleteGeoPoint(newPointIDColl, 0)
                            End If
                        End If
                End Select
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось удалить все точки.")
                tr.Commit()
                db.Dispose()
                Exit Sub
            End Try
            tr.Commit()
            db.Dispose()
            MsgBox("Всего удалено " & countDeletePoint & " точек.")
        End Using
    End Sub
    '=============================================================================================
    'Найти точку
    <CommandMethod("oldFindCadactralPointByNumber")>
    Public Sub oldFindCadactralPointByNumber()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim boolFindNumber As Boolean = False
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
        Dim arrayNameBlk As String() = {"PROJECT_POINT", "TEMP_POINT"}
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите номер точки: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                pStrOpts.DefaultValue = 1
                Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts)
                If pStrRes.Value > 0 Then
                    Dim numFindPoint As Integer = pStrRes.Value
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) 'открываем таблицу блоков на чтение
                    Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                    For Each id As ObjectId In btr 'цикл для перебора всех объектов на чертеже
                        Dim ent As Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                        If TypeOf ent Is BlockReference Then 'если объект имеет блок с атрибутом
                            Dim blk As BlockReference = ent 'ссылка на атрибут
                            Dim nameBlk As String = AcadFunction.FuncReadNameBlk(blk)
                            Dim boolZn As Boolean = MathFunction.FuncFindValueToArray(arrayNameBlk, nameBlk)
                            If boolZn = True Then
                                Dim numGeoPoint As String = AcadFunction.FuncReadValueAttr(blk, "NUMBER")
                                Dim result As System.Text.RegularExpressions.MatchCollection = New System.Text.RegularExpressions.Regex("[0-9]+").Matches(numGeoPoint)
                                If result.Count > 0 Then
                                    Dim value As Integer = Val(result.Item(0).Value)
                                    If value = numFindPoint Then

                                        Dim pMin As Double() = {blk.Position.X - blk.ScaleFactors.X, blk.Position.Y - blk.ScaleFactors.Y, 0}
                                        Dim pMax As Double() = {blk.Position.X + blk.ScaleFactors.X, blk.Position.Y + blk.ScaleFactors.Y, 0}
                                        Dim centrePoint As Point3d = blk.Position
                                        Dim minPoint1 As Point3d = New Point3d(centrePoint.X - 10, centrePoint.Y - 10, 0)
                                        Dim maxPoint1 As Point3d = New Point3d(centrePoint.X + 10, centrePoint.Y + 10, 0)
                                        SubZoom.Zoom(minPoint1, maxPoint1, New Point3d(), 1)
                                        Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                                        pIntOpts.Message = vbCrLf & "Искать еще точки с таким номером?"
                                        pIntOpts.Keywords.Add("Да")
                                        pIntOpts.Keywords.Add("Нет")
                                        pIntOpts.Keywords.Default = "Да"
                                        pIntOpts.AllowNone = True
                                        Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                                        If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                                            Throw New System.Exception("Прервано пользователем")
                                            Exit For 'если нажали ESC то выходим из программы
                                        End If
                                        If pIntRes.StringResult Like "нет" Then Exit For
                                        boolFindNumber = True
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Ошибка. Макро прервано")
                tr.Commit()
                db.Dispose()
                Exit Sub
            End Try
            If boolFindNumber - False Then MsgBox("Точки с таким номером нет на чертеже!!!")
            tr.Commit()
            db.Dispose()
        End Using
    End Sub
    '=============================================================================================
    'Изменить масштаб точки
    <CommandMethod("oldScaleCadastralPoint")>
    Public Sub oldScaleCadastralPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim CountPoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Изменить масштаб точкам?"
                pIntOpts.Keywords.Add("Все")
                pIntOpts.Keywords.Add("Полигон")
                pIntOpts.Keywords.Add("Диапазон")
                pIntOpts.Keywords.Default = "Полигон"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                'редактирование координат текстового файла
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите новый масштаб блока (точки): ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                pStrOpts.DefaultValue = 1
                Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
                If pStrRes.Value > 0 Then
                    Dim MBlk As Double = pStrRes.Value
                    Dim arrayNameBlk As String() = {"PROJECT_POINT", "TEMP_POINT"}
                    Select Case pIntRes.StringResult
                        Case "Все"
                            Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                            If pointIDColl.Count > 0 Then
                                For Each idBlk As ObjectId In pointIDColl
                                    Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                    Dim newScale As Double = MBlk / blk.ScaleFactors.X
                                    blk.TransformBy(Matrix3d.Scaling(newScale, blk.Position))
                                    CountPoint += 1
                                Next
                            End If
                        Case "Полигон"
                            'выбираем примитивы
                            Dim acSSPrompt As PromptSelectionResult
                            acSSPrompt = ed.GetSelection()
                            Dim acSSet1 As SelectionSet
                            If acSSPrompt.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                                acSSet1 = acSSPrompt.Value
                                For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                    Dim ent As Entity = tr.GetObject(IDObj, OpenMode.ForRead, False)
                                    If TypeOf ent Is Polyline Then
                                        Dim Pline As Polyline = ent
                                        Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                                        Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                            Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                            If IsNothing(idGeoPoint) = False Then
                                                newPointIDColl.Add(idGeoPoint)
                                            End If
                                        Next i
                                        If newPointIDColl.Count > 0 Then
                                            For Each idBlk As ObjectId In newPointIDColl
                                                Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                                Dim newScale As Double = MBlk / blk.ScaleFactors.X
                                                blk.TransformBy(Matrix3d.Scaling(newScale, blk.Position))
                                                CountPoint += 1
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        Case "Диапазон"
                            Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                            Dim pointStartPline As Point2d = Nothing
                            Dim numberVertexStart As Integer = -1
                            Dim pointEndPline As Point2d = Nothing
                            Dim numberVertexEnd As Integer = -1
                            Dim prPointRes1 As PromptEntityResult = Nothing
                            prPointRes1 = ed.GetEntity(vbLf & "Выберите первую точку ближайшую к узлу полилинии:")
                            If prPointRes1.Status = PromptStatus.OK Then
                                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                                If IsDBNull(entObjID) = False Then
                                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                        Pline = entObj
                                        Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                        Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                        Dim Dist As Double = 1000000
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                            Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                            If L < Dist Then
                                                Dist = L
                                                pointStartPline = tempPoint
                                                numberVertexStart = i
                                            End If
                                        Next i
                                    End If
                                End If
                            End If

                            prPointRes1 = ed.GetEntity(vbLf & "Выберите последнюю точку ближайшую к узлу полилинии:")
                            If prPointRes1.Status = PromptStatus.OK Then
                                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                                If IsDBNull(entObjID) = False Then
                                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                        If IsNothing(Pline) = False Then
                                            If Pline.Id <> entObjID Then
                                                MsgBox("Изменить статус точкам можно только если они находятся на одной и тойже полилинии. Макро превано!!!!")
                                                Exit Sub
                                            End If
                                        Else
                                            Pline = entObj
                                        End If
                                        Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                        Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                        Dim Dist As Double = 1000000
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                            Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                            If L < Dist Then
                                                Dist = L
                                                pointEndPline = tempPoint
                                                numberVertexEnd = i
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                            If IsNothing(Pline) = False Then
                                If numberVertexStart = -1 Then
                                    numberVertexStart = 0
                                End If
                                If numberVertexEnd = -1 Then
                                    numberVertexEnd = Pline.NumberOfVertices - 1
                                End If
                            End If
                            Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                            If idPointColl.Count > 0 Then
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'меняем статус точки
                                Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                                For i As Integer = numberVertexStart To numberVertexEnd
                                    Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                    'ищем геоточку в узле
                                    Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                    If IsNothing(idGeoPoint) = False Then
                                        If idGeoPoint.IsNull = False Then
                                            'меняем статус
                                            newPointIDColl.Add(idGeoPoint)
                                        End If
                                    End If
                                Next i
                                If newPointIDColl.Count > 0 Then
                                    For Each idBlk As ObjectId In newPointIDColl
                                        Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                        Dim newScale As Double = MBlk / blk.ScaleFactors.X
                                        blk.TransformBy(Matrix3d.Scaling(newScale, blk.Position))
                                        CountPoint += 1
                                    Next
                                End If
                            End If
                    End Select
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить масштаб точкам! Макро прервано")
                tr.Commit()
                db.Dispose()
                Exit Sub
            End Try
            tr.Commit()
            db.Dispose()
            MsgBox("Всего преобразовано " & CountPoint & " точек.")
        End Using
    End Sub

    '=============================================================================================
    'Изменить размер атрибута
    <CommandMethod("oldheightAttributeCadastralPoint")>
    Public Sub oldHeightAttributeCadastralPoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim CountPoint As Integer = 0
        Dim arrayNameBlk As String() = {}
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim prPointRes1 As PromptEntityResult = Nothing
                prPointRes1 = ed.GetEntity(vbLf & "Выберите один блок, для образца:")
                If prPointRes1.Status = PromptStatus.OK Then
                    Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                    If IsDBNull(entObjID) = False Then
                        Dim entObject As Entity = entObjID.GetObject(OpenMode.ForRead)
                        If TypeOf entObject Is BlockReference Then
                            Dim blk As BlockReference = entObject
                            ReDim arrayNameBlk(0)
                            arrayNameBlk(0) = blk.Name
                        End If
                    End If
                End If
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Изменить размер атрибута для блоков: "
                pIntOpts.Keywords.Add("Все")
                pIntOpts.Keywords.Add("Полигон")
                pIntOpts.Keywords.Add("Диапазон")
                pIntOpts.Keywords.Default = "Полигон"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                'редактирование координат текстового файла
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите новый размер атрибута: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                pStrOpts.DefaultValue = 1
                Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
                If pStrRes.Value > 0 Then
                    Dim heightBlk As Double = pStrRes.Value
                    Select Case pIntRes.StringResult
                        Case "Все"
                            Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                            If pointIDColl.Count > 0 Then
                                For Each idBlk As ObjectId In pointIDColl
                                    Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                    Dim point As Point2d = New Point2d(0, 0)
                                    Dim boolIzm As Boolean = AcadFunction.setHeightAndPositionArraibute(blk, heightBlk, point)
                                    CountPoint += 1
                                Next
                            End If
                        Case "Полигон"
                            'выбираем примитивы
                            Dim acSSPrompt As PromptSelectionResult
                            acSSPrompt = ed.GetSelection()
                            Dim acSSet1 As SelectionSet
                            If acSSPrompt.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                                acSSet1 = acSSPrompt.Value
                                For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                    Dim ent As Entity = tr.GetObject(IDObj, OpenMode.ForRead, False)
                                    If TypeOf ent Is Polyline Then
                                        Dim Pline As Polyline = ent
                                        Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                                        Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                            Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                            If IsNothing(idGeoPoint) = False Then
                                                newPointIDColl.Add(idGeoPoint)
                                            End If
                                        Next i
                                        If newPointIDColl.Count > 0 Then
                                            For Each idBlk As ObjectId In newPointIDColl
                                                Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                                Dim point As Point2d = New Point2d(0, 0)
                                                Dim boolIzm As Boolean = AcadFunction.setHeightAndPositionArraibute(blk, heightBlk, point)
                                                CountPoint += 1
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        Case "Диапазон"
                            Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                            Dim pointStartPline As Point2d = Nothing
                            Dim numberVertexStart As Integer = -1
                            Dim pointEndPline As Point2d = Nothing
                            Dim numberVertexEnd As Integer = -1
                            prPointRes1 = ed.GetEntity(vbLf & "Выберите первую точку ближайшую к узлу полилинии:")
                            If prPointRes1.Status = PromptStatus.OK Then
                                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                                If IsDBNull(entObjID) = False Then
                                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                        Pline = entObj
                                        Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                        Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                        Dim Dist As Double = 1000000
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                            Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                            If L < Dist Then
                                                Dist = L
                                                pointStartPline = tempPoint
                                                numberVertexStart = i
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                            prPointRes1 = ed.GetEntity(vbLf & "Выберите последнюю точку ближайшую к узлу полилинии:")
                            If prPointRes1.Status = PromptStatus.OK Then
                                Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                                If IsDBNull(entObjID) = False Then
                                    Dim entObj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(entObjID, OpenMode.ForRead)
                                    If TypeOf entObj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                        If IsNothing(Pline) = False Then
                                            If Pline.Id <> entObjID Then
                                                MsgBox("Изменить статус точкам можно только если они находятся на одной и тойже полилинии. Макро превано!!!!")
                                                Exit Sub
                                            End If
                                        Else
                                            Pline = entObj
                                        End If
                                        Dim point3d1 As Point3d = prPointRes1.PickedPoint 'смотрим координаты точки указания
                                        Dim point2dPic As Point2d = point3d1.Convert2d(New Plane)
                                        Dim Dist As Double = 1000000
                                        For i As Integer = 0 To Pline.NumberOfVertices - 1
                                            Dim tempPoint As Point2d = Pline.GetPoint2dAt(i)
                                            Dim L As Double = tempPoint.GetVectorTo(point2dPic).Length
                                            If L < Dist Then
                                                Dist = L
                                                pointEndPline = tempPoint
                                                numberVertexEnd = i
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                            If IsNothing(Pline) = False Then
                                If numberVertexStart = -1 Then
                                    numberVertexStart = 0
                                End If
                                If numberVertexEnd = -1 Then
                                    numberVertexEnd = Pline.NumberOfVertices - 1
                                End If
                            End If
                            Dim idPointColl As ObjectIdCollection = FuncCadastr.FuncFilterGeoPointByBoundObject(Pline, arrayNameBlk)
                            If idPointColl.Count > 0 Then
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'меняем статус точки
                                Dim newPointIDColl As ObjectIdCollection = New ObjectIdCollection
                                For i As Integer = numberVertexStart To numberVertexEnd
                                    Dim pt As Point2d = Pline.GetPoint2dAt(i)
                                    'ищем геоточку в узле
                                    Dim idGeoPoint As ObjectId = FuncCadastr.FuncFindGeoPoint(idPointColl, pt, 0.01)
                                    If IsNothing(idGeoPoint) = False Then
                                        If idGeoPoint.IsNull = False Then
                                            'меняем статус
                                            newPointIDColl.Add(idGeoPoint)
                                        End If
                                    End If
                                Next i
                                If newPointIDColl.Count > 0 Then
                                    For Each idBlk As ObjectId In newPointIDColl
                                        Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                        Dim point As Point2d = New Point2d(0, 0)
                                        Dim boolIzm As Boolean = AcadFunction.setHeightAndPositionArraibute(blk, heightBlk, point)
                                        CountPoint += 1
                                    Next
                                End If
                            End If
                    End Select
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить масштаб точкам! Макро прервано")
                tr.Commit()
                db.Dispose()
                Exit Sub
            End Try
            tr.Commit()
            db.Dispose()
            MsgBox("Всего преобразовано " & CountPoint & " точек.")
        End Using
    End Sub
    '=============================================================================================
    'Изменить местоположение атрибута
    <CommandMethod("oldPositionAttributePoint")>
    Public Sub oldPositionAttributePoint()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim CountPoint As Integer = 0
        Dim arrayNameBlk As String() = {}
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim arrayAttributeProperties As String(,) = {}
                Dim countArrayAttributeProperties As Integer = 0
                Dim prPointRes1 As PromptEntityResult = Nothing
                prPointRes1 = ed.GetEntity(vbLf & "Выберите один блок, для образца:")
                If prPointRes1.Status = PromptStatus.OK Then
                    Dim entObjID As Autodesk.AutoCAD.DatabaseServices.ObjectId = prPointRes1.ObjectId
                    If IsDBNull(entObjID) = False Then
                        Dim entObject As Entity = entObjID.GetObject(OpenMode.ForRead)
                        If TypeOf entObject Is BlockReference Then
                            Dim blk As BlockReference = entObject
                            If blk.AttributeCollection.Count > 0 Then
                                Dim attCol As AttributeCollection = blk.AttributeCollection
                                For Each attId As ObjectId In attCol
                                    Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForWrite)
                                    If attRef.Visible = True Then
                                        ReDim Preserve arrayAttributeProperties(2, countArrayAttributeProperties)
                                        arrayAttributeProperties(0, countArrayAttributeProperties) = attRef.Tag
                                        arrayAttributeProperties(1, countArrayAttributeProperties) = attRef.Position.X - blk.Position.X
                                        arrayAttributeProperties(2, countArrayAttributeProperties) = attRef.Position.Y - blk.Position.Y
                                        countArrayAttributeProperties += 1
                                        ReDim arrayNameBlk(0)
                                        arrayNameBlk(0) = blk.Name
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Изменить положение атрибутов для блоков: "
                pIntOpts.Keywords.Add("Все")
                pIntOpts.Keywords.Add("Диапазон")
                pIntOpts.Keywords.Default = "Диапазон"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If

                Select Case pIntRes.StringResult
                    Case "Все"
                        Dim pointIDColl As ObjectIdCollection = FuncCadastr.FuncFilterAllGeoPoint(arrayNameBlk)
                        If pointIDColl.Count > 0 Then
                            For Each idBlk As ObjectId In pointIDColl
                                Dim blk As BlockReference = idBlk.GetObject(OpenMode.ForWrite)
                                Dim boolIzm As Boolean = AcadFunction.setPositionAndPositionArraibute(blk, arrayAttributeProperties)
                                CountPoint += 1
                            Next
                        End If
                    Case "Диапазон"
                        'выбираем примитивы
                        Dim acSSPrompt As PromptSelectionResult
                        acSSPrompt = ed.GetSelection()
                        Dim acSSet1 As SelectionSet
                        If acSSPrompt.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                            acSSet1 = acSSPrompt.Value
                            For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                Dim ent As Entity = tr.GetObject(IDObj, OpenMode.ForWrite, False)
                                If TypeOf ent Is BlockReference Then
                                    Dim blk As BlockReference = ent
                                    Dim boolIzm As Boolean = AcadFunction.setPositionAndPositionArraibute(blk, arrayAttributeProperties)
                                    CountPoint += 1
                                End If
                            Next
                        End If
                End Select
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить масштаб точкам! Макро прервано")
                tr.Commit()
                db.Dispose()
                Exit Sub
            End Try
            tr.Commit()
            db.Dispose()
            MsgBox("Всего преобразовано " & CountPoint & " точек.")
        End Using
    End Sub
End Class
