Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Configuration
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Colors
Imports System.IO
'ведомость углов поворота прямых и кривых, ведомость разбивки кривых в плане
Public Class ListUPLinesAndCurves
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '================================================================================================================================
    'ведомость углов поворота прямых кривых
    '================================================================================================================================
    <CommandMethod("ListAlignAnglRot")>
    Public Sub ListAlignAnglRot()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acCurDb As Database = acDoc.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
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
        Dim Align As Alignment = Nothing 'трасса для отчета
        'временный каталог для автосохранения таблицы
        Dim TempDir As String = ""
        Try
            TempDir = System.IO.Path.GetTempPath()
        Catch ex As System.Security.SecurityException
            TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
        End Try
        '================================================================================================
        'вызываем диалоговое окно
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormRoadPlane As RoadPlane = New RoadPlane
        FormRoadPlane.ShowDialog()
        If FormRoadPlane.boolStatus = False Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim NameSelectAlign As String = FormRoadPlane.ComboBox5.Text
                    If NameSelectAlign.Trim.Length = 0 Then
                        MsgBox("Поле с наименованием трассы пустое. Выберите трассу из списка!!! Макро прервано.")
                        Exit Sub
                    End If
                    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    '=====================================================================================================================
                    'делаем ссылку на трассу
                    Dim AlignCollection As Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection = CivilDoc.GetAlignmentIds
                    If AlignCollection.Count > 0 Then
                        For Each idAlign As ObjectId In AlignCollection
                            Dim AlignTemp As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                            If AlignTemp.Name Like NameSelectAlign Then
                                Align = AlignTemp
                                Exit For
                            End If
                        Next
                    Else
                        MsgBox("В активном чертеже нет объектов типа ТРАССА!!! Макро прервано.")
                        Exit Sub
                    End If
                    If IsNothing(Align) = True Then
                        MsgBox("Трасса с выбранным именем, не найдена!!! Макро прервано")
                        Exit Sub
                    End If
                    If Align.Length = 0 Then
                        MsgBox("Трасса имеет нулевую длину!!! Макро прервано")
                        Exit Sub
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '============================================================================================================
                    'Raw пикет начала и конца трассы трассы
                    Dim PKStartAlign As Double = Val(FormRoadPlane.TextBox2.Tag)
                    'Raw пикет конца трассы трассы
                    Dim PKEndAlign As Double = Val(FormRoadPlane.TextBox3.Tag)
                    'проверка пикетов
                    If PKStartAlign > PKEndAlign Then
                        Dim TempStartPK As Double = PKStartAlign
                        PKStartAlign = PKEndAlign
                        PKEndAlign = TempStartPK
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '===============================================================================================================
                    'вычисляем координаты начала и конца трассы
                    Dim AlignStartX As Double = Align.StartPoint.X
                    Dim AlignStartY As Double = Align.StartPoint.Y
                    Dim AlignEndX As Double = Align.EndPoint.X
                    Dim AlignEndY As Double = Align.EndPoint.Y
                    Try
                        Align.PointLocation(PKStartAlign, 0, AlignStartX, AlignStartY)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                    End Try
                    Try
                        Align.PointLocation(PKEndAlign, 0, AlignEndX, AlignEndY)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '================================================================================================
                    'зАПРАШИВАЕМ ТОЧНОСТЬ
                    Dim RoundLineZn As Integer = FormRoadPlane.NumericUpDown6.Value 'точность линейных объектов
                    Dim RoundAngleZn As Integer = FormRoadPlane.NumericUpDown1.Value 'точность углов
                    'устанавливаем системную переменную указаннаю пользователем
                    Dim boolSist As Boolean = False
                    'смотрим точность системной переменной до запуска программы и сохраняем ее
                    Dim SistPerem As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value
                    If SistPerem <> RoundLineZn Then 'устанавливаем новое значение
                        CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value = RoundLineZn
                        boolSist = True
                    End If
                    '===============================================================================
                    'ЗАПРАШИВАЕМ ФОРМАТ ВЫВОДА расстояний
                    Dim indFormatPK As Integer = FormRoadPlane.ComboBox2.SelectedIndex
                    '===============================================================================
                    'ЗАПРАШИВАЕМ ФОРМАТ ВЫВОДА УГЛОВ
                    Dim indFormatAngle As Integer = FormRoadPlane.ComboBox1.SelectedIndex
                    '=========================================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'КОПИРОВАНИЕ СТИЛЕЙ
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim NameBlkVUP As String = "Вершина угла поворота трассы"
                    Dim NameBlkRubPK As String = "Рубленный пикет"
                    Dim NameBlkPK As String = "Пикет"
                    Dim NameBlkKM As String = "Километровый знак"
                    Dim NametypeLineTH As String = "Тангенциальный ход"
                    Dim NametypeLineRazm As String = "Размеры"
                    Dim ArrayPropBlkUPK As String(,) = Nothing
                    Dim ArrayPropBlkRPK As String(,) = Nothing
                    Dim ArrayPropBlkPK As String(,) = Nothing
                    Dim ArrayPropBlkKM As String(,) = Nothing
                    Dim ArrayPropLinTH As String(,) = Nothing
                    Dim ArrayPropLinPTH As String(,) = Nothing
                    Dim IDMLeaderStyleCoord As ObjectId = acCurDb.MLeaderstyle
                    Dim LayerMLeaderCoord As String = "0"
                    Dim ArrayPropertiesMleaderCoord As String(,) = Nothing

                    Dim IDMLeaderStyleVU As ObjectId = acCurDb.MLeaderstyle
                    Dim LayerMLeaderVU As String = "0"
                    Dim ArrayPropertiesMleaderVU As String(,) = Nothing

                    Dim IDMLeaderStyleCurve As ObjectId = acCurDb.MLeaderstyle
                    Dim LayerMLeaderCurve As String = "0"
                    Dim ArrayPropertiesMleaderCurve As String(,) = Nothing

                    Dim IDDimtyleDraw As ObjectId = acCurDb.Dimstyle
                    Dim LayerDim As String = "0"
                    Dim ArrayPropertiesDim As String(,) = Nothing
                    Dim IDTextStyleDraw As ObjectId = acCurDb.Textstyle
                    Dim HText As Double = 2
                    Dim offsetTextLin As Double = 2
                    '====================================================================================================================
                    'Dim arrayInsertObjectToDwg As String(,) = Nothing
                    'Dim countArrayInsertObjectToDwg As Integer = 0
                    'ReDim Preserve arrayInsertObjectToDwg(1, countArrayInsertObjectToDwg)
                    'arrayInsertObjectToDwg(0, 0) = "Координаты"
                    'arrayInsertObjectToDwg(1, 0) = "MLeaderType"
                    'arrayInsertObjectToDwg(0, 1) = "Элементы кривой трассы"
                    'arrayInsertObjectToDwg(1, 1) = "MLeaderType"
                    'countArrayInsertObjectToDwg += 1
                    '==================================================================================================================
                    'файл-шаблон со стилями
                    Dim PatchTemplateDWG As String = FormRoadPlane.PatchFileTemplate
                    Dim patchTemplateXML As String = PatchTemplateDWG.Replace(".dwg", ".xml")
                    If Not (File.Exists(PatchTemplateDWG) = True And File.Exists(patchTemplateXML) = True) Then
                        MsgBox("Шаблон не найден. Графика не построена!!!")
                        Exit Sub
                    End If
                    If FormRoadPlane.CheckBox1.Checked = True Then
                        Dim mlstyles As DBDictionary = tr.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForWrite)
                        Dim acDimStyleTbl As DimStyleTable = tr.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead)
                        Dim acLyrTbl As LayerTable = tr.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite)
                        '==========================================================================================================================
                        'СТИЛЬ МУЛЬТИВЫНОСОК ДЛЯ КООРДИНИНАТ И ЭЛЕМЕНТОВ КРИВОЙ
                        'для координат
                        'читаем характеристики стиля мультивыноски
                        Dim boolReadStyleML As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Мультивыноска", "Координаты", ArrayPropertiesMleaderCoord)
                        If IsArray(ArrayPropertiesMleaderCoord) = True Then
                            For i As Integer = 0 To ArrayPropertiesMleaderCoord.GetUpperBound(1)
                                If ArrayPropertiesMleaderCoord(0, i) Like "Name" Then
                                    Dim NameMleaderStyleCoord As String = ArrayPropertiesMleaderCoord(1, i)
                                    If IsNothing(NameMleaderStyleCoord) = False Then
                                        If NameMleaderStyleCoord.Trim.Length > 0 Then
                                            Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PatchTemplateDWG, NameMleaderStyleCoord)
                                            For Each item As DBDictionaryEntry In mlstyles
                                                If NameMleaderStyleCoord Like item.Key Then
                                                    IDMLeaderStyleCoord = item.Value
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                ElseIf ArrayPropertiesMleaderCoord(0, i) Like "Layer" Then
                                    Dim NameLayerTemp As String = ArrayPropertiesMleaderCoord(1, i)
                                    If IsNothing(NameLayerTemp) = False Then
                                        If NameLayerTemp.Trim.Length > 0 Then
                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                                LayerMLeaderCoord = NameLayerTemp
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If

                        'читаем характеристики стиля мультивыноски
                        Dim boolReadStyleMLVU As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Мультивыноска", "Элементы кривой трассы", ArrayPropertiesMleaderVU)
                        If IsArray(ArrayPropertiesMleaderVU) = True Then
                            For i As Integer = 0 To ArrayPropertiesMleaderVU.GetUpperBound(1)
                                If ArrayPropertiesMleaderVU(0, i) Like "Name" Then
                                    Dim NameMleaderStyleVU As String = ArrayPropertiesMleaderVU(1, i)
                                    If IsNothing(NameMleaderStyleVU) = False Then
                                        If NameMleaderStyleVU.Trim.Length > 0 Then
                                            Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PatchTemplateDWG, NameMleaderStyleVU)
                                            For Each item As DBDictionaryEntry In mlstyles
                                                If NameMleaderStyleVU Like item.Key Then
                                                    IDMLeaderStyleVU = item.Value
                                                    Exit For
                                                End If
                                            Next
                                        End If

                                    End If
                                ElseIf ArrayPropertiesMleaderCoord(0, i) Like "Layer" Then
                                    Dim NameLayerTemp As String = ArrayPropertiesMleaderVU(1, i)
                                    If IsNothing(NameLayerTemp) = False Then
                                        If NameLayerTemp.Trim.Length > 0 Then
                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                                LayerMLeaderVU = NameLayerTemp
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                        'для подписи элементов кривой
                        'читаем характеристики стиля мультивыноски
                        Dim boolReadStyleCurve As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Мультивыноска", "Пояснительная надпись", ArrayPropertiesMleaderCurve)
                        If IsArray(ArrayPropertiesMleaderCurve) = True Then
                            For i As Integer = 0 To ArrayPropertiesMleaderCurve.GetUpperBound(1)
                                If ArrayPropertiesMleaderCurve(0, i) Like "Name" Then
                                    Dim NameMleaderStyleCurve As String = ArrayPropertiesMleaderCurve(1, i)
                                    If IsNothing(NameMleaderStyleCurve) = False Then
                                        If NameMleaderStyleCurve.Trim.Length > 0 Then
                                            Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PatchTemplateDWG, NameMleaderStyleCurve)
                                            For Each item As DBDictionaryEntry In mlstyles
                                                If NameMleaderStyleCurve Like item.Key Then
                                                    IDMLeaderStyleCurve = item.Value
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                ElseIf ArrayPropertiesMleaderCurve(0, i) Like "Layer" Then
                                    Dim NameLayerTemp As String = ArrayPropertiesMleaderCurve(1, i)
                                    If IsNothing(NameLayerTemp) = False Then
                                        If NameLayerTemp.Trim.Length > 0 Then
                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                                LayerMLeaderCurve = NameLayerTemp
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'РАЗМЕРНЫЙ СТИЛЬ
                        '============================================================================================================================
                        Dim acDimStyleCollection As DimStyleTable = tr.GetObject(acCurDb.DimStyleTableId, OpenMode.ForRead)
                        'читаем характеристики стиля мультивыноски
                        Dim boolReadStyleDim As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Размер", "Размеры", ArrayPropertiesDim)
                        If IsArray(ArrayPropertiesDim) = True Then
                            For i As Integer = 0 To ArrayPropertiesDim.GetUpperBound(1)
                                If ArrayPropertiesDim(0, i) Like "Name" Then
                                    Dim NameStyleDim As String = ArrayPropertiesDim(1, i)
                                    If IsNothing(NameStyleDim) = False Then
                                        If NameStyleDim.Trim.Length > 0 Then
                                            Dim boolCopyStyleDim As Boolean = AcadFunction.FuncCreateDimStyleDestDraw(PatchTemplateDWG, NameStyleDim)
                                            If acDimStyleCollection.Has(NameStyleDim) Then
                                                IDDimtyleDraw = acDimStyleCollection.Item(NameStyleDim)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                ElseIf ArrayPropertiesDim(0, i) Like "Layer" Then
                                    Dim NameLayerTemp As String = ArrayPropertiesDim(1, i)
                                    If IsNothing(NameLayerTemp) = False Then
                                        If NameLayerTemp.Trim.Length > 0 Then
                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                            If acLyrTbl.Has(NameLayerTemp) = True Then
                                                LayerDim = NameLayerTemp
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                        'СТИЛЬ ТЕКСТОВ
                        Dim acDimStyleRec As DimStyleTableRecord = tr.GetObject(IDDimtyleDraw, OpenMode.ForRead)
                        IDTextStyleDraw = acDimStyleRec.Dimtxsty
                        'ВСТАВЛЯЕМ НEОБХОДИМЫЕ БЛОКИ
                        '=======================================================================
                        ''вставляем определение блока в базу (вершину угла)
                        If FormRoadPlane.CheckBox1.Checked = True Then
                            If IO.File.Exists(patchTemplateXML) = True Then
                                Dim boolPropBlkUPK As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Блок", NameBlkVUP, ArrayPropBlkUPK)
                                If IsArray(ArrayPropBlkUPK) = True Then
                                    For i As Integer = 0 To ArrayPropBlkUPK.GetUpperBound(1)
                                        Dim NameField As String = ArrayPropBlkUPK(0, i)
                                        If NameField.Trim Like "Name" Then
                                            Dim NameBlk As String = ArrayPropBlkUPK(1, i)
                                            If IsNothing(NameBlk) = False Then
                                                If NameBlk.Trim.Length > 0 Then
                                                    Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PatchTemplateDWG, NameBlk)
                                                    NameBlkVUP = NameBlk
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                                'блок пикета
                                Dim boolPropBlkPK As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Блок", NameBlkPK, ArrayPropBlkPK)
                                If IsArray(ArrayPropBlkPK) = True Then
                                    For i As Integer = 0 To ArrayPropBlkPK.GetUpperBound(1)
                                        Dim NameField As String = ArrayPropBlkPK(0, i)
                                        If NameField.Trim Like "Name" Then
                                            Dim NameBlk As String = ArrayPropBlkPK(1, i)
                                            If IsNothing(NameBlk) = False Then
                                                If NameBlk.Trim.Length > 0 Then
                                                    Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PatchTemplateDWG, NameBlk)
                                                    NameBlkPK = NameBlk
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                                'блок километра
                                Dim boolPropBlkKM As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Блок", NameBlkKM, ArrayPropBlkKM)
                                If IsArray(ArrayPropBlkKM) = True Then
                                    For i As Integer = 0 To ArrayPropBlkKM.GetUpperBound(1)
                                        Dim NameField As String = ArrayPropBlkKM(0, i)
                                        If NameField.Trim Like "Name" Then
                                            Dim NameBlk As String = ArrayPropBlkKM(1, i)
                                            If IsNothing(NameBlk) = False Then
                                                If NameBlk.Trim.Length > 0 Then
                                                    Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PatchTemplateDWG, NameBlk)
                                                    NameBlkKM = NameBlk
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                                'рубленный пикет
                                If Align.StationEquations.Count > 0 Then
                                    Dim boolPropBlkRPK As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Блок", NameBlkRubPK, ArrayPropBlkRPK)
                                    If IsArray(ArrayPropBlkRPK) = True Then
                                        For i As Integer = 0 To ArrayPropBlkRPK.GetUpperBound(1)
                                            Dim NameField As String = ArrayPropBlkRPK(0, i)
                                            If NameField.Trim Like "Name" Then
                                                Dim NameBlk As String = ArrayPropBlkRPK(1, i)
                                                If IsNothing(NameBlk) = False Then
                                                    If NameBlk.Trim.Length > 0 Then
                                                        Dim boolInsRefBlkPK As Boolean = AcadFunction.FuncCopyBlockDestDraw(PatchTemplateDWG, NameBlk)
                                                        NameBlkRubPK = NameBlk
                                                        Exit For
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Layer" Then
                                                Dim NameLayerTemp As String = ArrayPropLinTH(1, i)
                                                If IsNothing(NameLayerTemp) = False Then
                                                    If NameLayerTemp.Trim.Length > 0 Then
                                                        Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                                        If acLyrTbl.Has(NameLayerTemp) = True Then
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next i
                                    End If
                                End If
                                'вставляем тип линии Тангенциальный ход
                                Dim boolPropLinTH As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Полилиния", NametypeLineTH, ArrayPropLinTH)
                                If IsArray(ArrayPropLinTH) = True Then
                                    For i As Integer = 0 To ArrayPropLinTH.GetUpperBound(1)
                                        Dim NameField As String = ArrayPropLinTH(0, i)
                                        If NameField.Trim Like "Name" Then
                                            Dim NameLine As String = ArrayPropLinTH(1, i)
                                            If IsNothing(NameLine) = False Then
                                                If NameLine.Trim.Length > 0 Then
                                                    Dim boolTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PatchTemplateDWG, NameLine)
                                                    NametypeLineTH = NameLine
                                                    Exit For
                                                End If
                                            End If
                                        ElseIf NameField.Trim Like "Layer" Then
                                            Dim NameLayerTemp As String = ArrayPropLinTH(1, i)
                                            If IsNothing(NameLayerTemp) = False Then
                                                If NameLayerTemp.Trim.Length > 0 Then
                                                    Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayerTemp)
                                                    If acLyrTbl.Has(NameLayerTemp) = True Then
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If

                                Dim boolPropLinPTH As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(patchTemplateXML, "Проектирование автомобильных дорог", "Трасса автомобильной дороги", "Полилиния", NametypeLineRazm, ArrayPropLinPTH)
                                If IsArray(ArrayPropLinPTH) = True Then
                                    For i As Integer = 0 To ArrayPropLinPTH.GetUpperBound(1)
                                        Dim NameField As String = ArrayPropLinPTH(0, i)
                                        If NameField.Trim Like "Name" Then
                                            Dim NameLine As String = ArrayPropLinPTH(1, i)
                                            If IsNothing(NameLine) = False Then
                                                If NameLine.Trim.Length > 0 Then
                                                    Dim boolTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PatchTemplateDWG, NameLine)
                                                    NametypeLineRazm = NameLine
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '=============================================================
                    'закончили ввод исходных данных
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '===================================================================================================================
                    'записываем X и Y начала и конца трассы
                    Dim Xstart As Double = Math.Round(AlignStartY, 3)
                    Dim Ystart As Double = Math.Round(AlignStartX, 3)
                    Dim Xend As Double = Math.Round(AlignEndY, 3)
                    Dim Yend As Double = Math.Round(AlignEndX, 3)
                    Dim startNT As Point2d = New Point2d(Xstart, Ystart)
                    Dim endKT As Point2d = New Point2d(Xend, Yend)
                    '====================================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim ArrayEntity As Double(,) = Nothing 'массив для хранения элементов трассы (0 - пикет начала, 1- пикет конца, 2,3-центральная точка для дуги центр окружности для спирали точка пересечения тангенсов, 4-тип элемента 0-линия 1 дуга 2 спираль)
                    '(0-ПК НАЧАЛА, 1-ПК КОНЦА, 2 - X центра дуги (SPI для спирали), 3 - Y центра дуги (SPI для спирали), 4- тип объекта (0-линия,1-дуга,2-спираль), 5 - X начала элемента, 6 - Y начала элемента, 7 - X конца элемента, 8 - Y конца элемента, 9 - пк SPI или центра дуги, 10 - радиус дуги или спирали, 11 - длина объекта, 12 - направление объекта (дир.угол)
                    Dim CountObject As Integer = 0
                    '============================================================
                    'перебираем все элементы трассы для их сортировки в процессе возрастания
                    For Each myAe As AlignmentEntity In Align.Entities
                        Dim AlignEnt As AlignmentEntityType = myAe.EntityType
                        Select Case AlignEnt
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'прямой участок
                            Case AlignmentEntityType.Line 'прямой участок
                                Dim AlignLine As AlignmentLine = myAe
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsLine(Align, AlignLine, ArrayEntity, PKStartAlign, PKEndAlign)
     '===========================================================================================================================
'дуга одиночная
                            Case AlignmentEntityType.Arc
                                Dim AlignArc As AlignmentArc = myAe
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsArc(Align, AlignArc, ArrayEntity, PKStartAlign, PKEndAlign)
'========================================================================================================================================
'спираль одиночная
                            Case AlignmentEntityType.Spiral
                                If TypeOf myAe Is AlignmentSTS Then
                                    Dim AlignSTS As AlignmentSTS = myAe 'составная спираль спираль дуга спираль
                                    Dim AlignS1 As AlignmentSubEntitySpiral = Nothing
                                    Try
                                        AlignS1 = AlignSTS.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        AlignS1 = AlignSTS.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    'если спираль существует
                                    If IsNothing(AlignS1) = False Then
                                        Dim boolSpcChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, AlignS1, ArrayEntity, PKStartAlign, PKEndAlign)
                                    End If
                                    '==================================================================================================================================
                                ElseIf TypeOf myAe Is AlignmentSCS Then
                                    Dim AlignSCS As AlignmentSCS = myAe 'составная спираль спираль прямая спираль
                                    Dim AlignS2 As AlignmentSubEntitySpiral = Nothing
                                    Try
                                        AlignS2 = AlignSCS.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        AlignS2 = AlignSCS.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    'если спираль существует
                                    If IsNothing(AlignS2) = False Then
                                        Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, AlignS2, ArrayEntity, PKStartAlign, PKEndAlign)
                                    End If
                                    '=================================================================================================
                                ElseIf TypeOf myAe Is AlignmentSpiral Then
                                    Dim AlignS3 As AlignmentSpiral = Nothing 'отдельная спираль
                                    Try
                                        AlignS3 = myAe
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSpiral(Align, AlignS3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'=================================================================================================
'спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiral
                                Dim AlignSCS As AlignmentSCS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSCS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignSCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'=============================================================================================================================
'переходная кривая прямая переходная кривая
                            Case AlignmentEntityType.SpiralLineSpiral
                                Dim AlignSLS As AlignmentSTS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSLS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignSLS.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(Align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignSLS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'=========================================================================================================================
'переходная кривая прямая
                            Case AlignmentEntityType.SpiralLine
                                Dim AlignSL As AlignmentSTS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSL.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignSL.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(Align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'======================================================================================================================
'прямая спираль 
                            Case AlignmentEntityType.LineSpiral
                                Dim AlignLS As AlignmentSTS = myAe
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignLS.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(Align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignLS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'===================================================================================================================================
'переходная кривая дуга 
                            Case AlignmentEntityType.SpiralCurve
                                Dim AlignSC As AlignmentSCS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSC.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSC.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==========================================================================================================================================
                            Case AlignmentEntityType.CurveSpiral 'кривая переходная кривая
                                Dim AlignCS As AlignmentSCS = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = AlignCS.SpiralOut
                                Try
                                    SpOut = AlignCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==============================================================================================================================================
'переходная кривая переходная кривая дуга переходная кривая переходная кривая
                            Case AlignmentEntityType.SpiralSpiralCurveSpiralSpiral
                                Dim AlignSSCSS As AlignmentSSCSS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSSCSS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSSCSS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSSCSS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSSCSS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignSSCSS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'====================================================================================================================================
'дуга дуга обратная дуга
                            Case AlignmentEntityType.CurveCurveReverseCurve
                                Dim AlignCCRC As AlignmentCCRC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCCRC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCCRC.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc3 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc3 = AlignCCRC.Arc3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc3) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==============================================================================================================================
'дуга обратная дуга
                            Case AlignmentEntityType.CurveReverseCurve
                                Dim AlignCRC As AlignmentCRC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCRC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCRC.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==================================================================================================================================
'кривавая спираль спираль
                            Case AlignmentEntityType.CurveSpiralSpiral
                                Dim AlignCSS As AlignmentSSCSS = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCSS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignCSS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignCSS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==========================================================================================================================
'спираль спираль
                            Case AlignmentEntityType.SpiralSpiral
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                If TypeOf myAe Is AlignmentSTS Then
                                    Dim AlignCS1 As AlignmentSTS = myAe
                                    Try
                                        Sp1 = AlignCS1.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        Sp2 = AlignCS1.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                ElseIf TypeOf myAe Is AlignmentSCS Then
                                    Dim AlignCS2 As AlignmentSCS = myAe
                                    Try
                                        Sp1 = AlignCS2.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        Sp2 = AlignCS2.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                End If
                                If IsNothing(Sp1) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                If IsNothing(Sp2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'=============================================================================================================================
'спираль спираль дуга
                            Case AlignmentEntityType.SpiralSpiralCurve
                                Dim AlignSSC As AlignmentSSCSS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSSC.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSSC.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSSC.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'================================================================================================================================
'спираль дуга спираль спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiralSpiralCurveSpiral
                                Dim AlignSCSSCS As AlignmentSCSSCS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSCSSCS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCSSCS.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSCSSCS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSCSSCS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignSCSSCS.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignSCSSCS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
'==================================================================================================================================================
'спираль дуга спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiralCurveSpiral
                                Dim AlignSCSCS As AlignmentSCSCS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSCSCS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCSCS.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSCSCS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignSCSCS.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSCSCS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(Align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
 '=================================================================================================================                            
                            Case AlignmentEntityType.MultipleSegments
                                Dim AlignCS As AlignmentMultipleSegments = myAe
                                                        'структура не понятна
'=================================================================================================================
                            Case AlignmentEntityType.CurveLineCurve
                                Dim AlignCLC As AlignmentCTC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = AlignCLC.Arc1
                                Try
                                    Arc1 = AlignCLC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignCLC.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(Align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCLC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(Align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                        End Select
                    Next
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'второй этап
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'сортируем массив
                    If IsArray(ArrayEntity) = True Then
                        If ArrayEntity.GetUpperBound(1) > 0 Then
                            Dim boolArray As Boolean = MathFunction.FuncSortDblArray(ArrayEntity, 0)
                        End If
                    Else
                        MsgBox("Сегменты трассы не удалось записать в массив. Макро прервано!!!")
                        Exit Sub
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'ищем позицию первого и последнего вхождения для кривых
                    Dim i1 As Integer = -1 'позиция для кривой начала трассы
                    Dim i2 As Integer = -1 'позиция ждля кривой конца трассы
                    For i As Integer = 0 To ArrayEntity.GetUpperBound(1)
                        Dim NKr As Integer = ArrayEntity(4, i)
                        If NKr > 0 Then
                            If i1 = -1 Then
                                i1 = i
                            End If
                            i2 = i
                        End If
                    Next i
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'третий этап (основной)
                    '=======================================================================================================
                    'начинаем перебирать элементы массива
                    Dim NumbVU As Integer = 0 'номер вершины угла
                    Dim ArrayExcel As Double(,) = Nothing 'результирующий массив для Excel
                    Dim NTZY As Double = -1 'координаты начала трассы
                    Dim KTXY As Double = -1 'координаты конца трассы
                    Dim TLenghtStart As Double = 0 'длина прямой вставки (линии)
                    Dim TLenghtEnd As Double = 0 'длина прямой вставки (линии)
                    Dim countArrayExcel As Integer = 0
                    For i As Integer = 0 To ArrayEntity.GetUpperBound(1)
                        '====================================================================================
                        'начальные данные
                        Dim StartPoint As Point2d = Nothing 'точка начала обработки конструкции
                        Dim EndPoint As Point2d = Nothing 'точка конца обработки конструкции
                        Dim StartPK As Double = 0 'пикет начала обработки конструкции
                        Dim EndPK As Double = 0 'пикет конца обработки конструкции
                        Dim VUxy As Point2d = Nothing 'координаты вершины угла
                        Dim PKVU As Double = 0 'пикет вершины угла
                        Dim DeltaAngle As Double = 0 'дельта- угол
                        Dim R As Double = 0 'радиус
                        Dim L1 As Double = 0 'длина переходной кривой
                        Dim L2 As Double = 0 'длина переходной кривой
                        Dim T1 As Double = 0 'длина тангенса
                        Dim T2 As Double = 0 'длина тангенса
                        Dim K As Double = 0 'длина кривой
                        Dim B As Double = 0 'биссектриса
                        Dim D As Double = 0 'домер
                        Dim Ug As Double = 0 'угол поворота
                        Dim NPK As Double = -10000 'пикет начала переходной кривой
                        Dim NKK As Double = -10000 'пикет начала дуги
                        Dim KKK As Double = -10000 'пикет конца дуги
                        Dim KPK As Double = -10000 'пикет конца переходной кривой
                        Dim SK As Double = 0 'середина дуги
                        Dim boolFindSpLeft As Boolean = False 'флаг что спираль слева сонаправлена
                        Dim boolFindSpRight As Boolean = False 'флан сто спираль справа сонаправлена
                        '========================================================================================
                        'вспомогательные линии для построения
                        Dim Line1 As Line = Nothing
                        Dim Line2 As Line = Nothing
                        '========================================================================================
                        'считываем тип объекта
                        Dim TypeObj As Integer = ArrayEntity(4, i)
                        '==========================================================================================
                        'обрабатываем только линии, дуги и переходные кривые
                        '(0-ПК НАЧАЛА, 1-ПК КОНЦА, 2 - X центра дуги (SPI для спирали), 3 - Y центра дуги (SPI для спирали), 4- тип объекта (0-линия,1-дуга,2-спираль), 5 - X начала элемента, 6 - Y начала элемента, 7 - X конца элемента, 8 - Y конца элемента, 9 - пк SPI или центра дуги, 10 - радиус дуги или спирали (входящий),11 - радиус спирали (выходящий), 12 - длина объекта, 13 - направление объекта (дир.угол)
                        '============================================================================================
                        'находим дугу
                        If TypeObj = 1 Then 'это дуга
                            NKK = ArrayEntity(0, i) 'берем пикет начала и пикет конца дуги
                            KKK = ArrayEntity(1, i)
                            '=============================================================================
                            'считываем характеристики
                            '=============================================================================
                            'координаты начала и конца круговой кривой
                            StartPoint = New Point2d(ArrayEntity(5, i), ArrayEntity(6, i))
                            EndPoint = New Point2d(ArrayEntity(7, i), ArrayEntity(8, i))
                            'записываем координаты начала, конца и центра дуги в массив
                            Dim ArrayArc As Double() = Nothing
                            ReDim ArrayArc(8)
                            ArrayArc(0) = ArrayEntity(5, i) 'начало дуги
                            ArrayArc(1) = ArrayEntity(6, i)
                            ArrayArc(2) = ArrayEntity(7, i) 'конец дуги
                            ArrayArc(3) = ArrayEntity(8, i)
                            ArrayArc(4) = ArrayEntity(2, i) 'центр дуги
                            ArrayArc(5) = ArrayEntity(3, i)
                            'дублируем центр дуги
                            ArrayArc(7) = ArrayEntity(2, i) 'центр дуги
                            ArrayArc(8) = ArrayEntity(3, i)
                            SK = (NKK + KKK) / 2 'пикет середины дуги
                            ArrayArc(6) = SK
                            If IsArray(ArrayArc) = True Then
                                '======================================================================================
                                'смотрим элемент слева
                                Dim boolSPArcIn As Boolean = False 'флаг для элемента слева (элемент сонаправлен)
                                Dim boolSPArcOut As Boolean = False 'флаг для соседа справа (элемент сонаправлен)
                                'смотрим соседа слева
                                '=======================================================================================
                                'случай если элемент стоит не первым
                                If i > 0 Then
                                    'определяем тип объекта
                                    'обрабатываем только спирали
                                    Dim TypeObj1 As Integer = ArrayEntity(4, i - 1) 'тип объекта дуги или спираль
                                    If TypeObj1 = 2 Then 'если спираль тогда 
                                        'это спираль проверяем ее на сонаправленность к дуге
                                        Dim StSP As Double = ArrayEntity(0, i - 1) 'пикет начала переходной кривой
                                        Dim EndSP As Double = ArrayEntity(1, i - 1) 'пикет конца переходной кривой
                                        Dim StSp3d As Point3d = New Point3d(ArrayEntity(5, i - 1), ArrayEntity(5, i - 1), 0) 'координаты начала переходной кривой
                                        Dim SPI3d As Point3d = New Point3d(ArrayEntity(2, i - 1), ArrayEntity(3, i - 1), 0) 'координаты SPI точки
                                        Dim ArraySP As Double() = Nothing 'массив с координатами начала и конца переходной кривой
                                        ReDim ArraySP(8)
                                        ArraySP(0) = ArrayEntity(5, i - 1) 'начало переходной кривой
                                        ArraySP(1) = ArrayEntity(6, i - 1)
                                        ArraySP(2) = ArrayEntity(7, i - 1) 'конец переходной кривой
                                        ArraySP(3) = ArrayEntity(8, i - 1)
                                        ArraySP(4) = ArrayEntity(2, i - 1) 'координаты SpI точки
                                        ArraySP(5) = ArrayEntity(3, i - 1)
                                        ArraySP(6) = ArrayEntity(9, i - 1) 'пикет точки SPI
                                        ArraySP(7) = ArrayEntity(14, i - 1)
                                        ArraySP(8) = ArrayEntity(15, i - 1) 'пикет точки SPI
                                        'если массив состоялся
                                        If IsArray(ArraySP) = True Then
                                            'проверяем на сонаправленность входнцю спираль 
                                            boolSPArcIn = FuncAlignment.FuncSpiralArcNapr1(Align, ArraySP, ArrayArc) 'функция проверяет сонаправленность спирали и дуги
                                            'спираль слева сонаправлена с дугой
                                            If boolSPArcIn = True Then
                                                'строим вспомогательные линии для определения вершины угла поворота
                                                Dim StLine1 As Point3d = New Point3d(ArraySP(0), ArraySP(1), 0) 'координаты точки начала переходной кривой
                                                Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i - 1), ArrayEntity(3, i - 1), 0) 'координаты точки SPI переходной кривой
                                                Line1 = New Line(StLine1, StLine2) 'первая вспомогательная линия
                                                StartPoint = New Point2d(ArraySP(0), ArraySP(1)) 'координаты стартовой точки
                                                StartPK = StSP ''старт ПК начала обработки конструкции
                                                L1 = Math.Abs(EndSP - StSP) 'записываем длину первой переходной кривой
                                                NPK = StSP 'записываем пикет начала переходной кривой
                                            Else
                                                'переходная кривая не сонаправлена с данной дугой - считаем эту дугу первым элементом
                                                'ищем вектор центр дуги - начальный пикет
                                                Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                                Dim NK2d As Point2d = New Point2d(ArrayArc(0), ArrayArc(1)) 'точка - начало дуги
                                                Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d) 'вектор центр дуги - середина дуги
                                                Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2 'поворачиваем на 90 градусов
                                                If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                                '2.2 находим приближенные координаты на 1 тангенс
                                                Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                                Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                                Dim X As Double = ArrayArc(0) + dx
                                                Dim Y As Double = ArrayArc(1) + dy
                                                'строим вторую вспомогательную линию
                                                Dim StLine1 As Point3d = New Point3d(ArrayArc(0), ArrayArc(1), 0)
                                                Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                                Line1 = New Line(StLine1, StLine2)
                                            End If
                                        End If
                                    Else
                                        'дуга это первый элемент соседей слева нет
                                        'дуга слева соседа не имеет ищем направление для тангенса
                                        'ищем вектор центр дуги - начальный пикет
                                        Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                        Dim NK2d As Point2d = New Point2d(ArrayArc(0), ArrayArc(1)) 'точка - начало дуги
                                        Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d)
                                        Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2
                                        If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                        '2.2 находим приближенные координаты на 1 тангенс
                                        Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                        Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                        Dim X As Double = ArrayArc(0) + dx
                                        Dim Y As Double = ArrayArc(1) + dy
                                        'строим первую вспомогательную линию
                                        Dim StLine1 As Point3d = New Point3d(ArrayArc(0), ArrayArc(1), 0)
                                        Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                        Line1 = New Line(StLine1, StLine2)
                                    End If
                                Else
                                    'дуга первый элемент
                                    'дуга слева соседа не имеет ищем направление для тангенса
                                    'ищем вектор центр дуги - начальный пикет
                                    Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                    Dim NK2d As Point2d = New Point2d(ArrayArc(0), ArrayArc(1)) 'точка - начало дуги
                                    Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d)
                                    Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2
                                    If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                    '2.2 находим приближенные координаты на 1 тангенс
                                    Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                    Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                    Dim X As Double = ArrayArc(0) + dx
                                    Dim Y As Double = ArrayArc(1) + dy
                                    Dim StLine1 As Point3d = New Point3d(ArrayArc(0), ArrayArc(1), 0)
                                    Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                    'строим первую вспомогательную линию
                                    Line1 = New Line(StLine1, StLine2)
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                '================================================================================
                                'заходим с обратной стороны смотрим правый элемент
                                If i < ArrayEntity.GetUpperBound(1) Then
                                    'определяем тип объекта
                                    'обрабатываем только спирали
                                    Dim TypeObj1 As Integer = ArrayEntity(4, i + 1)
                                    If TypeObj1 = 2 Then
                                        'это спираль проверяем ее на сонаправленность к дуге
                                        Dim StSP As Double = ArrayEntity(0, i + 1) 'пикет начала переходной кривой
                                        Dim EndSP As Double = ArrayEntity(1, i + 1) 'пикет конца переходной кривой
                                        Dim StSp3d As Point3d = New Point3d(ArrayEntity(5, i + 1), ArrayEntity(5, i + 1), 0) 'координаты начала переходной кривой
                                        Dim SPI3d As Point3d = New Point3d(ArrayEntity(2, i + 1), ArrayEntity(3, i + 1), 0) 'координаты SPI точки
                                        Dim ArraySP As Double() = Nothing 'массив с координатами начала и конца переходной кривой
                                        ReDim ArraySP(8)
                                        ArraySP(0) = ArrayEntity(5, i + 1) 'начало переходной кривой
                                        ArraySP(1) = ArrayEntity(6, i + 1)
                                        ArraySP(2) = ArrayEntity(7, i + 1) 'конец переходной кривой
                                        ArraySP(3) = ArrayEntity(8, i + 1)
                                        ArraySP(4) = ArrayEntity(2, i + 1) 'координаты SpI точки
                                        ArraySP(5) = ArrayEntity(3, i + 1)
                                        ArraySP(6) = ArrayEntity(9, i + 1) 'пикет точки SPI
                                        ArraySP(7) = ArrayEntity(14, i + 1)
                                        ArraySP(8) = ArrayEntity(15, i + 1) 'пикет точки SPI
                                        'если массив состоялся
                                        If IsArray(ArraySP) = True Then
                                            'проверяем на сонаправленность
                                            boolSPArcOut = FuncAlignment.FuncArcSpiralNapr1(Align, ArrayArc, ArraySP)
                                            If boolSPArcOut = True Then
                                                'спирали сонаправлены вычисляем элементы кривой
                                                'строим вспомогательные линии для определения вершины угла поворота
                                                Dim StLine1 As Point3d = New Point3d(ArraySP(2), ArraySP(3), 0)
                                                Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i + 1), ArrayEntity(3, i + 1), 0)
                                                'строим вторую вспомогательную линию
                                                Line2 = New Line(StLine1, StLine2)
                                                EndPoint = New Point2d(ArraySP(2), ArraySP(3))
                                                EndPK = EndSP
                                                L2 = Math.Abs(EndSP - StSP) 'записываем длину первой переходной кривой
                                                KPK = EndSP 'записываем пикет конца переходной кривой
                                            Else
                                                'дуга слева соседа не имеет ищем направление для тангенса
                                                'ищем вектор центр дуги - начальный пикет
                                                Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                                Dim NK2d As Point2d = New Point2d(ArrayArc(2), ArrayArc(3)) 'точка - начало дуги
                                                Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d)
                                                Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2
                                                If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                                '2.2 находим приближенные координаты на 1 тангенс
                                                Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                                Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                                Dim X As Double = ArrayArc(2) + dx
                                                Dim Y As Double = ArrayArc(3) + dy
                                                Dim StLine1 As Point3d = New Point3d(ArrayArc(2), ArrayArc(3), 0)
                                                Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                                Line2 = New Line(StLine1, StLine2)
                                            End If
                                        End If
                                    Else
                                        'дуга слева соседа не имеет ищем направление для тангенса
                                        'ищем вектор центр дуги - начальный пикет
                                        Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                        Dim NK2d As Point2d = New Point2d(ArrayArc(2), ArrayArc(3)) 'точка - начало дуги
                                        Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d)
                                        Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2
                                        If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                        '2.2 находим приближенные координаты на 1 тангенс
                                        Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                        Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                        Dim X As Double = ArrayArc(2) + dx
                                        Dim Y As Double = ArrayArc(3) + dy
                                        Dim StLine1 As Point3d = New Point3d(ArrayArc(2), ArrayArc(3), 0)
                                        Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                        Line2 = New Line(StLine1, StLine2)
                                    End If
                                Else
                                    'дуга слева соседа не имеет ищем направление для тангенса
                                    'ищем вектор центр дуги - начальный пикет
                                    Dim CP2d As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'точка - центр дуги
                                    Dim NK2d As Point2d = New Point2d(ArrayArc(2), ArrayArc(3)) 'точка - начало дуги
                                    Dim Vect_CP_NK As Vector2d = CP2d.GetVectorTo(NK2d)
                                    Dim DirectionT1 As Double = Vect_CP_NK.Angle - Math.PI / 2
                                    If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                    '2.2 находим приближенные координаты на 1 тангенс
                                    Dim dx As Double = Math.Cos(DirectionT1) * Vect_CP_NK.Length
                                    Dim dy As Double = Math.Sin(DirectionT1) * Vect_CP_NK.Length
                                    Dim X As Double = ArrayArc(2) + dx
                                    Dim Y As Double = ArrayArc(3) + dy
                                    Dim StLine1 As Point3d = New Point3d(ArrayArc(2), ArrayArc(3), 0)
                                    Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                    Line2 = New Line(StLine1, StLine2)
                                End If
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'делаем анализ конструкции в состав которой входит дуга
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'вычисляем вершину угла
                            If IsNothing(Line1) = False And IsNothing(Line2) = False Then
                                'ищем точку пересечения двух линий
                                Dim PTVUxyz As Point3d = Nothing 'координата вершины дуги 3д
                                Dim intersectPoint As Point3dCollection = New Point3dCollection
                                Dim boolInterPoint As Boolean = False
                                Try
                                    Line1.IntersectWith(Line2, Intersect.ExtendBoth, intersectPoint, 1, 1)
                                    If intersectPoint.Count > 0 Then
                                        PTVUxyz = intersectPoint.Item(0)
                                        VUxy = New Point2d(PTVUxyz.X, PTVUxyz.Y)
                                        boolInterPoint = True
                                    End If
                                Catch ex As Exception
                                End Try
                                'точка пересечения тангенсов найдена продолжаем процедуру
                                If intersectPoint.Count = 0 Then
                                    'пытаемся математически найти точку пересечения
                                    Dim pt1 As Point2d = New Point2d(Line1.StartPoint.X, Line1.StartPoint.Y)
                                    Dim pt2 As Point2d = New Point2d(Line1.EndPoint.X, Line1.EndPoint.Y)
                                    Dim pt3 As Point2d = New Point2d(Line2.StartPoint.X, Line2.StartPoint.Y)
                                    Dim pt4 As Point2d = New Point2d(Line2.EndPoint.X, Line2.EndPoint.Y)
                                    boolInterPoint = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, VUxy)
                                    PTVUxyz = New Point3d(VUxy.X, VUxy.Y, 0)
                                End If
                                'находим длины 2 тангенсов
                                Dim Vect_T1 As Vector2d = StartPoint.GetVectorTo(VUxy)
                                Dim Vect_T2 As Vector2d = VUxy.GetVectorTo(EndPoint)
                                T1 = Vect_T1.Length '1 тангенс
                                T2 = Vect_T2.Length '2 тангенс
                                'вычисляем длину кривой
                                Dim LD As Double = ArrayEntity(1, i) - ArrayEntity(0, i) 'длина дуги (пикет конца - пикет начала)
                                K = L1 + L2 + LD 'общая длина дуги
                                D = T1 + T2 - K 'домер
                                'вычисляем отношение тангенса и длины кривой
                                Dim Dotn1 As Double = K / (T1 + T2)
                                'проверяем вешину угла на предмет если придется дугу делить на 2 части
                                'находим координаты середины конструкции
                                Dim SKX As Double = -100000 'координаты середины дуги
                                Dim SKY As Double = -100000
                                Try
                                    Align.PointLocation(SK, 0, SKX, SKY)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    'ошибка
                                    Dim StArc As Point2d = New Point2d(ArrayEntity(5, i), ArrayEntity(6, i))
                                    Dim EndArc As Point2d = New Point2d(ArrayEntity(7, i), ArrayEntity(8, i))
                                    Dim CtArc As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i))
                                    Dim MiddleArc As Point2d = MathFunction.FuncFindMiddleArc(StArc, EndArc, CtArc)
                                    SKX = MiddleArc.X
                                    SKY = MiddleArc.Y
                                End Try
                                'определяем точку середины дуги
                                Dim PtSk As Point2d = New Point2d(SKX, SKY) 'координаты середины дуги
                                'определяем точку центр дуги
                                Dim PtCk As Point2d = New Point2d(ArrayEntity(2, i), ArrayEntity(3, i)) 'центр дуги
                                'делаем вектора
                                Dim Vect_SK_VUxy As Vector2d = PtSk.GetVectorTo(VUxy)
                                Dim Vect_SK_PtCK As Vector2d = PtSk.GetVectorTo(PtCk)
                                'вычисляем радиус дуги
                                R = Vect_SK_PtCK.Length
                                'определяем угол если расхождения в углах более 90 градусов считаем что дуго развернута и ее следует поделить на 2 части
                                Dim DAngle As Double = Math.Abs(Vect_SK_PtCK.Angle - Vect_SK_VUxy.Angle)
                                '========================================================================================
                                'дуга развернута или тангенсы очень большие
                                Dim kT1 As Double = FormRoadPlane.NumericUpDown5.Value
                                If DAngle < Math.PI / 2 Or Dotn1 < kT1 Then
Lineint:

                                    'ищем вектор центр окружности - середина кривой
                                    Dim V_CP_SK As Vector2d = PtCk.GetVectorTo(PtSk)
                                    'строим перпендикуляр
                                    Dim DirectionT1 As Double = V_CP_SK.Angle - Math.PI / 2
                                    If DirectionT1 < 0 Then DirectionT1 = DirectionT1 + 2 * Math.PI
                                    '2.2 находим приближенные координаты на 1 тангенс
                                    Dim dx As Double = Math.Cos(DirectionT1) * V_CP_SK.Length
                                    Dim dy As Double = Math.Sin(DirectionT1) * V_CP_SK.Length
                                    Dim X As Double = PtSk.X + dx
                                    Dim Y As Double = PtSk.Y + dy
                                    Dim StLine1 As Point3d = New Point3d(PtSk.X, PtSk.Y, 0)
                                    Dim StLine2 As Point3d = New Point3d(X, Y, 0)
                                    Dim Line3 As Line = New Line(StLine1, StLine2)
                                    'ищем точки пересечения с линией 1
                                    Dim intersectPoint1 As Point3dCollection = New Point3dCollection
                                    Dim PTVUxy1 As Point2d = Nothing
                                    Dim boolIntPoint1 As Boolean = False
                                    Dim intersectPoint2 As Point3dCollection = New Point3dCollection
                                    Dim PTVUxy2 As Point2d = Nothing
                                    Dim boolIntPoint2 As Boolean = False
                                    Try
                                        Line3.IntersectWith(Line1, Intersect.ExtendBoth, intersectPoint1, 1, 1)
                                        Line3.IntersectWith(Line2, Intersect.ExtendBoth, intersectPoint2, 1, 1)
                                    Catch ex As Exception
                                    End Try
                                    'делаем проверку 1 пересечения
                                    If intersectPoint1.Count = 0 Then
                                        'пытаемся математически найти точку пересечения
                                        Dim pt1 As Point2d = New Point2d(Line1.StartPoint.X, Line1.StartPoint.Y)
                                        Dim pt2 As Point2d = New Point2d(Line1.EndPoint.X, Line1.EndPoint.Y)
                                        Dim pt3 As Point2d = New Point2d(Line3.StartPoint.X, Line3.StartPoint.Y)
                                        Dim pt4 As Point2d = New Point2d(Line3.EndPoint.X, Line3.EndPoint.Y)
                                        boolIntPoint1 = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, PTVUxy1)
                                    Else
                                        PTVUxy1 = New Point2d(intersectPoint1.Item(0).X, intersectPoint1.Item(0).Y)
                                        boolIntPoint1 = True
                                    End If
                                    'делаем проверку 2 пересечения
                                    If intersectPoint2.Count = 0 Then
                                        'пытаемся математически найти точку пересечения
                                        Dim pt1 As Point2d = New Point2d(Line2.StartPoint.X, Line2.StartPoint.Y)
                                        Dim pt2 As Point2d = New Point2d(Line2.EndPoint.X, Line2.EndPoint.Y)
                                        Dim pt3 As Point2d = New Point2d(Line3.StartPoint.X, Line3.StartPoint.Y)
                                        Dim pt4 As Point2d = New Point2d(Line3.EndPoint.X, Line3.EndPoint.Y)
                                        boolIntPoint2 = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, PTVUxy2)
                                    Else
                                        PTVUxy2 = New Point2d(intersectPoint2.Item(0).X, intersectPoint2.Item(0).Y)
                                        boolIntPoint2 = True
                                    End If
                                    '==============================================================================
                                    'координаты вершины угла 1
                                    If boolIntPoint1 = True Then
                                        Dim ldVu1 As Double = startNT.GetVectorTo(PTVUxy1).Length
                                        If ldVu1 < 0.02 Then Continue For
                                        'открываем 2 новые строки массива
                                        ReDim Preserve ArrayExcel(23, NumbVU + 1)
                                        '====================================================================================
                                        'номер 1вершины 
                                        ArrayExcel(0, NumbVU) = NumbVU + 1
                                        '===================================================================================
                                        PTVUxyz = New Point3d(PTVUxy1.X, PTVUxy1.Y, 0)
                                        ''координаты первой вершины в массив для отчета
                                        ArrayExcel(19, NumbVU) = PTVUxy1.X
                                        ArrayExcel(20, NumbVU) = PTVUxy1.Y
                                        'вычисляем характеристики 1 половины дуги
                                        Vect_T1 = StartPoint.GetVectorTo(PTVUxy1) 'начало конструкции - вершина угла
                                        Vect_T2 = PTVUxy1.GetVectorTo(PtSk) 'вершина угла - середина дуги
                                        T1 = Vect_T1.Length '1 тангенс
                                        T2 = Vect_T2.Length '2 тангенс
                                        ArrayExcel(6, NumbVU) = T1
                                        ArrayExcel(7, NumbVU) = T2
                                        '=============================================================================================
                                        'длины переходных кривых
                                        ArrayExcel(4, NumbVU) = L1 'первая половина дуги
                                        ArrayExcel(5, NumbVU) = 0
                                        '========================================================================================
                                        'определяесмя есть или нет переходная кривая в начале конструкции
                                        Dim StartPKKonstr As Double = NPK
                                        If L1 > 0 Then
                                            StartPKKonstr = NPK
                                        Else
                                            StartPKKonstr = NKK
                                        End If
                                        'записываем пикет вершины угла
                                        'находим пикет начала конструкции (с учетом рубленности)
                                        Dim NPKStrPK As String = "-1"
                                        Try
                                            NPKStrPK = Align.GetStationStringWithEquations(StartPKKonstr)
                                        Catch ex As System.ArgumentException
                                        End Try
                                        'переводим строку в число
                                        If Not (NPKStrPK Like "-1") Then
                                            Dim NPKDblPK As Double = Val(NPKStrPK.Replace("+", ""))
                                            'проверяем нахождение начала кривй на рубленном пикете
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, StartPKKonstr, SK)
                                            ArrayExcel(1, NumbVU) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        Else
                                            ArrayExcel(1, NumbVU) = 0 'в случае неудачи
                                        End If
                                        '================================================================================================
                                        'вычисляем длину кривой для первой половины дуги
                                        Dim Ksohr As Double = 0 'длина дуги без учета переходных кривых
                                        Ksohr = SK - NKK 'длина дуги (пикет середины дуги минус пикет начала дуги
                                        K = L1 + Ksohr
                                        ArrayExcel(8, NumbVU) = K
                                        ArrayExcel(9, NumbVU) = Ksohr
                                        '================================================================================================
                                        'определяем биссектрису
                                        Dim LineTX1 As Double = 10 * Math.Cos(Vect_T1.Angle + Math.PI) + PTVUxy1.X
                                        Dim LineTY1 As Double = 10 * Math.Sin(Vect_T1.Angle + Math.PI) + PTVUxy1.Y
                                        Dim LineTX2 As Double = 10 * Math.Cos(Vect_T2.Angle) + PTVUxy1.X
                                        Dim LineTY2 As Double = 10 * Math.Sin(Vect_T2.Angle) + PTVUxy1.Y
                                        Dim LineT1XSr As Double = (LineTX1 + LineTX2) / 2
                                        Dim LineT1YSr As Double = (LineTY1 + LineTY2) / 2
                                        Dim ptLineT1 As Point3d = New Point3d(LineT1XSr, LineT1YSr, 0)
                                        Dim LineB As Line = New Line(PTVUxyz, ptLineT1)
                                        Dim PointCol As Point3dCollection = New Point3dCollection
                                        Align.IntersectWith(LineB, Intersect.ExtendBoth, PointCol, 1, 1)
                                        If PointCol.Count > 0 Then
                                            For atp As Integer = 0 To PointCol.Count - 1
                                                Dim PtB As Point3d = PointCol.Item(atp)
                                                Dim PtB2d As Point2d = New Point2d(PtB.X, PtB.Y)
                                                Dim Vect_B As Vector2d = PtB2d.GetVectorTo(PTVUxy1)
                                                B = Vect_B.Length
                                                ArrayExcel(10, NumbVU) = B
                                                Dim PKB As Double = 0
                                                Dim PKOff As Double = 0
                                                Try
                                                    Align.StationOffset(PtB.X, PtB.Y, PKB, PKOff)
                                                    If PKB > StartPKKonstr And PKB < SK Then
                                                        If Math.Abs(PKOff) < 0.01 Then
                                                            ArrayExcel(23, NumbVU) = PKB
                                                            Exit For
                                                        End If
                                                    End If
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            Next
                                        End If
                                        '==========================================================================
                                        ''вычисляем домер
                                        D = T1 + T2 - K
                                        ArrayExcel(11, NumbVU) = D
                                        ''===================================================================================================
                                        ''вычисляем величину угла дельта для первой вершины
                                        Dim Dugol As Double = 0
                                        ''определяем дирекционный угол 1 тангенса
                                        Dim Vect_NSP_VU As Vector2d = StartPoint.GetVectorTo(PTVUxy1) 'вектор ночало спирали - вершина угла
                                        Dim Vect_VU_KSP As Vector2d = PTVUxy1.GetVectorTo(PtSk) 'вектор вершина угла - конец спирали
                                        Dim Vect_KSP_NSP As Vector2d = StartPoint.GetVectorTo(PtSk) 'вектор конец спирали - начало спирали
                                        ''вычисляем внутренний угол
                                        Dim k1 As Double = Vect_NSP_VU.Length ^ 2 + Vect_VU_KSP.Length ^ 2 - Vect_KSP_NSP.Length ^ 2
                                        Dim k2 As Double = 2 * Vect_NSP_VU.Length * Vect_VU_KSP.Length
                                        Dim ugol As Double = Math.PI - Math.Acos(k1 / k2) 'определили угол доворота
                                        ''делаем проверку
                                        ''прибавляем к направлению 1 тангенса величину угла доворота
                                        Dim Napr1 As Double = Vect_NSP_VU.Angle + ugol
                                        If Napr1 > Math.PI * 2 Then Napr1 = Napr1 - Math.PI * 2
                                        Dim Napr2 As Double = Vect_NSP_VU.Angle - ugol
                                        If Napr2 < 0 Then Napr2 = Napr2 + 2 * Math.PI
                                        ''проверяем
                                        Dim dNapr1 As Double = Math.Round(Vect_VU_KSP.Angle - Napr1, 5)
                                        If dNapr1 < 0 Then dNapr1 = dNapr1 + 2 * Math.PI
                                        Dim dNapr2 As Double = Math.Round(Vect_VU_KSP.Angle - Napr2, 5)
                                        If dNapr2 < 0 Then dNapr2 = dNapr2 + 2 * Math.PI
                                        If dNapr1 < dNapr2 Then
                                            Dugol = -1 * ugol
                                        Else
                                            Dugol = ugol
                                        End If
                                        If dNapr1 < dNapr2 Then
                                            Dugol = -1 * ugol
                                        Else
                                            Dugol = ugol
                                        End If
                                        ArrayExcel(2, NumbVU) = Dugol
                                        'записываем пикеты конструкций
                                        ArrayExcel(12, NumbVU) = NPK 'начало переходной кривой
                                        ArrayExcel(13, NumbVU) = NKK 'начало круговой кривой
                                        ArrayExcel(14, NumbVU) = SK 'конец круговой кривой
                                        ArrayExcel(15, NumbVU) = KPK 'конец переходной кривой
                                        '======================================================================================================
                                        'записываем радиус дуги
                                        ArrayExcel(3, NumbVU) = ArrayEntity(10, i)
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'координаты вершины угла 2
                                    If boolIntPoint2 = True Then
                                        Dim ldVu2 As Double = endKT.GetVectorTo(PTVUxy2).Length
                                        If ldVu2 < 0.02 Then Continue For
                                        'открываем 2 новые строки массива
                                        ReDim Preserve ArrayExcel(23, NumbVU + 1)
                                        PTVUxyz = New Point3d(PTVUxy2.X, PTVUxy2.Y, 0)
                                        ArrayExcel(19, NumbVU + 1) = PTVUxy2.X
                                        ArrayExcel(20, NumbVU + 1) = PTVUxy2.Y
                                        '===========================================================================================================
                                        'номер 2 вершины 
                                        ArrayExcel(0, NumbVU + 1) = NumbVU + 2
                                        '================================================================================================
                                        'вычисляем характеристики 2 половины дуги
                                        'определяем тангенсы второй половины дуги
                                        Vect_T1 = PtSk.GetVectorTo(PTVUxy2) 'середина дуги - вершина угла
                                        Vect_T2 = PTVUxy2.GetVectorTo(EndPoint) 'вершина угла - конец конструкции
                                        T1 = Vect_T1.Length '1 тангенс
                                        T2 = Vect_T2.Length '2 тангенс
                                        ArrayExcel(6, NumbVU + 1) = T1
                                        ArrayExcel(7, NumbVU + 1) = T2
                                        '=============================================================================================
                                        'длины переходных кривых
                                        ArrayExcel(4, NumbVU + 1) = 0 'вторая половина дуги
                                        ArrayExcel(5, NumbVU + 1) = L2
                                        '========================================================================================
                                        'определяесся с пикетом конца конструкции
                                        Dim EndPKKonstr As Double = KPK
                                        If L2 > 0 Then
                                            EndPKKonstr = KPK
                                        Else
                                            EndPKKonstr = KKK
                                        End If
                                        'записываем пикет вершины угла
                                        'находим пикет начала конструкции (с учетом рубленности)
                                        Dim SKStrPK As String = "-1"
                                        Try
                                            SKStrPK = Align.GetStationStringWithEquations(SK)
                                        Catch ex As System.ArgumentException
                                        End Try
                                        'переводим строку в число
                                        If Not (SKStrPK Like "-1") Then
                                            Dim NPKDblPK As Double = FuncAlignment.FuncFormatPKtoDbl(SKStrPK)
                                            'проверяем нахождение начала кривй на рубленном пикете
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, SK, EndPKKonstr)
                                            ArrayExcel(1, NumbVU + 1) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        Else
                                            ArrayExcel(1, NumbVU + 1) = 0 'в случае неудачи
                                        End If
                                        '================================================================================================
                                        'вычисляем длину кривой для первой половины дуги
                                        Dim Ksohr As Double = 0
                                        Ksohr = KKK - SK
                                        K = L2 + Ksohr
                                        ArrayExcel(8, NumbVU + 1) = K
                                        ArrayExcel(9, NumbVU + 1) = Ksohr
                                        '================================================================================================
                                        'определяем биссектрису
                                        Dim LineTX1 As Double = 10 * Math.Cos(Vect_T1.Angle + Math.PI) + PTVUxy2.X
                                        Dim LineTY1 As Double = 10 * Math.Sin(Vect_T1.Angle + Math.PI) + PTVUxy2.Y
                                        Dim LineTX2 As Double = 10 * Math.Cos(Vect_T2.Angle) + PTVUxy2.X
                                        Dim LineTY2 As Double = 10 * Math.Sin(Vect_T2.Angle) + PTVUxy2.Y
                                        Dim LineT1XSr As Double = (LineTX1 + LineTX2) / 2
                                        Dim LineT1YSr As Double = (LineTY1 + LineTY2) / 2
                                        Dim ptLineT1 As Point3d = New Point3d(LineT1XSr, LineT1YSr, 0)
                                        Dim LineB As Line = New Line(PTVUxyz, ptLineT1)
                                        Dim PointCol As Point3dCollection = New Point3dCollection
                                        Align.IntersectWith(LineB, Intersect.ExtendBoth, PointCol, 1, 1)
                                        If PointCol.Count > 0 Then
                                            For atp As Integer = 0 To PointCol.Count - 1
                                                Dim PtB As Point3d = PointCol.Item(atp)
                                                Dim PtB2d As Point2d = New Point2d(PtB.X, PtB.Y)
                                                Dim Vect_B As Vector2d = PtB2d.GetVectorTo(PTVUxy2)
                                                B = Vect_B.Length
                                                ArrayExcel(10, NumbVU + 1) = B
                                                Dim PKB As Double = 0
                                                Dim PKOff As Double = 0
                                                Try
                                                    Align.StationOffset(PtB.X, PtB.Y, PKB, PKOff)
                                                    If PKB > SK And PKB < EndPKKonstr Then
                                                        If Math.Abs(PKOff) < 0.01 Then
                                                            ArrayExcel(23, NumbVU + 1) = PKB
                                                        End If
                                                    End If
                                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                End Try
                                            Next
                                        End If
                                        '===============================================================================================
                                        'вычисляем домер
                                        D = T1 + T2 - K
                                        ArrayExcel(11, NumbVU + 1) = D
                                        '===================================================================================================
                                        '==================================================================================================================
                                        'вычисляем величину угла дельта для второй вершины
                                        'определяем дирекционный угол 1 тангенса
                                        Dim DUgol As Double = 0
                                        Dim Vect_SK_VU2 As Vector2d = PtSk.GetVectorTo(PTVUxy2) 'вектор ночало спирали - вершина угла
                                        Dim Vect_VU2_KSP As Vector2d = PTVUxy2.GetVectorTo(EndPoint) 'вектор вершина угла - конец спирали
                                        Dim Vect_SK_NSP As Vector2d = PtSk.GetVectorTo(EndPoint) 'вектор конец спирали - начало спирали
                                        'вычисляем внутренний угол
                                        Dim k1 As Double = Vect_SK_VU2.Length ^ 2 + Vect_VU2_KSP.Length ^ 2 - Vect_SK_NSP.Length ^ 2
                                        Dim k2 As Double = 2 * Vect_SK_VU2.Length * Vect_VU2_KSP.Length
                                        Dim ugol As Double = Math.PI - Math.Acos(k1 / k2) 'определили угол доворота
                                        'делаем проверку
                                        'прибавляем к направлению 1 тангенса величину угла доворота
                                        Dim Napr1 As Double = Vect_SK_VU2.Angle + ugol
                                        If Napr1 > Math.PI * 2 Then Napr1 = Napr1 - Math.PI * 2
                                        Dim Napr2 As Double = Vect_SK_VU2.Angle - ugol
                                        If Napr2 < 0 Then Napr2 = Napr2 + 2 * Math.PI
                                        'проверяем
                                        Dim dNapr1 As Double = Math.Round(Vect_VU2_KSP.Angle - Napr1, 5)
                                        If dNapr1 < 0 Then dNapr1 = dNapr1 + 2 * Math.PI
                                        Dim dNapr2 As Double = Math.Round(Vect_VU2_KSP.Angle - Napr2, 5)
                                        If dNapr2 < 0 Then dNapr2 = dNapr2 + 2 * Math.PI
                                        If dNapr1 < dNapr2 Then
                                            DUgol = -1 * ugol
                                        Else
                                            DUgol = ugol
                                        End If
                                        If dNapr1 < dNapr2 Then
                                            DUgol = -1 * ugol
                                        Else
                                            DUgol = ugol
                                        End If
                                        'форматируем угол
                                        ArrayExcel(2, NumbVU + 1) = DUgol
                                        '===================================================================================================
                                        ArrayExcel(12, NumbVU + 1) = NPK 'начало переходной кривой
                                        ArrayExcel(13, NumbVU + 1) = SK  'начало круговой кривой
                                        ArrayExcel(14, NumbVU + 1) = KKK 'конец круговой кривой
                                        ArrayExcel(15, NumbVU + 1) = KPK 'конец переходной кривой
                                        '======================================================================================================
                                        'записываем радиус дуги
                                        ArrayExcel(3, NumbVU + 1) = ArrayEntity(10, i)
                                    End If
                                    NumbVU += 2
                                Else
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'это простая дуга
                                    'открываем новую строку массива
                                    ReDim Preserve ArrayExcel(23, NumbVU)
                                    '==============================================================================
                                    'координаты вершины угла
                                    ArrayExcel(19, NumbVU) = PTVUxyz.X
                                    ArrayExcel(20, NumbVU) = PTVUxyz.Y
                                    '====================================================================================
                                    'номер вершины 
                                    ArrayExcel(0, NumbVU) = NumbVU + 1
                                    '================================================================================================
                                    'вычисляем характеристики 1 половины и второй половины дуги
                                    'определяем тангенсы дуги
                                    Vect_T1 = StartPoint.GetVectorTo(VUxy) 'начало конструкции - вершина угла
                                    Vect_T2 = VUxy.GetVectorTo(EndPoint) 'вершина угла - конец конструкции
                                    T1 = Vect_T1.Length '1 тангенс
                                    T2 = Vect_T2.Length '2 тангенс
                                    ArrayExcel(6, NumbVU) = T1
                                    ArrayExcel(7, NumbVU) = T2
                                    '=============================================================================================
                                    'длины переходных кривых
                                    ArrayExcel(4, NumbVU) = L1 'первая половина дуги
                                    ArrayExcel(5, NumbVU) = L2
                                    '========================================================================================
                                    'определяемся с пикетом начала и конца конструкции
                                    Dim StartPKKonstr As Double = NPK
                                    If L1 > 0 Then
                                        StartPKKonstr = NPK
                                    Else
                                        StartPKKonstr = NKK
                                    End If
                                    Dim EndPKKonstr As Double = KPK
                                    If L2 > 0 Then
                                        EndPKKonstr = KPK
                                    Else
                                        EndPKKonstr = KKK
                                    End If
                                    'записываем пикет вершины угла
                                    'находим пикет начала конструкции (с учетом рубленности)
                                    Dim NPKStrPK As String = "-1"
                                    Try
                                        NPKStrPK = Align.GetStationStringWithEquations(StartPKKonstr)
                                    Catch ex As System.ArgumentException
                                    End Try
                                    '================================================================================================
                                    'вычисляем длину кривой для первой половины дуги
                                    Dim Ksohr As Double = KKK - NKK 'длина дуги (пикет середины дуги минус пикет начала дуги
                                    K = L1 + L2 + Ksohr
                                    ArrayExcel(8, NumbVU) = K
                                    ArrayExcel(9, NumbVU) = Ksohr
                                    '================================================================================================
                                    'определяем биссектрису
                                    'рисуем вспомогательную линию
                                    Dim LineTX1 As Double = 10 * Math.Cos(Vect_T1.Angle + Math.PI) + VUxy.X
                                    Dim LineTY1 As Double = 10 * Math.Sin(Vect_T1.Angle + Math.PI) + VUxy.Y
                                    Dim LineTX2 As Double = 10 * Math.Cos(Vect_T2.Angle) + VUxy.X
                                    Dim LineTY2 As Double = 10 * Math.Sin(Vect_T2.Angle) + VUxy.Y
                                    Dim LineT1XSr As Double = (LineTX1 + LineTX2) / 2
                                    Dim LineT1YSr As Double = (LineTY1 + LineTY2) / 2
                                    Dim ptLineT1 As Point3d = New Point3d(LineT1XSr, LineT1YSr, 0)
                                    Dim LineB As Line = New Line(PTVUxyz, ptLineT1)
                                    Dim PointCol As Point3dCollection = New Point3dCollection
                                    Dim flagFindB As Boolean = False
                                    Align.IntersectWith(LineB, Intersect.ExtendBoth, PointCol, 1, 1)
                                    If PointCol.Count > 0 Then
                                        For atp As Integer = 0 To PointCol.Count - 1
                                            Dim PtB As Point3d = PointCol.Item(atp)
                                            Dim PtB2d As Point2d = New Point2d(PtB.X, PtB.Y)
                                            Dim Vect_B As Vector2d = PtB2d.GetVectorTo(VUxy)
                                            B = Vect_B.Length
                                            ArrayExcel(10, NumbVU) = B
                                            Dim PKB As Double = 0
                                            Dim PKOff As Double = 0
                                            Try
                                                Align.StationOffset(PtB.X, PtB.Y, PKB, PKOff)
                                                If PKB > StartPKKonstr And PKB < EndPKKonstr Then
                                                    If Math.Abs(PKOff) < 0.01 Then
                                                        ArrayExcel(23, NumbVU) = PKB
                                                        flagFindB = True
                                                    End If
                                                End If
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        Next
                                    End If
                                    '==================================================================================================
                                    'переводим строку в число
                                    If Not (NPKStrPK Like "-1") Then
                                        Dim NPKDblPK As Double = Val(NPKStrPK.Replace("+", ""))
                                        'проверяем нахождение начала кривй на рубленном пикете
                                        If flagFindB = True Then
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, StartPKKonstr, ArrayExcel(23, NumbVU))
                                            ArrayExcel(1, NumbVU) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        Else
                                            Dim SKTemp As Double = (StartPKKonstr + EndPKKonstr) / 2
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, StartPKKonstr, SKTemp)
                                            ArrayExcel(1, NumbVU) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        End If

                                    Else
                                        ArrayExcel(1, NumbVU) = 0 'в случае неудачи
                                    End If
                                    '===============================================================================================
                                    'вычисляем домер
                                    D = T1 + T2 - K
                                    ArrayExcel(11, NumbVU) = D
                                    '===================================================================================================
                                    'вычисляем величину угла дельта для первой вершины
                                    Dim Dugol As Double = 0
                                    'определяем дирекционный угол 1 тангенса
                                    Dim Vect_NSP_VU As Vector2d = StartPoint.GetVectorTo(VUxy) 'вектор ночало спирали - вершина угла
                                    Dim Vect_VU_KSP As Vector2d = VUxy.GetVectorTo(EndPoint) 'вектор вершина угла - конец спирали
                                    Dim Vect_KSP_NSP As Vector2d = StartPoint.GetVectorTo(EndPoint) 'вектор конец спирали - начало спирали
                                    'вычисляем внутренний угол
                                    Dim k1 As Double = Vect_NSP_VU.Length ^ 2 + Vect_VU_KSP.Length ^ 2 - Vect_KSP_NSP.Length ^ 2
                                    Dim k2 As Double = 2 * Vect_NSP_VU.Length * Vect_VU_KSP.Length
                                    Dim ugol As Double = Math.PI - Math.Acos(k1 / k2) 'определили угол доворота
                                    'делаем проверку
                                    'прибавляем к направлению 1 тангенса величину угла доворота
                                    Dim Napr1 As Double = Vect_NSP_VU.Angle + ugol
                                    If Napr1 > Math.PI * 2 Then Napr1 = Napr1 - Math.PI * 2
                                    Dim Napr2 As Double = Vect_NSP_VU.Angle - ugol
                                    If Napr2 < 0 Then Napr2 = Napr2 + 2 * Math.PI
                                    'проверяем
                                    Dim dNapr1 As Double = Math.Round(Vect_VU_KSP.Angle - Napr1, 5)
                                    If dNapr1 < 0 Then dNapr1 = dNapr1 + 2 * Math.PI
                                    Dim dNapr2 As Double = Math.Round(Vect_VU_KSP.Angle - Napr2, 5)
                                    If dNapr2 < 0 Then dNapr2 = dNapr2 + 2 * Math.PI
                                    If dNapr1 < dNapr2 Then
                                        Dugol = -1 * ugol
                                    Else
                                        Dugol = ugol
                                    End If
                                    If dNapr1 < dNapr2 Then
                                        Dugol = -1 * ugol
                                    Else
                                        Dugol = ugol
                                    End If
                                    'форматируем угол
                                    ArrayExcel(2, NumbVU) = Dugol
                                    '===================================================================================================
                                    'записываем пикеты конструкций
                                    ArrayExcel(12, NumbVU) = NPK
                                    ArrayExcel(13, NumbVU) = NKK
                                    ArrayExcel(14, NumbVU) = KKK
                                    ArrayExcel(15, NumbVU) = KPK
                                    '======================================================================================================
                                    'записываем радиус дуги
                                    ArrayExcel(3, NumbVU) = ArrayEntity(10, i)
                                    NumbVU += 1
                                End If

                            End If
                            '=========================================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'рассматриваем переходные кривые
                        ElseIf TypeObj = 2 Then 'это переходная кривая
                            NPK = ArrayEntity(0, i) 'берем пикет начала и пикет конца дуги
                            KPK = ArrayEntity(1, i)
                            'координаты начала и конца конструкции
                            StartPoint = New Point2d(ArrayEntity(5, i), ArrayEntity(6, i))
                            EndPoint = New Point2d(ArrayEntity(7, i), ArrayEntity(8, i))
                            L1 = KPK - NPK 'длина переходной кривой
                            'записываем координаты начала, конца и центра дуги в массив
                            Dim ArraySpi As Double() = Nothing
                            ReDim ArraySpi(8)
                            ArraySpi(0) = ArrayEntity(5, i) 'начало переходной кривой
                            ArraySpi(1) = ArrayEntity(6, i)
                            ArraySpi(2) = ArrayEntity(7, i) 'конец переходной кривой
                            ArraySpi(3) = ArrayEntity(8, i)
                            ArraySpi(4) = ArrayEntity(2, i) 'точка SPI
                            ArraySpi(5) = ArrayEntity(3, i)
                            ArraySpi(6) = ArrayEntity(9, i) 'пикет точки SPI
                            ArraySpi(7) = ArrayEntity(14, i) 'координаты радиальной точки
                            ArraySpi(8) = ArrayEntity(15, i)
                            'записываем предварительный радиус
                            Dim RinSp As Double = -1 'входящий радиус для спирали слева
                            Dim PKRinSp As Double = -1
                            Dim boolSPIn As Boolean = False 'флаг для элемента слева (элемент сонаправлен)
                            Dim boolSPOut As Boolean = False 'флаг для элемента справа (элемент сонаправлен)
                            'смотрим элемент слева
                            'условие - если это не первый элемент в массиве
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'заходим с лева (смотрим элемент слева от выбранной спирали)
                            If i > 0 Then
                                'берем элемент слева и смотрим
                                Dim TypeObj1 As Integer = ArrayEntity(4, i - 1) 'тип элемента
                                '==============================================================================
                                If TypeObj1 = 2 Then ''это спираль
                                    'это спираль проверяем ее на сонаправленность к дуге
                                    'записываем в массив ее характеристики
                                    Dim StSP As Double = ArrayEntity(0, i - 1) 'пикет начала переходной кривой
                                    Dim EndSP As Double = ArrayEntity(1, i - 1) 'пикет конца переходной кривой
                                    Dim StSp3d As Point3d = New Point3d(ArrayEntity(5, i - 1), ArrayEntity(5, i - 1), 0) 'координаты начала переходной кривой
                                    Dim SPI3d As Point3d = New Point3d(ArrayEntity(2, i - 1), ArrayEntity(3, i - 1), 0) 'координаты SPI точки
                                    Dim ArraySP As Double() = Nothing 'массив с координатами начала и конца переходной кривой
                                    ReDim ArraySP(8)
                                    ArraySP(0) = ArrayEntity(5, i - 1) 'начало переходной кривой
                                    ArraySP(1) = ArrayEntity(6, i - 1)
                                    ArraySP(2) = ArrayEntity(7, i - 1) 'конец переходной кривой
                                    ArraySP(3) = ArrayEntity(8, i - 1)
                                    ArraySP(4) = ArrayEntity(2, i - 1) 'координаты SpI точки
                                    ArraySP(5) = ArrayEntity(3, i - 1)
                                    ArraySP(6) = ArrayEntity(9, i - 1) 'пикет точки SPI
                                    ArraySpi(7) = ArrayEntity(14, i - 1) 'координаты радиальной точки
                                    ArraySpi(8) = ArrayEntity(15, i - 1)
                                    'если массив состоялся тогда
                                    If IsArray(ArraySP) = True Then
                                        'проверяем на сонаправленность
                                        boolSPIn = FuncAlignment.FuncSpiralSpiralNapr1(Align, ArraySP, ArraySpi)
                                        If boolSPIn = True Then
                                            'берем точки начала спирали и ии SPI (спирали слева)
                                            Dim StLine1 As Point3d = New Point3d(ArraySP(0), ArraySP(1), 0)
                                            Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i - 1), ArrayEntity(3, i - 1), 0)
                                            'строим вспомогательную линию
                                            Line1 = New Line(StLine1, StLine2)
                                            L2 = L1 'выбранный элемент становится вторым, спираль слева первой
                                            L1 = EndSP - StSP 'длина переходной кривой которая слева
                                            'новая стартовая точка (спирали слева)
                                            StartPoint = New Point2d(ArraySP(0), ArraySP(1))
                                            SK = EndSP 'середина конструкции
                                            NPK = StSP 'новое начало конструкции из переходных кривых
                                            NKK = EndSP 'новое начало конструкции из переходных кривых
                                            KKK = ArrayEntity(0, i)
                                            KPK = ArrayEntity(1, i)
                                            RinSp = ArrayEntity(10, i - 1) 'входящий радиус для спирали слева
                                            PKRinSp = ArrayEntity(0, i - 1)
                                        Else
                                            'спираль самостоятельная (самостоятельная слева)
                                            Dim StLine1 As Point3d = New Point3d(ArraySpi(0), ArraySpi(1), 0) 'начало переходной кривой
                                            Dim StLine2 As Point3d = New Point3d(ArraySpi(4), ArraySpi(5), 0) 'конец переходной кривой
                                            Line1 = New Line(StLine1, StLine2)
                                        End If
                                    End If
                                    '================================================================================================
                                    'если элемент слева это дуга
                                ElseIf TypeObj1 = 1 Then ''это дуга
                                    'заносим массив с характеристиками дуги
                                    Dim ArrayArc As Double() = Nothing
                                    ReDim ArrayArc(8)
                                    ArrayArc(0) = ArrayEntity(5, i - 1) 'начало дуги
                                    ArrayArc(1) = ArrayEntity(6, i - 1)
                                    ArrayArc(2) = ArrayEntity(7, i - 1) 'конец дуги
                                    ArrayArc(3) = ArrayEntity(8, i - 1)
                                    ArrayArc(4) = ArrayEntity(2, i - 1) 'центр дуги
                                    ArrayArc(5) = ArrayEntity(3, i - 1)
                                    ArrayArc(6) = (ArrayEntity(0, i - 1) + ArrayEntity(1, i - 1)) / 2 'пикет середины дуги
                                    ArrayArc(7) = ArrayEntity(2, i - 1) 'координаты центра дуги
                                    ArrayArc(8) = ArrayEntity(3, i - 1)
                                    If IsArray(ArrayArc) = True Then
                                        'проверяем на сонаправленность
                                        boolSPIn = FuncAlignment.FuncArcSpiralNapr1(Align, ArrayArc, ArraySpi)
                                        'если сонаправлены то это значит данная спираль уже была обработана ранее переходим к следующему элементу
                                        If boolSPIn = True Then
                                            GoTo Line3d
                                        Else
                                            'спираль самостоятельная
                                            Dim StLine1 As Point3d = New Point3d(ArraySpi(0), ArraySpi(1), 0) 'начало переходной кривой
                                            Dim StLine2 As Point3d = New Point3d(ArraySpi(4), ArraySpi(5), 0) 'конец переходной кривой
                                            Line1 = New Line(StLine1, StLine2)
                                        End If
                                    End If
                                Else
                                    'спираль самостоятельная
                                    Dim StLine1 As Point3d = New Point3d(ArraySpi(0), ArraySpi(1), 0) 'начало переходной кривой
                                    Dim StLine2 As Point3d = New Point3d(ArraySpi(4), ArraySpi(5), 0) 'конец переходной кривой
                                    Line1 = New Line(StLine1, StLine2)
                                End If
                            Else
                                'спираль самостоятельная
                                Dim StLine1 As Point3d = New Point3d(ArraySpi(0), ArraySpi(1), 0) 'начало переходной кривой
                                Dim StLine2 As Point3d = New Point3d(ArraySpi(4), ArraySpi(5), 0) 'конец переходной кривой
                                Line1 = New Line(StLine1, StLine2)
                            End If
                            '================================================================================
                            'заходим с обратной стороны если есть спираль справа и она сонаправлена то вообще ни чего не делаем иначе это обособленная спираль
                            If i < ArrayEntity.GetUpperBound(1) Then
                                Dim TypeObj1 As Integer = ArrayEntity(4, i + 1)
                                If TypeObj1 = 2 Then ''это спираль
                                    'это спираль проверяем ее на сонаправленность к дуге
                                    Dim StSP As Double = ArrayEntity(0, i + 1) 'пикет начала переходной кривой
                                    Dim EndSP As Double = ArrayEntity(1, i + 1) 'пикет конца переходной кривой
                                    L2 = EndSP - StSP
                                    Dim StSp3d As Point3d = New Point3d(ArrayEntity(5, i + 1), ArrayEntity(5, i + 1), 0) 'координаты начала переходной кривой
                                    Dim SPI3d As Point3d = New Point3d(ArrayEntity(2, i + 1), ArrayEntity(3, i + 1), 0) 'координаты SPI точки
                                    Dim ArraySP As Double() = Nothing 'массив с координатами начала и конца переходной кривой
                                    ReDim ArraySP(8)
                                    ArraySP(0) = ArrayEntity(5, i + 1) 'начало переходной кривой
                                    ArraySP(1) = ArrayEntity(6, i + 1)
                                    ArraySP(2) = ArrayEntity(7, i + 1) 'конец переходной кривой
                                    ArraySP(3) = ArrayEntity(8, i + 1)
                                    ArraySP(4) = ArrayEntity(2, i + 1) 'координаты SpI точки
                                    ArraySP(5) = ArrayEntity(3, i + 1)
                                    ArraySP(6) = ArrayEntity(9, i + 1) 'пикет точки SPI
                                    ArraySP(7) = ArrayEntity(14, i + 1) 'координаты радиальной точки
                                    ArraySP(8) = ArrayEntity(15, i + 1)
                                    'если массив состоялся тогда
                                    If IsArray(ArraySP) = True Then
                                        'проверяем на сонаправленность
                                        boolSPIn = FuncAlignment.FuncSpiralSpiralNapr1(Align, ArraySpi, ArraySP)
                                        If boolSPIn = True Then
                                            'спирали сонаправлены, выходим из ветки этого цикла, обработаем ее со следующего элемента
                                            GoTo Line3d
                                        Else
                                            'дуга справа соседа не имеет ищем направление для тангенса
                                            'ищем вектор центр дуги - начальный пикет
                                            Dim StLine1 As Point3d = New Point3d(ArraySpi(2), ArraySpi(3), 0)
                                            Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i), ArrayEntity(3, i), 0)
                                            Line2 = New Line(StLine1, StLine2)
                                        End If
                                    End If
                                    'если это дуга
                                ElseIf TypeObj1 = 1 Then
                                    Dim ArrayArc As Double() = Nothing
                                    ReDim ArrayArc(8)
                                    ArrayArc(0) = ArrayEntity(5, i + 1) 'начало дуги
                                    ArrayArc(1) = ArrayEntity(6, i + 1)
                                    ArrayArc(2) = ArrayEntity(7, i + 1) 'конец дуги
                                    ArrayArc(3) = ArrayEntity(8, i + 1)
                                    ArrayArc(4) = ArrayEntity(2, i + 1) 'центр дуги
                                    ArrayArc(5) = ArrayEntity(3, i + 1)
                                    ArrayArc(6) = (ArrayEntity(1, i + 1) + ArrayEntity(2, i + 1)) / 2 'пикет середины дуги
                                    ArrayArc(7) = ArrayEntity(2, i + 1) 'начало дуги
                                    ArrayArc(8) = ArrayEntity(3, i + 1)
                                    If IsArray(ArrayArc) = True Then
                                        'проверяем на сонаправленность
                                        boolSPIn = FuncAlignment.FuncSpiralArcNapr1(Align, ArraySpi, ArrayArc)
                                        If boolSPIn = True Then
                                            'спирали сонаправлены, выходим из ветки этого цикла, обработаем ее со следующего элемента
                                            GoTo Line3d
                                        Else
                                            'дуга слева соседа не имеет ищем направление для тангенса
                                            'ищем вектор центр дуги - начальный пикет
                                            Dim StLine1 As Point3d = New Point3d(ArraySpi(2), ArraySpi(3), 0)
                                            Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i), ArrayEntity(3, i), 0)
                                            Line2 = New Line(StLine1, StLine2)
                                        End If
                                    End If
                                Else
                                    'дуга первый элемент
                                    'дуга слева соседа не имеет ищем направление для тангенса
                                    'ищем вектор центр дуги - начальный пикет
                                    Dim StLine1 As Point3d = New Point3d(ArraySpi(2), ArraySpi(3), 0)
                                    Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i), ArrayEntity(3, i), 0)
                                    Line2 = New Line(StLine1, StLine2)
                                End If
                            Else
                                Dim StLine1 As Point3d = New Point3d(ArraySpi(2), ArraySpi(3), 0)
                                Dim StLine2 As Point3d = New Point3d(ArrayEntity(2, i), ArrayEntity(3, i), 0)
                                Line2 = New Line(StLine1, StLine2)
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'вычисление элементов
                            'вычисляем вершину угла
                            If IsNothing(Line1) = False And IsNothing(Line2) = False Then
                                Dim PTVUxyz As Point3d = Nothing 'координата вершины дуги 3д
                                Dim intersectPoint As Point3dCollection = New Point3dCollection
                                Try
                                    Line1.IntersectWith(Line2, Intersect.ExtendBoth, intersectPoint, 1, 1)
                                    If intersectPoint.Count > 0 Then
                                        PTVUxyz = intersectPoint.Item(0)
                                        VUxy = New Point2d(PTVUxyz.X, PTVUxyz.Y)
                                    End If
                                Catch ex As Exception
                                End Try
                                'точка пересечения тангенсов найдена продолжаем процедуру
                                If intersectPoint.Count = 0 Then
                                    'пытаемся математически найти точку пересечения
                                    Dim pt1 As Point2d = New Point2d(Line1.StartPoint.X, Line1.StartPoint.Y)
                                    Dim pt2 As Point2d = New Point2d(Line1.EndPoint.X, Line1.EndPoint.Y)
                                    Dim pt3 As Point2d = New Point2d(Line2.StartPoint.X, Line2.StartPoint.Y)
                                    Dim pt4 As Point2d = New Point2d(Line2.EndPoint.X, Line2.EndPoint.Y)
                                    Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, VUxy)
                                End If
                                If intersectPoint.Count > 0 Then
                                    'ищем точку пересечения двух линий
                                    ReDim Preserve ArrayExcel(23, NumbVU)
                                    '==============================================================================
                                    'координаты вершины угла 1
                                    ArrayExcel(19, NumbVU) = PTVUxyz.X
                                    ArrayExcel(20, NumbVU) = PTVUxyz.Y
                                    '====================================================================================
                                    'номер вершины 
                                    ArrayExcel(0, NumbVU) = NumbVU + 1
                                    '================================================================================================
                                    'вычисляем характеристики 1 половины и второй половины дуги
                                    'определяем тангенсы 1 половины дуги
                                    Dim Vect_T1 As Vector2d = StartPoint.GetVectorTo(VUxy) 'начало конструкции - вершина угла
                                    Dim Vect_T2 As Vector2d = VUxy.GetVectorTo(EndPoint) 'вершина угла - конец конструкции
                                    T1 = Vect_T1.Length '1 тангенс
                                    T2 = Vect_T2.Length '2 тангенс
                                    ArrayExcel(6, NumbVU) = T1
                                    ArrayExcel(7, NumbVU) = T2
                                    '=============================================================================================
                                    'длины переходных кривых
                                    ArrayExcel(4, NumbVU) = L1 'первая половина дуги
                                    ArrayExcel(5, NumbVU) = L2
                                    '========================================================================================
                                    'записываем пикет вершины угла
                                    'находим пикет начала конструкции (с учетом рубленности)
                                    Dim NPKStrPK As String = "-1"
                                    Try
                                        NPKStrPK = Align.GetStationStringWithEquations(NPK)
                                    Catch ex As System.ArgumentException
                                    End Try
                                    '================================================================================================
                                    'вычисляем длину кривой для первой половины дуги
                                    Dim Kpoln1 As Double = KPK - NPK 'длина дуги (пикет середины дуги минус пикет начала дуги
                                    K = L1 + L2
                                    ArrayExcel(8, NumbVU) = K
                                    '================================================================================================
                                    'определяем биссектрису
                                    'рисуем вспомогательную линию
                                    Dim LineTX1 As Double = 10 * Math.Cos(Vect_T1.Angle + Math.PI) + VUxy.X
                                    Dim LineTY1 As Double = 10 * Math.Sin(Vect_T1.Angle + Math.PI) + VUxy.Y
                                    Dim LineTX2 As Double = 10 * Math.Cos(Vect_T2.Angle) + VUxy.X
                                    Dim LineTY2 As Double = 10 * Math.Sin(Vect_T2.Angle) + VUxy.Y
                                    Dim LineT1XSr As Double = (LineTX1 + LineTX2) / 2
                                    Dim LineT1YSr As Double = (LineTY1 + LineTY2) / 2
                                    Dim ptLineT1 As Point3d = New Point3d(LineT1XSr, LineT1YSr, 0)
                                    Dim LineB As Line = New Line(PTVUxyz, ptLineT1)
                                    Dim PointCol As Point3dCollection = New Point3dCollection
                                    Dim boolFindB As Boolean = False
                                    Align.IntersectWith(LineB, Intersect.ExtendBoth, PointCol, 1, 1)
                                    If PointCol.Count > 0 Then
                                        For atp As Integer = 0 To PointCol.Count - 1
                                            Dim PtB As Point3d = PointCol.Item(atp)
                                            Dim PtB2d As Point2d = New Point2d(PtB.X, PtB.Y)
                                            Dim Vect_B As Vector2d = PtB2d.GetVectorTo(VUxy)
                                            B = Vect_B.Length
                                            ArrayExcel(10, NumbVU) = B
                                            Dim PKB As Double = 0
                                            Dim PKOff As Double = 0
                                            Try
                                                Align.StationOffset(PtB.X, PtB.Y, PKB, PKOff)
                                                If PKB > NPK And PKB < KPK Then
                                                    If Math.Abs(PKOff) < 0.01 Then
                                                        ArrayExcel(23, NumbVU) = PKB
                                                        boolFindB = True
                                                        Exit For
                                                    End If
                                                End If
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                        Next
                                    End If
                                    'переводим строку в число
                                    If Not (NPKStrPK Like "-1") Then
                                        Dim NPKDblPK As Double = FuncAlignment.FuncFormatPKtoDbl(NPKStrPK)
                                        'проверяем нахождение начала кривй на рубленном пикете
                                        If boolFindB = True Then
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, NPK, ArrayExcel(23, NumbVU))
                                            ArrayExcel(1, NumbVU) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        Else
                                            Dim PKBTemp As Double = (NPK + KPK) / 2
                                            Dim DeltaPK As Double = FuncAlignment.FuncFindRubkPK(Align, NPK, PKBTemp)
                                            ArrayExcel(1, NumbVU) = NPKDblPK + T1 + DeltaPK 'начало переходной кривой (начало круговой кривой) + тангенс1
                                        End If
                                    Else
                                        ArrayExcel(1, NumbVU) = 0 'в случае неудачи
                                    End If
                                    '===============================================================================================
                                    'вычисляем домер
                                    D = T1 + T2 - Kpoln1
                                    ArrayExcel(11, NumbVU) = D
                                    '===================================================================================================
                                    'вычисляем величину угла дельта для первой вершины
                                    Dim Dugol As Double = 0
                                    'определяем дирекционный угол 1 тангенса
                                    Dim Vect_NSP_VU As Vector2d = StartPoint.GetVectorTo(VUxy) 'вектор ночало спирали - вершина угла
                                    Dim Vect_VU_KSP As Vector2d = VUxy.GetVectorTo(EndPoint) 'вектор вершина угла - конец спирали
                                    Dim Vect_KSP_NSP As Vector2d = StartPoint.GetVectorTo(EndPoint) 'вектор конец спирали - начало спирали
                                    'вычисляем внутренний угол
                                    Dim k1 As Double = Vect_NSP_VU.Length ^ 2 + Vect_VU_KSP.Length ^ 2 - Vect_KSP_NSP.Length ^ 2
                                    Dim k2 As Double = 2 * Vect_NSP_VU.Length * Vect_VU_KSP.Length
                                    Dim ugol As Double = Math.PI - Math.Acos(k1 / k2) 'определили угол доворота
                                    'делаем проверку
                                    'прибавляем к направлению 1 тангенса величину угла доворота
                                    Dim Napr1 As Double = Vect_NSP_VU.Angle + ugol
                                    If Napr1 > Math.PI * 2 Then Napr1 = Napr1 - Math.PI * 2
                                    Dim Napr2 As Double = Vect_NSP_VU.Angle - ugol
                                    If Napr2 < 0 Then Napr2 = Napr2 + 2 * Math.PI
                                    'проверяем
                                    Dim dNapr1 As Double = Math.Round(Vect_VU_KSP.Angle - Napr1, 5)
                                    If dNapr1 < 0 Then dNapr1 = dNapr1 + 2 * Math.PI
                                    Dim dNapr2 As Double = Math.Round(Vect_VU_KSP.Angle - Napr2, 5)
                                    If dNapr2 < 0 Then dNapr2 = dNapr2 + 2 * Math.PI
                                    If dNapr1 < dNapr2 Then
                                        Dugol = -1 * ugol
                                    Else
                                        Dugol = ugol
                                    End If
                                    If dNapr1 < dNapr2 Then
                                        Dugol = -1 * ugol
                                    Else
                                        Dugol = ugol
                                    End If
                                    'форматируем угол
                                    ArrayExcel(2, NumbVU) = Dugol
                                    '===================================================================================================
                                    'записываем пикеты конструкций
                                    ArrayExcel(12, NumbVU) = NPK
                                    ArrayExcel(13, NumbVU) = NKK
                                    ArrayExcel(14, NumbVU) = KKK
                                    ArrayExcel(15, NumbVU) = KPK
                                    '======================================================================================================
                                    'записываем радиус переходной кривой
                                    ArrayExcel(3, NumbVU) = RinSp
                                    ArrayExcel(21, NumbVU) = ArrayEntity(10, i)
                                    ArrayExcel(22, NumbVU) = ArrayEntity(11, i)
                                    NumbVU += 1
                                End If
                            End If
                            ' ===================================================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'прямые участки
                        ElseIf TypeObj = 0 Then 'это прямая вставка
                            'обрабатываем исключительный случай если трасса состоит из одного элемента (одной единственной прямой)
                            'ищем соседа слева
                            If i > 0 Then
                                '======================================================================
                                'ищем соседа слева
                                Dim TypeObj1 As Integer = ArrayEntity(4, i - 1)
                                'если предыдущий объект не линия то записываем ее длину в массив (длину прямой вставки)
                                If TypeObj1 <> 0 Then
                                    ArrayExcel(17, NumbVU - 1) = ArrayEntity(12, i)
                                Else
                                    'берем характеристики линии слева
                                    Dim LinDistPrev As Double = ArrayEntity(12, i - 1) 'длина линии
                                    Dim LinNaprPrev As Double = ArrayEntity(13, i - 1) 'направление линии
                                    'берем характеристики линии исходной
                                    Dim LinDist As Double = ArrayEntity(12, i) 'длина линии
                                    Dim LinNapr As Double = ArrayEntity(13, i) 'направление линии
                                    'если направления совпадают тогда линии объединяем
                                    If LinNaprPrev = LinNapr Then
                                        'если в массиве уже есть углы поворота, то просто ссумируем длины
                                        If ArrayExcel.GetUpperBound(1) > 0 Then
                                            ArrayExcel(17, NumbVU - 1) = ArrayExcel(17, NumbVU - 1) + LinDist
                                        Else
                                            If i < i1 Then 'это начало конструкции (ссумируем начальную длину)
                                                TLenghtStart = TLenghtStart + LinDist
                                            End If
                                            If i > i2 Then
                                                TLenghtEnd = TLenghtEnd + LinDist
                                            End If
                                        End If
                                    Else
                                        'делаем новый угол поворота
                                        Dim StartPointLinPrev As Point2d = New Point2d(ArrayEntity(5, i - 1), ArrayEntity(6, i - 1)) 'начальная точка предыдущей линии
                                        Dim StartPointLin As Point2d = New Point2d(ArrayEntity(5, i), ArrayEntity(6, i)) 'начальная точка линии
                                        Dim EndPointLin As Point2d = New Point2d(ArrayEntity(7, i), ArrayEntity(8, i)) 'конечная точка линии
                                        ReDim Preserve ArrayExcel(23, NumbVU)
                                        '==============================================================================
                                        'координаты вершины угла 1
                                        ArrayExcel(19, NumbVU) = ArrayEntity(5, i)
                                        ArrayExcel(20, NumbVU) = ArrayEntity(6, i)
                                        '====================================================================================
                                        'номер вершины 
                                        ArrayExcel(0, NumbVU) = NumbVU + 1
                                        'пикет вершины
                                        ArrayExcel(1, NumbVU) = ArrayEntity(0, i)
                                        'записываем угол поворота
                                        'вычисляем величину угла дельта для первой вершины
                                        Dim Dugol As Double = 0
                                        'определяем дирекционный угол 1 тангенса
                                        Dim Vect_NSP_VU As Vector2d = StartPointLinPrev.GetVectorTo(StartPointLin) 'вектор ночало спирали - вершина угла
                                        Dim Vect_VU_KSP As Vector2d = StartPointLin.GetVectorTo(EndPointLin) 'вектор вершина угла - конец спирали
                                        Dim Vect_KSP_NSP As Vector2d = StartPointLinPrev.GetVectorTo(EndPointLin) 'вектор конец спирали - начало спирали
                                        'вычисляем внутренний угол
                                        Dim k1 As Double = Vect_NSP_VU.Length ^ 2 + Vect_VU_KSP.Length ^ 2 - Vect_KSP_NSP.Length ^ 2
                                        Dim k2 As Double = 2 * Vect_NSP_VU.Length * Vect_VU_KSP.Length
                                        Dim ugol As Double = Math.PI - Math.Acos(k1 / k2) 'определили угол доворота
                                        'делаем проверку
                                        'прибавляем к направлению 1 тангенса величину угла доворота
                                        Dim Napr1 As Double = Vect_NSP_VU.Angle + ugol
                                        If Napr1 > Math.PI * 2 Then Napr1 = Napr1 - Math.PI * 2
                                        Dim Napr2 As Double = Vect_NSP_VU.Angle - ugol
                                        If Napr2 < 0 Then Napr2 = Napr2 + 2 * Math.PI
                                        'проверяем
                                        Dim dNapr1 As Double = Math.Round(Vect_VU_KSP.Angle - Napr1, 5)
                                        If dNapr1 < 0 Then dNapr1 = dNapr1 + 2 * Math.PI
                                        Dim dNapr2 As Double = Math.Round(Vect_VU_KSP.Angle - Napr2, 5)
                                        If dNapr2 < 0 Then dNapr2 = dNapr2 + 2 * Math.PI
                                        If dNapr1 < dNapr2 Then
                                            Dugol = -1 * ugol
                                        Else
                                            Dugol = ugol
                                        End If
                                        If dNapr1 < dNapr2 Then
                                            Dugol = -1 * ugol
                                        Else
                                            Dugol = ugol
                                        End If
                                        ArrayExcel(2, NumbVU) = Dugol
                                        'записываем длину прямой вставки
                                        If NumbVU > 0 Then
                                            ArrayExcel(17, NumbVU - 1) = LinDistPrev
                                        End If
                                        '============================================================================================
                                        'рисуем дельта угол
                                        'флажек - рисовать графику
                                        If FormRoadPlane.CheckBox1.Checked = True Then
                                            '3. Вставляем дельта-угол
                                            Dim AngleT1 As Double = Vect_NSP_VU.Angle
                                            Dim AngleT2 As Double = Vect_VU_KSP.Angle
                                            Dim VnT1X As Double = 40 * Math.Cos(AngleT1) + StartPointLin.X
                                            Dim VnT1Y As Double = 40 * Math.Sin(AngleT1) + StartPointLin.Y
                                            Dim VnT2X As Double = (40) * Math.Cos(AngleT2) + StartPointLin.X 'вторая точка вставки для угла
                                            Dim VnT2Y As Double = (40) * Math.Sin(AngleT2) + StartPointLin.Y
                                            'точка вставки текста
                                            Dim ArcPtX As Double = (VnT1X + VnT2X) / 2
                                            Dim ArcPtY As Double = (VnT1Y + VnT2Y) / 2
                                            'вставляем угол
                                            Dim acLinAngDim As Point3AngularDimension = New Point3AngularDimension
                                            acLinAngDim.SetDatabaseDefaults()
                                            acLinAngDim.XLine1Point = New Point3d(VnT1X, VnT1Y, 0)
                                            acLinAngDim.XLine2Point = New Point3d(VnT2X, VnT2Y, 0)
                                            acLinAngDim.CenterPoint = New Point3d(StartPointLin.X, StartPointLin.Y, 0)
                                            acLinAngDim.ArcPoint = New Point3d(ArcPtX, ArcPtY, 0)
                                            acLinAngDim.DimensionStyle = acCurDb.Dimstyle
                                            acLinAngDim.Dimaunit = 1
                                            acLinAngDim.Dimdec = 5
                                            acLinAngDim.Prefix = "\U+0394="
                                            acLinAngDim.Layer = LayerDim
                                            acBlkTblRec.AppendEntity(acLinAngDim)
                                            tr.AddNewlyCreatedDBObject(acLinAngDim, True)
                                            'вставляем линиюпродолжения тангенса
                                            Dim ArrayVertexContTan As Double(,) = Nothing
                                            ReDim Preserve ArrayVertexContTan(1, 1)
                                            ArrayVertexContTan(0, 0) = StartPointLin.X
                                            ArrayVertexContTan(1, 0) = StartPointLin.Y
                                            ArrayVertexContTan(0, 1) = VnT1X
                                            ArrayVertexContTan(1, 1) = VnT1Y
                                            Dim acPolyT1 As Autodesk.AutoCAD.DatabaseServices.Polyline = AcadFunction.FuncDrawPolylineByArrayProperties(ArrayVertexContTan, ArrayPropLinPTH)
                                        End If
                                        NumbVU += 1
                                    End If
                                End If
                            ElseIf i = 0 Then
                                TLenghtStart = ArrayEntity(12, i)
                            End If
                            If i = ArrayEntity.GetUpperBound(1) Then
                                TLenghtEnd = TLenghtEnd + ArrayEntity(12, i)
                            End If
                            '============================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'вставляем графику
                            'значение градусов
                            'записываем румб
                            If FormRoadPlane.CheckBox1.Checked = True Then
                                Dim DirectMath As Double = ArrayEntity(13, i) 'угол в радианах (румб)
                                'Dim DirectGeodesic As Double = FuncConvertMathAngleToDirectionAngle(DirectMath)
                                Dim Grad As Double = (DirectMath * 180) / Math.PI 'переводим в градусы
                                Dim RumbLine As String = FuncSurvey.FuncDirectToRumbDDMMSS(Grad, FormRoadPlane.NumericUpDown1.Value)
                                '==============================================================================================
                                'длина линии
                                Dim S As Double = ArrayEntity(12, i)
                                'форматируем значение
                                Dim SLine As String = FuncFormatZN.FuncFormatCoordinate(S, FormRoadPlane.NumericUpDown6.Value) 'длина линии
                                '================================================================================================
                                'форматируем значение
                                Dim PtLineStartPK As Double = ArrayEntity(0, i)
                                Dim PtLineEndPK As Double = ArrayEntity(1, i)
                                Dim PKInsTxt As Double = (PtLineStartPK + PtLineEndPK) / 2
                                'ишем точку вставки текстов
                                Dim andLine As Double = ArrayEntity(13, i) 'направление объекта
                                Dim XTextA As Double = -1
                                Dim YTextA As Double = -1
                                Dim XTextD As Double = -1
                                Dim YTextD As Double = -1
                                Dim AngleText As Double = 0
                                Dim AttLine As Integer = 22
                                Dim AttAng As Integer = 22
                                Dim AnglM As Double = FuncSurvey.FuncDirectionAngleToMathAngle(andLine)
                                Select Case AnglM
                                    Case 0 To Math.PI / 2
                                        Try
                                            Align.PointLocation(PKInsTxt, -offsetTextLin, XTextA, YTextA)
                                            Align.PointLocation(PKInsTxt, offsetTextLin, XTextD, YTextD)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            XTextA = (ArrayEntity(5, i) + ArrayEntity(7, i)) / 2
                                            YTextA = (ArrayEntity(6, i) + ArrayEntity(8, i)) / 2
                                            XTextD = XTextA
                                            YTextD = YTextA
                                        End Try
                                        AngleText = AnglM
                                        AttLine = 24
                                        AttAng = 22
                                    Case Math.PI / 2 To Math.PI
                                        Try
                                            Align.PointLocation(PKInsTxt, offsetTextLin, XTextA, YTextA)
                                            Align.PointLocation(PKInsTxt, -offsetTextLin, XTextD, YTextD)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            XTextA = (ArrayEntity(5, i) + ArrayEntity(7, i)) / 2
                                            YTextA = (ArrayEntity(6, i) + ArrayEntity(8, i)) / 2
                                            XTextD = XTextA
                                            YTextD = YTextA
                                        End Try
                                        AngleText = AnglM + Math.PI
                                        AttLine = 24
                                        AttAng = 22
                                    Case Math.PI To (3 * Math.PI) / 2
                                        Try
                                            Align.PointLocation(PKInsTxt, offsetTextLin, XTextA, YTextA)
                                            Align.PointLocation(PKInsTxt, -offsetTextLin, XTextD, YTextD)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            XTextA = (ArrayEntity(5, i) + ArrayEntity(7, i)) / 2
                                            YTextA = (ArrayEntity(6, i) + ArrayEntity(8, i)) / 2
                                            XTextD = XTextA
                                            YTextD = YTextA
                                        End Try
                                        AngleText = AnglM - Math.PI
                                        AttLine = 24
                                        AttAng = 22
                                    Case (3 * Math.PI) / 2 To Math.PI * 2
                                        Try
                                            Align.PointLocation(PKInsTxt, -offsetTextLin, XTextA, YTextA)
                                            Align.PointLocation(PKInsTxt, offsetTextLin, XTextD, YTextD)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            XTextA = (ArrayEntity(5, i) + ArrayEntity(7, i)) / 2
                                            YTextA = (ArrayEntity(6, i) + ArrayEntity(8, i)) / 2
                                            XTextD = XTextA
                                            YTextD = YTextA
                                        End Try
                                        AngleText = AnglM
                                        AttLine = 24
                                        AttAng = 22
                                End Select
                                Dim insTextL1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(XTextA, YTextA, 0), "r=" & RumbLine, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleText, AttAng)
                                Dim insTextL2 As Boolean = AcadFunction.FuncDrawMText(New Point3d(XTextD, YTextD, 0), "S=" & SLine, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleText, AttLine)
                            End If
                        End If
Line3d:
                    Next i


                    '===================================================================================================================
                    'записываем X и Y
                    'Dim Xstart As Double = Math.Round(AlignStartY, 3)
                    'Dim Ystart As Double = Math.Round(AlignStartX, 3)
                    'Dim Xend As Double = Math.Round(AlignEndY, 3)
                    ' Dim Yend As Double = Math.Round(AlignEndX, 3)
                    'Dim startNT As Point2d = New Point2d(Xstart, Ystart)
                    'Dim endKT As Point2d = New Point2d(Xend, Yend)



                    '==========================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выводим отчет
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    If FormRoadPlane.RadioButton1.Checked = True And FormRoadPlane.CheckBox2.Checked = True Then
                        'стартовый пикет
                        Dim PiketStartStr As String = PKStartAlign
                        Try
                            PiketStartStr = Align.GetStationStringWithEquations(PKStartAlign) 'преобразовываем RawPiket
                        Catch ex As System.ArgumentException
                        End Try
                        'конечный пикет
                        Dim PiketEndStr As String = PKEndAlign
                        Try
                            PiketEndStr = Align.GetStationStringWithEquations(PKEndAlign) 'преобразовываем RawPiket
                        Catch ex As System.ArgumentException
                        End Try
                        '=================================================================================================================
                        'выводим отчет в MS Excel
                        If FormRoadPlane.ComboBox10.SelectedIndex = 0 Then
                            Dim IndTableExcel As Integer = FormRoadPlane.ComboBox4.SelectedIndex
                            'находим все директории с файлами
                            Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                            Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                            'отделяем имя файла от папки
                            Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                            'определяем имя родительской папки для папки где находится сборка
                            Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                            Dim PutchTables As String = parentPatch2 & "\TemplateTables\Roads\Ведомости углов поворота\"  'пути для шаблона к образцам таблиц
                            If IO.Directory.Exists(PutchTables) = True Then
                                'запускаем Excel
                                Dim xlApp As Excel.Application = Nothing
                                Dim xlBook As Excel.Workbook = Nothing
                                Dim xlSheets As Excel.Worksheet = Nothing
                                Dim cell1 As Excel.Range 'первая ячейка в строке
                                Dim cell2 As Excel.Range 'последняя ячейка в строке
                                Dim CellRng As Excel.Range 'массив ячеек
                                xlApp = CreateObject("Excel.Application")
                                xlApp.Visible = True
                                'форма отчета Расширенный
                                If IndTableExcel = 0 Then
                                    Dim ColTemplateExcel As Integer = 23
                                    Dim PutchDoc As String = PutchTables & "Форма 4_ГОСТ 21.1701-97_Расширенная.xlsx"
                                    Dim rng As Random = New Random
                                    Dim DestDoc As String = TempDir & "Форма 4_ГОСТ 21.1701-97_Расширенная_" & rng.Next.ToString & ".xlsx"
                                    If IO.File.Exists(PutchDoc) = True Then
                                        Try
                                            If IO.File.Exists(PutchDoc) = True Then
                                                IO.File.Copy(PutchDoc, DestDoc)
                                            End If
                                        Catch ex As UnauthorizedAccessException
                                            Exit Sub
                                        Catch ex As ArgumentException
                                            Exit Sub
                                        Catch ex As PathTooLongException
                                            Exit Sub
                                        Catch ex As DirectoryNotFoundException
                                            Exit Sub
                                        Catch ex As FileNotFoundException
                                            Exit Sub
                                        Catch ex As IOException
                                            Exit Sub
                                        Catch ex As NotSupportedException
                                            Exit Sub
                                        End Try
                                        xlBook = xlApp.Workbooks.Open(DestDoc)
                                    Else
                                        Dim avFiles As String = xlApp.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Excel файлы", , True)
                                        xlBook = xlApp.Workbooks.Open(avFiles)
                                    End If
                                    'находим нужный лист(Ведомость объемов работ)
                                    Dim SheetCountZu1 As Integer = 1
                                    Dim xlsheet As Excel.Worksheet
                                    For Each xlsheet In xlBook.Worksheets
                                        If xlsheet.Name.ToString.Trim Like "Ведомость углов поворота" Then
                                            SheetCountZu1 = xlsheet.Index
                                            Exit For
                                        End If
                                    Next
                                    xlsheet = xlBook.Worksheets.Item(SheetCountZu1) 'делаем ссылку на этот лист
                                    xlsheet.Activate()
                                    '1. конец записей
                                    Dim exlPos As Integer = 1
                                    For exl As Integer = 1 To 200
                                        Dim exlStr As String = xlsheet.Cells(exl, 1).value
                                        If exlStr Like "Ведомость углов поворота, прямых и кривых" Then
                                            exlPos = exl + 1
                                            Exit For
                                        End If
                                    Next exl
                                    'записываем обозначение участка
                                    xlsheet.Cells(exlPos, 1).value = "Наименование трассы" & Align.Name
                                    xlsheet.Cells(exlPos + 1, 1).value = "Описание трассы" & Align.Description
                                    exlPos = exlPos + 6
                                    '==============================================================================================
                                    'запись значений (начало трассы)
                                    cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    xlsheet.Cells(exlPos, 1) = "НТ"
                                    'проверяем стартовый пикет на рубленность
                                    Dim PiketStartStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKStartAlign, FormRoadPlane.NumericUpDown6.Value, indFormatPK)
                                    'вычленяем километр
                                    Dim KMInt As Integer = Val(PiketStartStrFormat)
                                    Dim KM1 As Integer = 0
                                    If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                        'вычленяем километр
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                        xlsheet.Cells(exlPos, 3) = PiketStartStr
                                    Else
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                        xlsheet.Cells(exlPos, 3) = PiketStartStr
                                    End If
                                    'записываем угол начальный
                                    Dim strAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    xlsheet.Cells(exlPos, 4) = strAngle
                                    xlsheet.Cells(exlPos, 5) = "'-"

                                    xlsheet.Cells(exlPos, 22) = FuncFormatZN.FuncFormatCoordinate(AlignStartY, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    xlsheet.Cells(exlPos, 23) = FuncFormatZN.FuncFormatCoordinate(AlignStartX, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    '===================================================================================================================
                                    'записываем длину прямой вставки
                                    Dim PosStartLg As Integer = exlPos + 1 'номер позиции для начального отрезка
                                    exlPos += 2
                                    If IsArray(ArrayExcel) = True Then
                                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                            Dim Xvu As Double = Val(ArrayExcel(20, i)) 'Y
                                            Dim Yvu As Double = Val(ArrayExcel(19, i)) 'X
                                            Dim ld As Double = startNT.GetVectorTo(New Point2d(Xvu, Yvu)).Length
                                            If ld <= 0.02 Then

                                            End If

                                            cell1 = xlsheet.Cells(exlPos - 1, 1) 'ссылка на первую ячейку
                                            cell2 = xlsheet.Cells(exlPos - 1, 23) 'ссылка на 6 ячейку в строке
                                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                            CellRng.EntireRow.Insert() 'вставляем строку
                                            cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                            cell2 = xlsheet.Cells(exlPos, 23) 'ссылка на 6 ячейку в строке
                                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                            CellRng.EntireRow.Insert() 'вставляем строку
                                            'записываем значения
                                            xlsheet.Cells(exlPos, 1) = "ВУ " & ArrayExcel(0, i) 'вершина угла (номер)
                                            'находим пикет
                                            Dim PiketStr As String = ArrayExcel(1, i)
                                            If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                                PiketStr = FuncFormatZN.FuncFormatPK(Val(PiketStr), FormRoadPlane.NumericUpDown6.Value)
                                                'вычленяем километр
                                                KMInt = Val(PiketStr)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                                xlsheet.Cells(exlPos, 2) = KM1
                                                xlsheet.Cells(exlPos, 3) = PiketStr
                                            Else
                                                KMInt = Val(PiketStr)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                                xlsheet.Cells(exlPos, 2) = KM1
                                                xlsheet.Cells(exlPos, 3) = PiketStr
                                            End If
                                            'угол дельта
                                            Dim Ugol As Double = ArrayExcel(2, i)
                                            Dim AbsUgol As Double = Math.Abs(Ugol)
                                            Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                            If Ugol <= 0 Then
                                                xlsheet.Cells(exlPos, 4) = DAngle
                                                xlsheet.Cells(exlPos, 5) = "'-"
                                            Else
                                                xlsheet.Cells(exlPos, 4) = "'-"
                                                xlsheet.Cells(exlPos, 5) = DAngle
                                            End If
                                            'записываем радиус
                                            Dim rad1 As Double = ArrayExcel(3, i)
                                            Dim rad2 As Double = ArrayExcel(21, i)
                                            Dim rad3 As Double = ArrayExcel(22, i)
                                            Dim StrRAd As String = "" 'результирующая строка
                                            If rad1 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad1 - Int(rad1))
                                                If flrad < 0.01 Then
                                                    StrRAd = Int(rad1) & vbLf 'радиус
                                                Else
                                                    StrRAd = FuncFormatZN.FuncFormatCoordinate(rad1, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad2 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad2 - Int(rad2))
                                                If flrad < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad2) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad2, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad3 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad3 - Int(rad3))
                                                If flrad < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad3) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad3, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            StrRAd = StrRAd.Trim
                                            If StrRAd.Length > 0 Then
                                                xlsheet.Cells(exlPos, 6) = StrRAd.Trim
                                            Else
                                                xlsheet.Cells(exlPos, 6) = 0
                                            End If
                                            'записываем элементы кривой
                                            Dim numbz As Double = ArrayExcel(4, i)
                                            Dim flnumb As Double = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 7) = Int(ArrayExcel(4, i)) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 7) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем L1
                                            End If
                                            numbz = ArrayExcel(5, i) 'L2
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 8) = Int(ArrayExcel(5, i)) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 8) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем L2
                                            End If
                                            numbz = ArrayExcel(6, i) 'T1
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 9) = ArrayExcel(6, i) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 9) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем T1
                                            End If
                                            numbz = ArrayExcel(7, i) 'T2
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 10) = Int(ArrayExcel(7, i)) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 10) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем T2
                                            End If
                                            numbz = ArrayExcel(8, i) 'Kpoln
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 11) = Int(ArrayExcel(8, i)) 'Kpoln
                                            Else
                                                xlsheet.Cells(exlPos, 11) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем Kpoln
                                            End If
                                            numbz = ArrayExcel(9, i) 'Kpoln
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 12) = Int(ArrayExcel(9, i)) 'Ksohr
                                            Else
                                                xlsheet.Cells(exlPos, 12) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем Ksohr
                                            End If
                                            'B
                                            xlsheet.Cells(exlPos, 13) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(10, i), FormRoadPlane.NumericUpDown6.Value)
                                            'D
                                            xlsheet.Cells(exlPos, 14) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(11, i), FormRoadPlane.NumericUpDown6.Value)
                                            'записываем пикетаж
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'начало переходной кривой
                                            Try
                                                Dim NPK As Double = ArrayExcel(12, i) 'RAW пикет начала переходной кривой
                                                If NPK = -10000 Then NPK = ArrayExcel(13, i)
                                                Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NPK, RoundLineZn, indFormatPK)
                                                If PiketNPKStrFormat Like "" Then
                                                    xlsheet.Cells(exlPos, 15) = ArrayExcel(12, i)
                                                Else
                                                    xlsheet.Cells(exlPos, 15) = PiketNPKStrFormat 'NPK
                                                End If
                                            Catch
                                            End Try
                                            'начало круговой кривой
                                            Try
                                                Dim NKK As Double = ArrayExcel(13, i) 'RAW пикет начала круговой кривой
                                                If NKK = -10000 Then NKK = ArrayExcel(12, i)
                                                Dim PiketNKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NKK, RoundLineZn, indFormatPK)
                                                If PiketNKKStrFormat Like "" Then
                                                    xlsheet.Cells(exlPos, 16) = ArrayExcel(13, i) 'NKK
                                                Else
                                                    xlsheet.Cells(exlPos, 16) = PiketNKKStrFormat 'NKK
                                                End If
                                            Catch
                                            End Try
                                            'конец круговой кривой
                                            Try
                                                Dim KKK As Double = ArrayExcel(14, i) 'RAW пикет конца круговой кривой
                                                If KKK = -10000 Then KKK = ArrayExcel(15, i)
                                                Dim PiketKKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KKK, RoundLineZn, indFormatPK)
                                                If PiketKKKStrFormat Like "" Then
                                                    xlsheet.Cells(exlPos, 17) = ArrayExcel(14, i) 'KKK
                                                Else
                                                    xlsheet.Cells(exlPos, 17) = PiketKKKStrFormat 'KKK
                                                End If
                                            Catch
                                            End Try
                                            'конец переходной кривой
                                            Try
                                                Dim KPK As Double = ArrayExcel(15, i) 'RAW пикет конца круговой кривой
                                                If KPK = -10000 Then KPK = ArrayExcel(14, i)
                                                Dim PiketKPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KPK, RoundLineZn, indFormatPK)
                                                If PiketKPKStrFormat Like "" Then
                                                    xlsheet.Cells(exlPos, 18) = ArrayExcel(15, i) 'KPK
                                                Else
                                                    xlsheet.Cells(exlPos, 18) = PiketKPKStrFormat 'KPK
                                                End If
                                            Catch ex As System.ArgumentException
                                            End Try
                                            '=================================================================================================
                                            'записываем длины прямых участков
                                            If i <> 0 Then
                                                xlsheet.Cells(exlPos - 1, 20) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(17, i - 1), RoundLineZn) 'Длина прямого участка
                                            End If
                                            xlsheet.Cells(exlPos, 22) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(20, i), RoundLineZn) 'Y
                                            xlsheet.Cells(exlPos, 23) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(19, i), RoundLineZn) 'X
                                            'рассчитываем рассояния между вершинами
                                            Dim vect As Vector2d = Nothing
                                            Dim XY1 As Point2d = Nothing
                                            Dim XY2 As Point2d = Nothing
                                            If i = 0 Then
                                                XY1 = New Point2d(AlignStartX, AlignStartY)
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            Else
                                                XY1 = New Point2d(ArrayExcel(19, i - 1), ArrayExcel(20, i - 1))
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            End If
                                            vect = XY1.GetVectorTo(XY2)
                                            xlsheet.Cells(exlPos - 1, 19) = FuncFormatZN.FuncFormatCoordinate(vect.Length, RoundLineZn) 'Расстояние между вершинами
                                            'записываем румб
                                            Dim DirectMath As Double = vect.Angle 'угол в радианах (румб)
                                            Dim DirectGeodesic As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath)
                                            Dim Grad As Double = (DirectGeodesic * 180) / Math.PI 'переводим в градусы
                                            xlsheet.Cells(exlPos - 1, 21) = FuncSurvey.FuncDirectToRumbDDMMSS(Grad, FormRoadPlane.NumericUpDown1.Value)
                                            exlPos = exlPos + 2
                                        Next i
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем конец трассы
                                    cell1 = xlsheet.Cells(exlPos - 1, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos - 1, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    xlsheet.Cells(exlPos, 1) = "КТ"
                                    'вычленяем километр
                                    Dim PiketEndStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKEndAlign, RoundLineZn, indFormatPK)
                                    If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                        'вычленяем километр
                                        KMInt = Val(PiketEndStrFormat)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                        xlsheet.Cells(exlPos, 3) = PiketEndStrFormat
                                    Else
                                        KMInt = Val(PiketEndStrFormat)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                        xlsheet.Cells(exlPos, 3) = PiketEndStrFormat
                                    End If
                                    'записываем угол конечный
                                    strAngle = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    xlsheet.Cells(exlPos, 4) = strAngle
                                    xlsheet.Cells(exlPos, 5) = "'-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    xlsheet.Cells(exlPos, 22) = FuncFormatZN.FuncFormatCoordinate(AlignEndY, RoundLineZn) 'записываем координаты вершины
                                    xlsheet.Cells(exlPos, 23) = FuncFormatZN.FuncFormatCoordinate(AlignEndX, RoundLineZn) 'записываем координаты вершины
                                    'рассчитываем расстояния между вершинами
                                    Dim XYk As Point2d = Nothing
                                    If IsArray(ArrayExcel) = True Then
                                        Dim PosC As Integer = ArrayExcel.GetUpperBound(1)
                                        XYk = New Point2d(ArrayExcel(19, PosC), ArrayExcel(20, PosC))
                                    Else
                                        XYk = New Point2d(AlignStartX, AlignStartY)
                                        xlsheet.Cells(exlPos - 1, 6) = "'-" 'радиус
                                        xlsheet.Cells(exlPos - 1, 7) = "'-" 'L1
                                        xlsheet.Cells(exlPos - 1, 8) = "'-" 'L2
                                        xlsheet.Cells(exlPos - 1, 9) = "'-" 'T1
                                        xlsheet.Cells(exlPos - 1, 10) = "'-" 'T2
                                        xlsheet.Cells(exlPos - 1, 11) = "'-" 'Kполн
                                        xlsheet.Cells(exlPos - 1, 12) = "'-" 'Kсохр
                                        xlsheet.Cells(exlPos - 1, 13) = "'-" 'Б
                                        xlsheet.Cells(exlPos - 1, 14) = "'-" 'Д
                                        xlsheet.Cells(exlPos - 1, 15) = "'-" 'НПК
                                        xlsheet.Cells(exlPos - 1, 16) = "'-" 'НКК
                                        xlsheet.Cells(exlPos - 1, 17) = "'-" 'ККК
                                        xlsheet.Cells(exlPos - 1, 18) = "'-" 'КПК
                                    End If
                                    Dim XYk2 As Point2d = New Point2d(AlignEndX, AlignEndY)
                                    Dim vect1 As Vector2d = XYk.GetVectorTo(XYk2)
                                    xlsheet.Cells(exlPos - 1, 19) = FuncFormatZN.FuncFormatCoordinate(vect1.Length, FormRoadPlane.NumericUpDown6.Value) 'Расстояние между вершинами
                                    'записываем румб
                                    Dim DirectMath2 As Double = vect1.Angle 'угол в радианах (румб)
                                    Dim DirectGeodesic2 As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath2)
                                    Dim Grad2 As Double = (DirectGeodesic2 * 180) / Math.PI 'переводим в градусы
                                    xlsheet.Cells(exlPos - 1, 21) = FuncSurvey.FuncDirectToRumbDDMMSS(Grad2, 1)
                                    'записываем длину конечного отрезка
                                    xlsheet.Cells(exlPos - 1, 20) = FuncFormatZN.FuncFormatCoordinate(TLenghtEnd, FormRoadPlane.NumericUpDown6.Value) 'длина конечной прямой
                                    'записываем длину прямой вставки начального отрезка
                                    xlsheet.Cells(PosStartLg, 20) = FuncFormatZN.FuncFormatCoordinate(TLenghtStart, FormRoadPlane.NumericUpDown6.Value) 'длина начальной прямой
                                    xlSheets = Nothing
                                    xlBook = Nothing
                                    xlApp = Nothing
                                    'форма отчета ГОСТ
                                ElseIf IndTableExcel = 1 Then
                                    Dim PutchDoc As String = PutchTables & "Форма 4_ГОСТ 21.1701-97.xlsx"
                                    Dim rng As Random = New Random
                                    Dim DestDoc As String = TempDir & "Форма 4_ГОСТ 21.1701-97.xlsx_" & rng.Next.ToString & ".xlsx"
                                    If IO.File.Exists(PutchDoc) = True Then
                                        Try
                                            If IO.File.Exists(PutchDoc) = True Then
                                                IO.File.Copy(PutchDoc, DestDoc)
                                            End If
                                        Catch ex As UnauthorizedAccessException
                                            Exit Sub
                                        Catch ex As ArgumentException
                                            Exit Sub
                                        Catch ex As PathTooLongException
                                            Exit Sub
                                        Catch ex As DirectoryNotFoundException
                                            Exit Sub
                                        Catch ex As FileNotFoundException
                                            Exit Sub
                                        Catch ex As IOException
                                            Exit Sub
                                        Catch ex As NotSupportedException
                                            Exit Sub
                                        End Try
                                        xlBook = xlApp.Workbooks.Open(DestDoc)
                                    Else
                                        Dim avFiles As String = xlApp.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Excel файлы", , True)
                                        xlBook = xlApp.Workbooks.Open(avFiles)
                                    End If
                                    xlApp.Visible = True
                                    'находим нужный лист(Ведомость объемов работ)
                                    Dim SheetCountZu1 As Integer = 1
                                    Dim xlsheet As Excel.Worksheet
                                    For Each xlsheet In xlBook.Worksheets
                                        If xlsheet.Name.ToString.Trim Like "Ведомость углов поворота" Then
                                            SheetCountZu1 = xlsheet.Index
                                            Exit For
                                        End If
                                    Next
                                    xlsheet = xlBook.Worksheets.Item(SheetCountZu1) 'делаем ссылку на этот лист
                                    xlsheet.Activate()
                                    '1. конец записей
                                    Dim exlPos As Integer = 1
                                    For exl As Integer = 1 To 200
                                        Dim exlStr As String = xlsheet.Cells(exl, 1).value
                                        If exlStr Like "Ведомость углов поворота, прямых и кривых" Then
                                            exlPos = exl + 1
                                            Exit For
                                        End If
                                    Next exl
                                    'записываем обозначение участка
                                    xlsheet.Cells(exlPos, 1).value = "Наименование трассы: " & Align.Name
                                    xlsheet.Cells(exlPos + 1, 1).value = "Описание трассы: " & Align.Description
                                    exlPos = exlPos + 6
                                    '==============================================================================================
                                    'запись значений (начало трассы)
                                    cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    xlsheet.Cells(exlPos, 1) = "НТ"
                                    'проверяем стартовый пикет на рубленность
                                    Dim PiketStartStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKStartAlign, RoundLineZn, indFormatPK)
                                    'вычленяем километр
                                    Dim KMInt As Integer = Val(PiketStartStrFormat)
                                    Dim KM1 As Integer = 0
                                    If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                        'вычленяем километр
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                    Else
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                    End If
                                    'записываем пикет
                                    Dim StrPKArray As String() = PiketStartStrFormat.Replace("+", " ").Split
                                    If IsArray(StrPKArray) = True Then
                                        If StrPKArray.Count > 1 Then
                                            xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                            xlsheet.Cells(exlPos, 4) = StrPKArray(1)
                                        Else
                                            xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                            xlsheet.Cells(exlPos, 4) = 0
                                        End If
                                    End If
                                    'записываем угол начальный
                                    Dim strAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    xlsheet.Cells(exlPos, 5) = strAngle
                                    xlsheet.Cells(exlPos, 6) = "'-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    xlsheet.Cells(exlPos, 25) = FuncFormatZN.FuncFormatCoordinate(AlignStartY, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    xlsheet.Cells(exlPos, 26) = FuncFormatZN.FuncFormatCoordinate(AlignStartX, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    '===================================================================================================================
                                    'записываем длину прямой вставки
                                    Dim PosStartLg As Integer = exlPos + 1 'номер позиции для начального отрезка
                                    exlPos += 2
                                    If IsArray(ArrayExcel) = True Then
                                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                            cell1 = xlsheet.Cells(exlPos - 1, 1) 'ссылка на первую ячейку
                                            cell2 = xlsheet.Cells(exlPos - 1, 26) 'ссылка на 6 ячейку в строке
                                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                            CellRng.EntireRow.Insert() 'вставляем строку
                                            cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                            cell2 = xlsheet.Cells(exlPos, 26) 'ссылка на 6 ячейку в строке
                                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                            CellRng.EntireRow.Insert() 'вставляем строку
                                            'записываем значения
                                            xlsheet.Cells(exlPos, 1) = "ВУ " & ArrayExcel(0, i) 'вершина угла (номер)
                                            'находим пикет
                                            'Dim PiketStr As String = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(1, i), RoundLineZn, indFormatPK)
                                            Dim PiketStr As String = FuncFormatZN.FuncFormatPK(Val(ArrayExcel(1, i)), RoundLineZn)
                                            If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                                Dim PiketStrKM As String = FuncFormatZN.FuncFormatPK(Val(ArrayExcel(1, i)), FormRoadPlane.NumericUpDown6.Value)
                                                'вычленяем километр
                                                KMInt = Val(PiketStrKM)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                                xlsheet.Cells(exlPos, 2) = KM1
                                            Else
                                                Dim PiketStrKM As String = FuncFormatZN.FuncFormatPK(Val(ArrayExcel(1, i)), FormRoadPlane.NumericUpDown6.Value)
                                                KMInt = Val(PiketStrKM)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                                xlsheet.Cells(exlPos, 2) = KM1
                                            End If
                                            'записываем пикет
                                            Erase StrPKArray
                                            StrPKArray = PiketStr.Replace("+", " ").Split
                                            If IsArray(StrPKArray) = True Then
                                                If StrPKArray.Count > 1 Then
                                                    xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                                    xlsheet.Cells(exlPos, 4) = StrPKArray(1)
                                                Else
                                                    xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                                    xlsheet.Cells(exlPos, 4) = 0
                                                End If
                                            End If
                                            'угол дельта
                                            Dim Ugol As Double = ArrayExcel(2, i)
                                            Dim AbsUgol As Double = Math.Abs(Ugol)
                                            Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                            If Ugol <= 0 Then
                                                xlsheet.Cells(exlPos, 5) = DAngle
                                                xlsheet.Cells(exlPos, 6) = "'-"
                                            Else
                                                xlsheet.Cells(exlPos, 5) = "'-"
                                                xlsheet.Cells(exlPos, 6) = DAngle
                                            End If
                                            'записываем радиус
                                            Dim rad1 As Double = ArrayExcel(3, i)
                                            Dim rad2 As Double = ArrayExcel(21, i)
                                            Dim rad3 As Double = ArrayExcel(22, i)
                                            Dim StrRAd As String = "" 'результирующая строка
                                            If rad1 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad1 - Int(rad1))
                                                If flrad < 0.01 Then
                                                    StrRAd = Int(rad1) & vbLf 'радиус
                                                Else
                                                    StrRAd = FuncFormatZN.FuncFormatCoordinate(rad1, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad2 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad2 - Int(rad2))
                                                If flrad < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad2) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad2, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad3 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad3 - Int(rad3))
                                                If flrad < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad3) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad3, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            StrRAd = StrRAd.Trim
                                            If StrRAd.Length > 0 Then
                                                xlsheet.Cells(exlPos, 7) = StrRAd.Trim
                                            Else
                                                xlsheet.Cells(exlPos, 7) = 0
                                            End If

                                            'записываем элементы кривой
                                            Dim L1 As Double = ArrayExcel(4, i)
                                            Dim flnumb As Double = L1 - Int(L1)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 10) = Int(ArrayExcel(4, i)) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 10) = FuncFormatZN.FuncFormatCoordinate(L1, FormRoadPlane.NumericUpDown6.Value) 'записываем L1
                                            End If
                                            Dim L2 As Double = ArrayExcel(5, i) 'L2
                                            flnumb = L2 - Int(L2)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 11) = Int(ArrayExcel(5, i)) 'L2
                                            Else
                                                xlsheet.Cells(exlPos, 11) = FuncFormatZN.FuncFormatCoordinate(L2, FormRoadPlane.NumericUpDown6.Value) 'записываем L2
                                            End If
                                            Dim numbz As Double = ArrayExcel(6, i) 'T1
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 8) = ArrayExcel(6, i) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 8) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем T1
                                            End If
                                            numbz = ArrayExcel(7, i) 'T2
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 9) = Int(ArrayExcel(7, i)) 'L1
                                            Else
                                                xlsheet.Cells(exlPos, 9) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем T2
                                            End If
                                            numbz = ArrayExcel(9, i) 'Kpoln
                                            flnumb = numbz - Int(numbz)
                                            If flnumb = 0 Then
                                                xlsheet.Cells(exlPos, 12) = Int(ArrayExcel(9, i)) 'Ksohr
                                            Else
                                                xlsheet.Cells(exlPos, 12) = FuncFormatZN.FuncFormatCoordinate(numbz, FormRoadPlane.NumericUpDown6.Value) 'записываем Ksohr
                                            End If
                                            'B
                                            xlsheet.Cells(exlPos, 13) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(10, i), FormRoadPlane.NumericUpDown6.Value)
                                            'записываем пикетаж
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'начало переходной кривой
                                            Dim NPK As Double = ArrayExcel(12, i) 'RAW пикет начала переходной кривой
                                            Dim NKK As Double = ArrayExcel(13, i) 'RAW пикет начала круговой кривой
                                            Dim KKK As Double = ArrayExcel(14, i) 'RAW пикет конца круговой кривой
                                            Dim KPK As Double = ArrayExcel(15, i) 'RAW пикет конца круговой кривой
                                            If L1 > 0 Then
                                                Try
                                                    If NPK <> -10000 Then
                                                        Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NPK, RoundLineZn, indFormatPK)
                                                        If PiketNPKStrFormat Like "" Then
                                                            xlsheet.Cells(exlPos, 14) = ArrayExcel(12, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketNPKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    xlsheet.Cells(exlPos, 14) = StrPKArray(0)
                                                                    xlsheet.Cells(exlPos, 15) = StrPKArray(1)
                                                                Else
                                                                    xlsheet.Cells(exlPos, 14) = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        xlsheet.Cells(exlPos, 14) = "'-"
                                                        xlsheet.Cells(exlPos, 15) = "'-"
                                                    End If
                                                Catch
                                                End Try
                                                'начало круговой кривой
                                                Try
                                                    If NKK <> -10000 Then
                                                        Dim PiketNKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NKK, RoundLineZn, indFormatPK)
                                                        If PiketNKKStrFormat Like "" Then
                                                            xlsheet.Cells(exlPos, 16) = ArrayExcel(13, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketNKKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    xlsheet.Cells(exlPos, 16) = StrPKArray(0)
                                                                    xlsheet.Cells(exlPos, 17) = StrPKArray(1)
                                                                Else
                                                                    xlsheet.Cells(exlPos, 16) = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        xlsheet.Cells(exlPos, 16) = "'-"
                                                        xlsheet.Cells(exlPos, 17) = "'-"
                                                    End If
                                                Catch
                                                End Try
                                            Else
                                                xlsheet.Cells(exlPos, 14) = "'-"
                                                xlsheet.Cells(exlPos, 15) = "'-"
                                                xlsheet.Cells(exlPos, 16) = "'-"
                                                xlsheet.Cells(exlPos, 17) = "'-"
                                            End If
                                            'конец круговой кривой
                                            If L2 > 0 Then
                                                Try
                                                    If KKK <> -10000 Then
                                                        Dim PiketKKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KKK, RoundLineZn, indFormatPK)
                                                        If PiketKKKStrFormat Like "" Then
                                                            xlsheet.Cells(exlPos, 18) = ArrayExcel(14, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketKKKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    xlsheet.Cells(exlPos, 18) = StrPKArray(0)
                                                                    xlsheet.Cells(exlPos, 19) = StrPKArray(1)
                                                                Else
                                                                    xlsheet.Cells(exlPos, 18) = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        xlsheet.Cells(exlPos, 18) = "'-"
                                                        xlsheet.Cells(exlPos, 19) = "'-"
                                                    End If
                                                Catch
                                                End Try
                                                'конец переходной кривой
                                                Try
                                                    If KPK <> -10000 Then
                                                        Dim PiketKPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KPK, RoundLineZn, indFormatPK)
                                                        If PiketKPKStrFormat Like "" Then
                                                            xlsheet.Cells(exlPos, 20) = ArrayExcel(15, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketKPKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    xlsheet.Cells(exlPos, 20) = StrPKArray(0)
                                                                    xlsheet.Cells(exlPos, 21) = StrPKArray(1)
                                                                Else
                                                                    xlsheet.Cells(exlPos, 20) = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        xlsheet.Cells(exlPos, 20) = "'-"
                                                        xlsheet.Cells(exlPos, 21) = "'-"
                                                    End If
                                                Catch
                                                End Try
                                            Else
                                                xlsheet.Cells(exlPos, 18) = "'-"
                                                xlsheet.Cells(exlPos, 19) = "'-"
                                                xlsheet.Cells(exlPos, 20) = "'-"
                                                xlsheet.Cells(exlPos, 21) = "'-"
                                            End If

                                            '=================================================================================================
                                            'записываем длины прямых участков
                                            If i <> 0 Then
                                                xlsheet.Cells(exlPos - 1, 23) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(17, i - 1), FormRoadPlane.NumericUpDown6.Value) 'Длина прямого участка
                                            End If
                                            xlsheet.Cells(exlPos, 25) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(20, i), FormRoadPlane.NumericUpDown6.Value) 'Y
                                            xlsheet.Cells(exlPos, 26) = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(19, i), FormRoadPlane.NumericUpDown6.Value) 'X
                                            'рассчитываем рассояния между вершинами
                                            Dim vect As Vector2d = Nothing
                                            Dim XY1 As Point2d = Nothing
                                            Dim XY2 As Point2d = Nothing
                                            If i = 0 Then
                                                XY1 = New Point2d(AlignStartX, AlignStartY)
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            Else
                                                XY1 = New Point2d(ArrayExcel(19, i - 1), ArrayExcel(20, i - 1))
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            End If
                                            vect = XY1.GetVectorTo(XY2)
                                            xlsheet.Cells(exlPos - 1, 22) = FuncFormatZN.FuncFormatCoordinate(vect.Length, FormRoadPlane.NumericUpDown6.Value) 'Расстояние между вершинами
                                            'записываем румб
                                            Dim DirectMath As Double = vect.Angle 'угол в радианах (румб)
                                            Dim DirectGeodesic As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath)
                                            Dim Grad As Double = (DirectGeodesic * 180) / Math.PI 'переводим в градусы
                                            xlsheet.Cells(exlPos - 1, 24) = FuncSurvey.FuncDirectToRumbDDMMSS(Grad, FormRoadPlane.NumericUpDown1.Value)
                                            exlPos = exlPos + 2
                                        Next i
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем конец трассы
                                    cell1 = xlsheet.Cells(exlPos - 1, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos - 1, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(exlPos, 23) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.EntireRow.Insert() 'вставляем строку
                                    xlsheet.Cells(exlPos, 1) = "КТ"
                                    'вычленяем километр
                                    Dim PiketEndStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKEndAlign, RoundLineZn, indFormatPK)
                                    If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                        'вычленяем километр
                                        KMInt = Val(PiketEndStrFormat)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                    Else
                                        KMInt = Val(PiketEndStrFormat)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                        xlsheet.Cells(exlPos, 2) = KM1
                                    End If
                                    'записываем пикет
                                    Erase StrPKArray
                                    StrPKArray = PiketEndStrFormat.Replace("+", " ").Split
                                    If IsArray(StrPKArray) = True Then
                                        If StrPKArray.Count > 1 Then
                                            xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                            xlsheet.Cells(exlPos, 4) = StrPKArray(1)
                                        Else
                                            xlsheet.Cells(exlPos, 3) = StrPKArray(0)
                                            xlsheet.Cells(exlPos, 4) = 0
                                        End If
                                    End If
                                    'записываем угол конечный
                                    strAngle = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    xlsheet.Cells(exlPos, 5) = strAngle
                                    xlsheet.Cells(exlPos, 6) = "'-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    xlsheet.Cells(exlPos, 25) = FuncFormatZN.FuncFormatCoordinate(AlignEndY, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    xlsheet.Cells(exlPos, 26) = FuncFormatZN.FuncFormatCoordinate(AlignEndX, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    'рассчитываем расстояния между вершинами
                                    Dim XYk As Point2d = Nothing
                                    If IsArray(ArrayExcel) = True Then
                                        Dim PosC As Integer = ArrayExcel.GetUpperBound(1)
                                        XYk = New Point2d(ArrayExcel(19, PosC), ArrayExcel(20, PosC))
                                    Else
                                        XYk = New Point2d(AlignStartX, AlignStartY)
                                        xlsheet.Cells(exlPos - 1, 7) = "'-" 'радиус
                                        xlsheet.Cells(exlPos - 1, 8) = "'-" 'T1
                                        xlsheet.Cells(exlPos - 1, 9) = "'-" 'T2
                                        xlsheet.Cells(exlPos - 1, 10) = "'-" 'L1
                                        xlsheet.Cells(exlPos - 1, 11) = "'-" 'L2
                                        xlsheet.Cells(exlPos - 1, 12) = "'-" 'K
                                        xlsheet.Cells(exlPos - 1, 13) = "'-" 'Б
                                        xlsheet.Cells(exlPos - 1, 14) = "'-" 'НПК
                                        xlsheet.Cells(exlPos - 1, 15) = "'-" 'НПК
                                        xlsheet.Cells(exlPos - 1, 16) = "'-" 'НКК
                                        xlsheet.Cells(exlPos - 1, 17) = "'-" 'НКК
                                        xlsheet.Cells(exlPos - 1, 18) = "'-" 'ККК
                                        xlsheet.Cells(exlPos - 1, 19) = "'-" 'ККК
                                        xlsheet.Cells(exlPos - 1, 20) = "'-" 'КПК
                                        xlsheet.Cells(exlPos - 1, 21) = "'-" 'КПК
                                    End If
                                    Dim XYk2 As Point2d = New Point2d(AlignEndX, AlignEndY)
                                    Dim vect1 As Vector2d = XYk.GetVectorTo(XYk2)
                                    xlsheet.Cells(exlPos - 1, 22) = FuncFormatZN.FuncFormatCoordinate(vect1.Length, FormRoadPlane.NumericUpDown6.Value) 'Расстояние между вершинами
                                    'записываем румб
                                    Dim DirectMath2 As Double = vect1.Angle 'угол в радианах (румб)
                                    Dim DirectGeodesic2 As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath2)
                                    Dim Grad2 As Double = (DirectGeodesic2 * 180) / Math.PI 'переводим в градусы
                                    xlsheet.Cells(exlPos - 1, 24) = FuncSurvey.FuncDirectToRumbDDMMSS(Grad2, 1)
                                    'записываем длину конечного отрезка
                                    xlsheet.Cells(exlPos - 1, 23) = FuncFormatZN.FuncFormatCoordinate(TLenghtEnd, FormRoadPlane.NumericUpDown6.Value) 'длина конечной прямой
                                    'записываем длину прямой вставки начального отрезка
                                    xlsheet.Cells(PosStartLg, 23) = FuncFormatZN.FuncFormatCoordinate(TLenghtStart, FormRoadPlane.NumericUpDown6.Value) 'длина начальной прямой
                                    xlSheets = Nothing
                                    xlBook = Nothing
                                    xlApp = Nothing
                                End If
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'ТАБЛИЦА АВТОКАД
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        ElseIf FormRoadPlane.ComboBox10.SelectedIndex = 1 Then
                            'СТИЛЬ ТАБЛИЦЫ ПО УМОЛЧАНИЮ
                            Dim IDTablelStyleDraw As ObjectId = acCurDb.Tablestyle
                            'СЛОЙ ПО УМОЛЧАНИЮ
                            Dim NameLayerTlb As String = "0"
                            Dim IndTableACAD As Integer = FormRoadPlane.ComboBox4.SelectedIndex
                            '========================================================================================================================
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'стиль расширенный
                            If IndTableACAD = 0 Then
                                'копируем стиль таблицы
                                If IO.File.Exists(PatchTemplateDWG) = True Then
                                    'копируем характеристики таблицы из строки текста
                                    Dim ArrayPropertiesTableRow As String(,) = Nothing
                                    Dim boolPropTable As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PatchTemplateDWG, "Ведомость углов поворота прямых и кривых", ArrayPropertiesTableRow, "Таблица", "Трасса автомобильной дороги")
                                    Dim NameTableTemplates As String = ""
                                    If IsArray(ArrayPropertiesTableRow) = True Then
                                        For i As Integer = 0 To ArrayPropertiesTableRow.GetUpperBound(1)
                                            Dim NameField As String = ArrayPropertiesTableRow(0, i)
                                            If NameField.Trim Like "Имя объекта" Then
                                                NameTableTemplates = ArrayPropertiesTableRow(1, i)
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                    If NameTableTemplates.Trim.Length > 0 Then
                                        Dim boolCopyTextStyle As Boolean = AcadFunction.FuncCreateTableStyleDestDraw(PatchTemplateDWG, NameTableTemplates)
                                        Dim TblDictionary As DBDictionary = tr.GetObject(acCurDb.TableStyleDictionaryId, OpenMode.ForWrite)
                                        If TblDictionary.Contains(NameTableTemplates) Then
                                            IDTablelStyleDraw = TblDictionary.GetAt(NameTableTemplates)
                                        Else
                                            MsgBox("Шаблон ведомости не найден!!! Макро прервано.")
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox("Шаблон ведомости не найден!!! Макро прервано.")
                                        Exit Sub
                                    End If
                                    '===============================================================================================
                                    Dim bt As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                                    Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                                    pPtOpts.Message = vbLf & "Укажите точку вставки таблицы: "
                                    Dim pPtRes As PromptPointResult = acDoc.Editor.GetPoint(pPtOpts)
                                    Dim ptStartTlb As Point3d = pPtRes.Value
                                    Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                                    Dim ts As Autodesk.AutoCAD.DatabaseServices.TableStyle = tr.GetObject(IDTablelStyleDraw, OpenMode.ForRead)
                                    Dim template As Autodesk.AutoCAD.DatabaseServices.TableTemplate = tr.GetObject(ts.Template, OpenMode.ForRead)
                                    tb.CopyFrom(template, TableCopyOptions.FillTarget)
                                    tb.TableStyle = IDTablelStyleDraw
                                    tb.SetDatabaseDefaults()
                                    ms.UpgradeOpen()
                                    ms.AppendEntity(tb)
                                    tb.Position = ptStartTlb
                                    tr.AddNewlyCreatedDBObject(tb, True)
                                    Dim HRowsTbl As Double = 5
                                    Dim HTextTbl As Double = 2
                                    Dim ColorTextTbl As Integer = 256
                                    If IsArray(ArrayPropertiesTableRow) = True Then
                                        For i As Integer = 0 To ArrayPropertiesTableRow.GetUpperBound(1)
                                            Dim NameField As String = ArrayPropertiesTableRow(0, i)
                                            If NameField.Trim Like "Наименование слоя" Then
                                                Dim NameLayer As String = ArrayPropertiesTableRow(1, i)
                                                If IsNothing(NameLayer) = False Then
                                                    If NameLayer.Trim.Length > 0 Then
                                                        Dim acLyrTbl As LayerTable = tr.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                                        If acLyrTbl.Has(NameLayer) = True Then
                                                            tb.Layer = NameLayer
                                                        Else
                                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayer)
                                                            If acLyrTbl.Has(NameLayer) = True Then
                                                                tb.Layer = NameLayer
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Цвет" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    Dim ColorBlk As Integer = Val(ArrayPropertiesTableRow(1, i))
                                                    If ColorBlk < 0 Then
                                                        ColorBlk = 0
                                                    ElseIf ColorBlk > 256 Then
                                                        ColorBlk = 256
                                                    End If
                                                    tb.ColorIndex = ColorBlk
                                                End If
                                            ElseIf NameField.Trim Like "Тип линии" Then
                                                Dim NameTypeLine As String = ArrayPropertiesTableRow(1, i)
                                                If IsNothing(NameTypeLine) = False Then
                                                    If NameTypeLine.Trim.Length > 0 Then
                                                        If Not (NameTypeLine Like "ПоСлою" Or NameTypeLine Like "ПоБлоку") Then
                                                            Dim acLineTypTbl As LinetypeTable = tr.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                            If acLineTypTbl.Has(NameTypeLine) = True Then
                                                                tb.Linetype = NameTypeLine
                                                            Else
                                                                Dim boolTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PatchTemplateDWG, NameTypeLine)
                                                                If acLineTypTbl.Has(NameTypeLine) = True Then
                                                                    tb.Linetype = NameTypeLine
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Вес линии" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    Dim WidthLin As Integer = Val(ArrayPropertiesTableRow(1, i))
                                                    If WidthLin > 0 Then
                                                        tb.LineWeight = WidthLin
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Высота текста" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    HTextTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HTextTbl <= 0 Then
                                                        HTextTbl = 2
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Цвет текста" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    ColorTextTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HTextTbl <= 0 Then
                                                        ColorTextTbl = 256
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Высота строки" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    HRowsTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HRowsTbl <= 0 Then
                                                        HRowsTbl = 5
                                                    End If
                                                End If
                                            End If
                                        Next i
                                    End If
                                    'смотрим количество строк
                                    Dim RowsTbl As Integer = tb.Rows.Count
                                    Dim NumRows As Integer = RowsTbl - 1
                                    tb.Rows(NumRows).Height = HRowsTbl
                                    tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                    tb.Rows(NumRows).TextHeight = HTextTbl
                                    '==========================================================================================================
                                    'записываем значение
                                    tb.Cells(NumRows, 0).TextString = "НТ"
                                    'преобразовываем его в форматированный пикет
                                    Dim PiketStartStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKStartAlign, RoundLineZn, indFormatPK)
                                    If PiketStartStrFormat.Trim.Length = 0 Then
                                        Try
                                            PiketStartStrFormat = Align.GetStationStringWithEquations(PKStartAlign)
                                        Catch ex As ArgumentException
                                        End Try
                                    End If
                                    '==================================================================================================================
                                    'начальное значение километра
                                    Dim PK As Integer = -1
                                    Dim Plus As Double = 0
                                    If PiketStartStrFormat.Trim.Length > 0 Then
                                        Dim posPlus As Integer = PiketStartStrFormat.IndexOf("+")
                                        If posPlus > -1 Then
                                            PK = Val(Mid(PiketStartStrFormat, 1, posPlus))
                                            Plus = Val(Mid(PiketStartStrFormat, posPlus + 1))
                                        End If
                                    Else
                                        PK = Int(PKStartAlign) / 100
                                        Plus = PKStartAlign - PK * 100
                                    End If
                                    tb.Cells(NumRows, 1).TextString = Int(PK / 10)
                                    tb.Cells(NumRows, 2).TextString = PK
                                    tb.Cells(NumRows, 3).TextString = Plus
                                    '===================================================================================================================
                                    'записываем угол начальный
                                    Dim strAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, RoundAngleZn)
                                    tb.Cells(NumRows, 4).TextString = strAngle
                                    tb.Cells(NumRows, 5).TextString = "-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    tb.Cells(NumRows, 22).TextString = FuncFormatZN.FuncFormatCoordinate(AlignStartY, RoundLineZn) 'записываем координаты вершины X
                                    tb.Cells(NumRows, 23).TextString = FuncFormatZN.FuncFormatCoordinate(AlignStartX, RoundLineZn) 'записываем координаты вершины Y
                                    '===================================================================================================================
                                    'запоминаем позицию для вставки значения "длина прямой вставки"
                                    Dim PosStartLg As Integer = NumRows + 1 'номер позиции для начального отрезка
                                    NumRows += 2
                                    If IsArray(ArrayExcel) = True Then
                                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                            tb.InsertRows(NumRows, HRowsTbl, 2)
                                            '=======================================================================================================
                                            'форматируем строку
                                            tb.Rows(NumRows).Height = HRowsTbl
                                            tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                            tb.Rows(NumRows).TextHeight = HTextTbl
                                            tb.Rows(NumRows - 1).Height = HRowsTbl
                                            tb.Rows(NumRows - 1).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                            tb.Rows(NumRows - 1).TextHeight = HTextTbl
                                            '========================================================================================================
                                            'записываем номер вершины угла
                                            Dim NumbVertex As String = ArrayExcel(0, i)
                                            If IsNothing(NumbVertex) = False Then
                                                tb.Cells(NumRows, 0).TextString = "ВУ " & ArrayExcel(0, i) 'вершина угла (номер)
                                            End If
                                            '=======================================================================================================
                                            'находим пикет вершины угла
                                            Dim boolRez As Boolean = False
                                            Dim PiketStr As String = ArrayExcel(1, i)
                                            If IsNothing(PiketStr) = False Then
                                                Dim PKDbl As Double = Val(PiketStr)
                                                PiketStr = FuncAlignment.FuncFormatRublPiket(Align, PKDbl, RoundLineZn, 0)
                                                If PiketStr.Trim.Length = 0 Then
                                                    Try
                                                        PiketStr = Align.GetStationStringWithEquations(PKDbl)
                                                    Catch ex As System.ArgumentException
                                                        PiketStr = FuncFormatZN.FuncFormatPK(PKDbl, RoundLineZn)
                                                        boolRez = True
                                                    End Try
                                                End If
                                                '===============================================================================================
                                                'отделяем целую часть от дробной
                                                Dim ArrayPKSplit As String() = PiketStr.Split("+")
                                                If IsArray(ArrayPKSplit) = True Then
                                                    PK = Val(ArrayPKSplit(0))
                                                    If ArrayPKSplit.Length > 1 Then
                                                        Plus = Val(ArrayPKSplit(1))
                                                    Else
                                                        Plus = 0
                                                    End If
                                                Else
                                                    PK = 0
                                                    Plus = 0
                                                End If
                                                'вычленяем километр
                                                Dim KM11 As Integer = MathFunction.FuncTrimDigitFloo((PK) / 10, 0)
                                                tb.Cells(NumRows, 1).TextString = KM11
                                                'записываем пикет
                                                tb.Cells(NumRows, 2).TextString = PK
                                                tb.Cells(NumRows, 3).TextString = Plus
                                            End If
                                            '=============================================================================================================
                                            'угол дельта
                                            Dim Ugol As Double = ArrayExcel(2, i)
                                            Dim AbsUgol As Double = Math.Abs(Ugol)
                                            Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, RoundAngleZn)
                                            If Ugol <= 0 Then
                                                tb.Cells(NumRows, 4).TextString = DAngle
                                                tb.Cells(NumRows, 5).TextString = "-"
                                            Else
                                                tb.Cells(NumRows, 5).TextString = DAngle
                                                tb.Cells(NumRows, 4).TextString = "-"
                                            End If
                                            '==================================================================================================================
                                            'записываем радиус
                                            Dim rad1 As Double = ArrayExcel(3, i)
                                            Dim rad2 As Double = ArrayExcel(21, i)
                                            Dim rad3 As Double = ArrayExcel(22, i)
                                            Dim StrRAd As String = "" 'результирующая строка
                                            If rad1 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad1 - Int(rad1))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = Int(rad1) & vbLf 'радиус
                                                Else
                                                    StrRAd = FuncFormatZN.FuncFormatCoordinate(rad1, RoundLineZn) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad2 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad2 - Int(rad2))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad2) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad2, RoundLineZn) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad3 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad3 - Int(rad3))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad3) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad3, RoundLineZn) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            StrRAd = StrRAd.Trim
                                            If StrRAd.Length > 0 Then
                                                tb.Cells(NumRows, 6).TextString = StrRAd.Trim
                                            Else
                                                tb.Cells(NumRows, 6).TextString = "0"
                                            End If
                                            '============================================================================================================================
                                            'записываем элементы кривой
                                            Dim numbz As Double = ArrayExcel(4, i)
                                            Dim flnumb As Double = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 7).TextString = Int(ArrayExcel(4, i)) 'L1
                                            Else
                                                tb.Cells(NumRows, 7).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем L1
                                            End If
                                            numbz = ArrayExcel(5, i) 'L2
                                            flnumb = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 8).TextString = Int(ArrayExcel(5, i)) 'L2
                                            Else
                                                tb.Cells(NumRows, 8).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем L2
                                            End If
                                            numbz = ArrayExcel(6, i) 'T1
                                            flnumb = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 9).TextString = Int(ArrayExcel(6, i)) 'T1
                                            Else
                                                tb.Cells(NumRows, 9).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем T1
                                            End If
                                            numbz = ArrayExcel(7, i) 'T2
                                            flnumb = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 10).TextString = Int(ArrayExcel(7, i)) 'L1
                                            Else
                                                tb.Cells(NumRows, 10).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем T2
                                            End If
                                            numbz = ArrayExcel(8, i) 'Kpoln
                                            flnumb = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 11).TextString = Int(ArrayExcel(8, i)) 'Kpoln
                                            Else
                                                tb.Cells(NumRows, 11).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем Kpoln
                                            End If
                                            numbz = ArrayExcel(9, i) 'Kpoln
                                            flnumb = Math.Abs(numbz - Int(numbz))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 12).TextString = Int(ArrayExcel(9, i)) 'Ksohr
                                            Else
                                                tb.Cells(NumRows, 12).TextString = FuncFormatZN.FuncFormatCoordinate(numbz, RoundLineZn) 'записываем Ksohr
                                            End If
                                            'B
                                            tb.Cells(NumRows, 13).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(10, i), RoundLineZn)
                                            'D
                                            tb.Cells(NumRows, 14).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(11, i), RoundLineZn)
                                            'записываем пикетаж
                                            '=================================================================================================================
                                            'начало переходной кривой
                                            Try
                                                Dim NPK As Double = ArrayExcel(12, i) 'RAW пикет начала переходной кривой
                                                If NPK = -10000 Then NPK = ArrayExcel(13, i)
                                                Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NPK, RoundLineZn, indFormatPK)
                                                If PiketNPKStrFormat Like "" Then
                                                    tb.Cells(NumRows, 15).TextString = ArrayExcel(12, i)
                                                Else
                                                    tb.Cells(NumRows, 15).TextString = PiketNPKStrFormat 'NPK
                                                End If
                                            Catch
                                            End Try
                                            'начало круговой кривой
                                            Try
                                                Dim NKK As Double = ArrayExcel(13, i) 'RAW пикет начала круговой кривой
                                                If NKK = -10000 Then NKK = ArrayExcel(12, i)
                                                Dim PiketNKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NKK, RoundLineZn, indFormatPK)
                                                If PiketNKKStrFormat Like "" Then
                                                    tb.Cells(NumRows, 16).TextString = ArrayExcel(13, i) 'NKK
                                                Else
                                                    tb.Cells(NumRows, 16).TextString = PiketNKKStrFormat 'NKK
                                                End If
                                            Catch
                                            End Try
                                            'конец круговой кривой
                                            Try
                                                Dim KKK As Double = ArrayExcel(14, i) 'RAW пикет конца круговой кривой
                                                If KKK = -10000 Then KKK = ArrayExcel(15, i)
                                                Dim PiketKKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KKK, RoundLineZn, indFormatPK)
                                                If PiketKKKStrFormat Like "" Then
                                                    tb.Cells(NumRows, 17).TextString = ArrayExcel(14, i) 'KKK
                                                Else
                                                    tb.Cells(NumRows, 17).TextString = PiketKKKStrFormat 'KKK
                                                End If
                                            Catch
                                            End Try
                                            'конец переходной кривой
                                            Try
                                                Dim KPK As Double = ArrayExcel(15, i) 'RAW пикет конца круговой кривой
                                                If KPK = -10000 Then KPK = ArrayExcel(14, i)
                                                Dim PiketKPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KPK, RoundLineZn, indFormatPK)
                                                If PiketKPKStrFormat Like "" Then
                                                    tb.Cells(NumRows, 18).TextString = ArrayExcel(15, i) 'KPK
                                                Else
                                                    tb.Cells(NumRows, 18).TextString = PiketKPKStrFormat 'KPK
                                                End If
                                            Catch ex As System.ArgumentException
                                            End Try
                                            '=================================================================================================
                                            'записываем длины прямых участков (21 столбец)
                                            If i <> 0 Then
                                                tb.Cells(NumRows - 1, 20).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(17, i - 1), RoundLineZn) 'Длина прямого участка
                                            End If
                                            '===================================================================================================================================
                                            'записываем координаты
                                            tb.Cells(NumRows, 22).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(20, i), RoundLineZn) 'Y
                                            tb.Cells(NumRows, 23).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(19, i), RoundLineZn) 'X
                                            '=====================================================================================================================
                                            'рассчитываем рассояния между вершинами
                                            Dim vect As Vector2d = Nothing
                                            Dim XY1 As Point2d = Nothing
                                            Dim XY2 As Point2d = Nothing
                                            If i = 0 Then
                                                XY1 = New Point2d(AlignStartX, AlignStartY)
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            Else
                                                XY1 = New Point2d(ArrayExcel(19, i - 1), ArrayExcel(20, i - 1))
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            End If
                                            vect = XY1.GetVectorTo(XY2)
                                            tb.Cells(NumRows - 1, 19).TextString = FuncFormatZN.FuncFormatCoordinate(vect.Length, RoundLineZn) 'Расстояние между вершинами
                                            'записываем румб
                                            Dim DirectMath As Double = vect.Angle 'угол в радианах (румб)
                                            Dim DirectGeodesic As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath)
                                            Dim Grad As Double = (DirectGeodesic * 180) / Math.PI 'переводим в градусы
                                            tb.Cells(NumRows - 1, 21).TextString = FuncSurvey.FuncDirectToRumbDDMMSS(Grad, RoundAngleZn)
                                            NumRows += 2
                                        Next i
                                    End If
                                    '=================================================================================================================================
                                    'записываем конец трассы
                                    tb.InsertRows(NumRows, HRowsTbl, 2)
                                    tb.Rows(NumRows).Height = HRowsTbl
                                    tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                    tb.Rows(NumRows).TextHeight = HTextTbl
                                    tb.Cells(NumRows, 0).TextString = "КТ"
                                    'преобразовываем его в форматированный пикет
                                    Dim PiketEndStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKEndAlign, RoundLineZn, 0)
                                    If PiketEndStrFormat.Trim.Length = 0 Then
                                        Try
                                            PiketEndStrFormat = Align.GetStationStringWithEquations(PKEndAlign)
                                        Catch ex As System.ArgumentException
                                            PiketEndStrFormat = FuncFormatZN.FuncFormatPK(PKEndAlign, RoundLineZn)
                                        End Try
                                    End If
                                    '===============================================================================================
                                    'отделяем целую часть от дробной
                                    Dim ArrayPKSplitEnd As String() = PiketEndStrFormat.Split("+")
                                    If IsArray(ArrayPKSplitEnd) = True Then
                                        PK = Val(ArrayPKSplitEnd(0))
                                        If ArrayPKSplitEnd.Length > 1 Then
                                            Plus = Val(ArrayPKSplitEnd(1))
                                        Else
                                            Plus = 0
                                        End If
                                    Else
                                        PK = 0
                                        Plus = 0
                                    End If
                                    'вычленяем километр
                                    Dim KMEnd As Integer = MathFunction.FuncTrimDigitFloo((PK) / 10, 0)
                                    tb.Cells(NumRows, 1).TextString = KMEnd
                                    'записываем пикет
                                    tb.Cells(NumRows, 2).TextString = PK
                                    tb.Cells(NumRows, 3).TextString = Plus
                                    'записываем угол конечный
                                    strAngle = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, RoundAngleZn)
                                    tb.Cells(NumRows, 4).TextString = strAngle
                                    tb.Cells(NumRows, 5).TextString = "-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    'записываем координаты вершины угла
                                    tb.Cells(NumRows, 22).TextString = FuncFormatZN.FuncFormatCoordinate(AlignEndY, RoundLineZn) 'записываем координаты вершины
                                    tb.Cells(NumRows, 23).TextString = FuncFormatZN.FuncFormatCoordinate(AlignEndX, RoundLineZn) 'записываем координаты вершины
                                    Dim XYk As Point2d = Nothing
                                    If IsArray(ArrayExcel) = True Then
                                        Dim PosC As Integer = ArrayExcel.GetUpperBound(1)
                                        XYk = New Point2d(ArrayExcel(19, PosC), ArrayExcel(20, PosC))
                                    Else
                                        XYk = New Point2d(AlignStartX, AlignStartY)
                                        tb.Cells(NumRows - 1, 6).TextString = "-" 'радиус
                                        tb.Cells(NumRows - 1, 7).TextString = "-" 'L1
                                        tb.Cells(NumRows - 1, 8).TextString = "-" 'L2
                                        tb.Cells(NumRows - 1, 9).TextString = "-" 'T1
                                        tb.Cells(NumRows - 1, 10).TextString = "-" 'T2
                                        tb.Cells(NumRows - 1, 11).TextString = "-" 'Kполн
                                        tb.Cells(NumRows - 1, 12).TextString = "-" 'Kсохр
                                        tb.Cells(NumRows - 1, 13).TextString = "-" 'Б
                                        tb.Cells(NumRows - 1, 14).TextString = "-" 'Д
                                        tb.Cells(NumRows - 1, 15).TextString = "-" 'НПК
                                        tb.Cells(NumRows - 1, 16).TextString = "-" 'НКК
                                        tb.Cells(NumRows - 1, 17).TextString = "-" 'ККК
                                        tb.Cells(NumRows - 1, 18).TextString = "-" 'КПК
                                    End If
                                    Dim XYk2 As Point2d = New Point2d(AlignEndX, AlignEndY)
                                    Dim vect1 As Vector2d = XYk.GetVectorTo(XYk2)
                                    tb.Cells(NumRows - 1, 19).TextString = FuncFormatZN.FuncFormatCoordinate(vect1.Length, RoundLineZn) 'Расстояние между вершинами
                                    'записываем румб
                                    Dim DirectMath2 As Double = vect1.Angle 'угол в радианах (румб)
                                    Dim DirectGeodesic2 As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath2)
                                    Dim Grad2 As Double = (DirectGeodesic2 * 180) / Math.PI 'переводим в градусы
                                    tb.Cells(NumRows - 1, 21).TextString = FuncSurvey.FuncDirectToRumbDDMMSS(Grad2, 1)
                                    'записываем длину конечного отрезка
                                    tb.Cells(NumRows - 1, 20).TextString = FuncFormatZN.FuncFormatCoordinate(TLenghtEnd, RoundLineZn) 'длина конечной прямой
                                    'записываем длину прямой вставки начального отрезка
                                    tb.Cells(PosStartLg, 20).TextString = FuncFormatZN.FuncFormatCoordinate(TLenghtStart, RoundLineZn) 'длина начальной прямой
                                End If
                                '======================================================================================================================
                                'ПО ГОСТУ оформление
                            ElseIf IndTableACAD = 1 Then
                                'копируем стиль таблицы
                                If IO.File.Exists(PatchTemplateDWG) = True Then
                                    'копируем характеристики таблицы из строки текста
                                    Dim ArrayPropertiesTableRow As String(,) = Nothing
                                    Dim boolPropTable As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PatchTemplateDWG, "Ведомость углов поворота прямых и кривых ГОСТ 21.1701-97", ArrayPropertiesTableRow, "Таблица", "Трасса автомобильной дороги")
                                    Dim NameTableTemplates As String = ""
                                    If IsArray(ArrayPropertiesTableRow) = True Then
                                        For i As Integer = 0 To ArrayPropertiesTableRow.GetUpperBound(1)
                                            Dim NameField As String = ArrayPropertiesTableRow(0, i)
                                            If NameField.Trim Like "Имя объекта" Then
                                                NameTableTemplates = ArrayPropertiesTableRow(1, i)
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                    If NameTableTemplates.Trim.Length > 0 Then
                                        Dim boolCopyTextStyle As Boolean = AcadFunction.FuncCreateTableStyleDestDraw(PatchTemplateDWG, NameTableTemplates)
                                        Dim TblDictionary As DBDictionary = tr.GetObject(acCurDb.TableStyleDictionaryId, OpenMode.ForWrite)
                                        If TblDictionary.Contains(NameTableTemplates) Then
                                            IDTablelStyleDraw = TblDictionary.GetAt(NameTableTemplates)
                                        Else
                                            MsgBox("Шаблон ведомости не найден!!! Макро прервано.")
                                            Exit Sub
                                        End If
                                    Else
                                        MsgBox("Шаблон ведомости не найден!!! Макро прервано.")
                                        Exit Sub
                                    End If
                                    '===============================================================================================
                                    Dim bt As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                                    Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                                    pPtOpts.Message = vbLf & "Укажите точку вставки таблицы: "
                                    Dim pPtRes As PromptPointResult = acDoc.Editor.GetPoint(pPtOpts)
                                    Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                                    Dim ts As Autodesk.AutoCAD.DatabaseServices.TableStyle = tr.GetObject(IDTablelStyleDraw, OpenMode.ForRead)
                                    Dim template As Autodesk.AutoCAD.DatabaseServices.TableTemplate = tr.GetObject(ts.Template, OpenMode.ForRead)
                                    tb.CopyFrom(template, TableCopyOptions.FillTarget)
                                    tb.TableStyle = IDTablelStyleDraw
                                    tb.SetDatabaseDefaults()
                                    ms.UpgradeOpen()
                                    ms.AppendEntity(tb)
                                    tb.Position = pPtRes.Value
                                    tr.AddNewlyCreatedDBObject(tb, True)
                                    Dim HRowsTbl As Double = 5
                                    Dim HTextTbl As Double = 2
                                    Dim ColorTextTbl As Integer = 256
                                    If IsArray(ArrayPropertiesTableRow) = True Then
                                        For i As Integer = 0 To ArrayPropertiesTableRow.GetUpperBound(1)
                                            Dim NameField As String = ArrayPropertiesTableRow(0, i)
                                            If NameField.Trim Like "Наименование слоя" Then
                                                Dim NameLayer As String = ArrayPropertiesTableRow(1, i)
                                                If IsNothing(NameLayer) = False Then
                                                    If NameLayer.Trim.Length > 0 Then
                                                        Dim acLyrTbl As LayerTable = tr.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                                                        If acLyrTbl.Has(NameLayer) = True Then
                                                            tb.Layer = NameLayer
                                                        Else
                                                            Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplateDWG, NameLayer)
                                                            If acLyrTbl.Has(NameLayer) = True Then
                                                                tb.Layer = NameLayer
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Цвет" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    Dim ColorBlk As Integer = Val(ArrayPropertiesTableRow(1, i))
                                                    If ColorBlk < 0 Then
                                                        ColorBlk = 0
                                                    ElseIf ColorBlk > 256 Then
                                                        ColorBlk = 256
                                                    End If
                                                    tb.ColorIndex = ColorBlk
                                                End If
                                            ElseIf NameField.Trim Like "Тип линии" Then
                                                Dim NameTypeLine As String = ArrayPropertiesTableRow(1, i)
                                                If IsNothing(NameTypeLine) = False Then
                                                    If NameTypeLine.Trim.Length > 0 Then
                                                        If Not (NameTypeLine Like "ПоСлою" Or NameTypeLine Like "ПоБлоку") Then
                                                            Dim acLineTypTbl As LinetypeTable = tr.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                                                            If acLineTypTbl.Has(NameTypeLine) = True Then
                                                                tb.Linetype = NameTypeLine
                                                            Else
                                                                Dim boolTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PatchTemplateDWG, NameTypeLine)
                                                                If acLineTypTbl.Has(NameTypeLine) = True Then
                                                                    tb.Linetype = NameTypeLine
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Вес линии" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    Dim WidthLin As Integer = Val(ArrayPropertiesTableRow(1, i))
                                                    If WidthLin > 0 Then
                                                        tb.LineWeight = WidthLin
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Высота текста" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    HTextTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HTextTbl <= 0 Then
                                                        HTextTbl = 2
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Цвет текста" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    ColorTextTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HTextTbl <= 0 Then
                                                        ColorTextTbl = 256
                                                    End If
                                                End If
                                            ElseIf NameField.Trim Like "Высота строки" Then
                                                If IsNothing(ArrayPropertiesTableRow(1, i)) = False Then
                                                    HRowsTbl = Val(ArrayPropertiesTableRow(1, i))
                                                    If HRowsTbl <= 0 Then
                                                        HRowsTbl = 5
                                                    End If
                                                End If
                                            End If
                                        Next i
                                    End If
                                    'смотрим количество строк
                                    Dim RowsTbl As Integer = tb.Rows.Count
                                    Dim NumRows As Integer = RowsTbl - 1
                                    tb.Rows(NumRows).Height = HRowsTbl
                                    tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                    tb.Rows(NumRows).TextHeight = HTextTbl
                                    tb.Cells(NumRows, 0).TextString = "НТ"
                                    Dim PKStartAcad As Double = PKStartAlign 'пикет начала трассы
                                    'преобразовываем его в форматированный пикет
                                    Dim PiketStartStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKStartAlign, RoundLineZn, indFormatPK)
                                    'вычленяем километр
                                    Dim KMInt As Integer = Val(PiketStartStrFormat)
                                    Dim KM1 As Integer = 0
                                    'записываем километр
                                    If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                        tb.Cells(NumRows, 1).TextString = KM1
                                    Else
                                        KMInt = Val(PiketStartStr)
                                        KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                        tb.Cells(NumRows, 1).TextString = KM1
                                    End If
                                    'записываем пикет
                                    Dim StrPKArray As String() = PiketStartStrFormat.Replace("+", " ").Split
                                    If IsArray(StrPKArray) = True Then
                                        If StrPKArray.Count > 1 Then
                                            tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                            tb.Cells(NumRows, 3).TextString = StrPKArray(1)
                                        Else
                                            tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                        End If
                                    End If
                                    'записываем угол начальный
                                    Dim strAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    tb.Cells(NumRows, 4).TextString = strAngle
                                    tb.Cells(NumRows, 5).TextString = "-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    tb.Cells(NumRows, 24).TextString = FuncFormatZN.FuncFormatCoordinate(AlignStartY, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    tb.Cells(NumRows, 25).TextString = FuncFormatZN.FuncFormatCoordinate(AlignStartX, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    '===================================================================================================================
                                    'запоминаем позицию для вставки значения "длина прямой вставки"
                                    Dim PosStartLg As Integer = NumRows + 1 'номер позиции для начального отрезка
                                    NumRows += 2
                                    '======================================================================================================================
                                    'промежуточные значения
                                    If IsArray(ArrayExcel) = True Then
                                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                            tb.InsertRows(NumRows, HRowsTbl, 2)
                                            tb.Rows(NumRows).Height = HRowsTbl
                                            tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                            tb.Rows(NumRows).TextHeight = HTextTbl
                                            tb.Rows(NumRows - 1).Height = HRowsTbl
                                            tb.Rows(NumRows - 1).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                            tb.Rows(NumRows - 1).TextHeight = HTextTbl
                                            tb.Cells(NumRows, 0).TextString = "ВУ " & ArrayExcel(0, i) 'вершина угла (номер)
                                            'находим пикет
                                            Dim PiketStr As String = ArrayExcel(1, i)
                                            If FormRoadPlane.ComboBox2.SelectedIndex = 0 Then
                                                PiketStr = FuncFormatZN.FuncFormatPK(Val(PiketStr), RoundLineZn)
                                                'вычленяем километр
                                                KMInt = Val(PiketStr)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10, 0)
                                                tb.Cells(NumRows, 1).TextString = KM1
                                            Else
                                                KMInt = Val(PiketStr)
                                                KM1 = MathFunction.FuncTrimDigitFloo((KMInt) / 10000, 0)
                                                tb.Cells(NumRows, 1).TextString = KM1
                                            End If
                                            'записываем пикет
                                            Erase StrPKArray
                                            StrPKArray = PiketStr.Replace("+", " ").Split
                                            If IsArray(StrPKArray) = True Then
                                                If StrPKArray.Count > 1 Then
                                                    tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                                    tb.Cells(NumRows, 3).TextString = StrPKArray(1)
                                                Else
                                                    tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                                End If
                                            End If
                                            'угол дельта
                                            Dim Ugol As Double = ArrayExcel(2, i)
                                            Dim AbsUgol As Double = Math.Abs(Ugol)
                                            Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                            If Ugol <= 0 Then
                                                tb.Cells(NumRows, 4).TextString = DAngle
                                                tb.Cells(NumRows, 5).TextString = "-"
                                            Else
                                                tb.Cells(NumRows, 5).TextString = DAngle
                                                tb.Cells(NumRows, 4).TextString = "-"
                                            End If
                                            'записываем радиус
                                            Dim rad1 As Double = ArrayExcel(3, i)
                                            Dim rad2 As Double = ArrayExcel(21, i)
                                            Dim rad3 As Double = ArrayExcel(22, i)
                                            Dim StrRAd As String = "" 'результирующая строка
                                            If rad1 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad1 - Int(rad1))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = Int(rad1) & vbLf 'радиус
                                                Else
                                                    StrRAd = FuncFormatZN.FuncFormatCoordinate(rad1, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad2 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad2 - Int(rad2))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad2) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad2, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            If rad3 > 0 Then
                                                Dim flrad As Double = Math.Abs(rad3 - Int(rad3))
                                                If Math.Abs(flrad) < 0.01 Then
                                                    StrRAd = StrRAd & Int(rad3) & vbLf 'радиус
                                                Else
                                                    StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad3, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                                End If
                                            End If
                                            'записываем радиус
                                            StrRAd = StrRAd.Trim
                                            If StrRAd.Length > 0 Then
                                                tb.Cells(NumRows, 6).TextString = StrRAd
                                            Else
                                                tb.Cells(NumRows, 6).TextString = "0"
                                            End If
                                            '===================================================================================================================
                                            'записываем элементы кривой
                                            'ТАНГЕНСЫ
                                            Dim T1 As Double = ArrayExcel(6, i) 'T1
                                            Dim flnumb As Double = Math.Abs(T1 - Int(T1))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 7).TextString = Int(T1) 'T1
                                            Else
                                                tb.Cells(NumRows, 7).TextString = FuncFormatZN.FuncFormatCoordinate(T1, RoundLineZn) 'записываем T1
                                            End If
                                            Dim T2 As Double = ArrayExcel(7, i) 'T2
                                            flnumb = Math.Abs(T2 - Int(T2))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 8).TextString = Int(T2) 'T2
                                            Else
                                                tb.Cells(NumRows, 8).TextString = FuncFormatZN.FuncFormatCoordinate(T2, RoundLineZn) 'записываем T2
                                            End If
                                            'ПЕРЕХОДНЫЕ КРИВЫЕ
                                            Dim L1 As Double = ArrayExcel(4, i)
                                            flnumb = Math.Abs(L1 - Int(L1))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 9).TextString = Int(L1) 'L1
                                            Else
                                                tb.Cells(NumRows, 9).TextString = FuncFormatZN.FuncFormatCoordinate(L1, RoundLineZn) 'записываем L1
                                            End If
                                            Dim L2 As Double = ArrayExcel(5, i) 'L2
                                            flnumb = Math.Abs(L2 - Int(L2))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 10).TextString = Int(L2) 'L2
                                            Else
                                                tb.Cells(NumRows, 10).TextString = FuncFormatZN.FuncFormatCoordinate(L2, RoundLineZn) 'записываем L2
                                            End If
                                            'КРИВЫЕ
                                            Dim K As Double = ArrayExcel(9, i)
                                            flnumb = Math.Abs(K - Int(K))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 11).TextString = Int(K) 'K
                                            Else
                                                tb.Cells(NumRows, 11).TextString = FuncFormatZN.FuncFormatCoordinate(K, RoundLineZn) 'записываем K
                                            End If
                                            'БИССЕКТРИСА
                                            Dim B As Double = ArrayExcel(10, i)
                                            flnumb = Math.Abs(B - Int(B))
                                            If flnumb < 0.01 Then
                                                tb.Cells(NumRows, 12).TextString = Int(B) 'B
                                            Else
                                                tb.Cells(NumRows, 12).TextString = FuncFormatZN.FuncFormatCoordinate(B, RoundLineZn) 'записываем K
                                            End If
                                            'записываем пикетаж
                                            '=================================================================================================================
                                            'начало переходной кривой
                                            Dim NPK As Double = ArrayExcel(12, i) 'RAW пикет начала переходной кривой
                                            Dim NKK As Double = ArrayExcel(13, i) 'RAW пикет начала круговой кривой
                                            Dim KKK As Double = ArrayExcel(14, i) 'RAW пикет конца круговой кривой
                                            Dim KPK As Double = ArrayExcel(15, i) 'RAW пикет конца круговой кривой
                                            If L1 > 0 Then
                                                Try
                                                    If NPK <> -10000 Then
                                                        Dim PiketNPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NPK, RoundLineZn, indFormatPK)
                                                        If PiketNPKStrFormat Like "" Then
                                                            tb.Cells(NumRows, 13).TextString = ArrayExcel(12, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketNPKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    tb.Cells(NumRows, 13).TextString = StrPKArray(0)
                                                                    tb.Cells(NumRows, 14).TextString = StrPKArray(1)
                                                                Else
                                                                    tb.Cells(NumRows, 13).TextString = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        tb.Cells(NumRows, 13).TextString = "-"
                                                        tb.Cells(NumRows, 14).TextString = "-"
                                                    End If
                                                Catch
                                                End Try
                                                'начало круговой кривой
                                                Try
                                                    If NKK <> -10000 Then
                                                        Dim PiketNKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NKK, RoundLineZn, indFormatPK)
                                                        If PiketNKKStrFormat Like "" Then
                                                            tb.Cells(NumRows, 15).TextString = ArrayExcel(13, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketNKKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    tb.Cells(NumRows, 15).TextString = StrPKArray(0)
                                                                    tb.Cells(NumRows, 16).TextString = StrPKArray(1)
                                                                Else
                                                                    tb.Cells(NumRows, 15).TextString = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        tb.Cells(NumRows, 15).TextString = "-"
                                                        tb.Cells(NumRows, 16).TextString = "-"
                                                    End If
                                                Catch
                                                End Try
                                            Else
                                                tb.Cells(NumRows, 13).TextString = "-"
                                                tb.Cells(NumRows, 14).TextString = "-"
                                                tb.Cells(NumRows, 15).TextString = "-"
                                                tb.Cells(NumRows, 16).TextString = "-"
                                            End If
                                            'конец круговой кривой
                                            If L2 > 0 Then
                                                Try
                                                    If KKK <> -10000 Then
                                                        Dim PiketKKKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KKK, RoundLineZn, indFormatPK)
                                                        If PiketKKKStrFormat Like "" Then
                                                            tb.Cells(NumRows, 17).TextString = ArrayExcel(14, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketKKKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    tb.Cells(NumRows, 17).TextString = StrPKArray(0)
                                                                    tb.Cells(NumRows, 18).TextString = StrPKArray(1)
                                                                Else
                                                                    tb.Cells(NumRows, 17).TextString = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        tb.Cells(NumRows, 17).TextString = "-"
                                                        tb.Cells(NumRows, 18).TextString = "-"
                                                    End If
                                                Catch
                                                End Try
                                                'конец переходной кривой
                                                Try
                                                    If KPK <> -10000 Then
                                                        Dim PiketKPKStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, KPK, RoundLineZn, indFormatPK)
                                                        If PiketKPKStrFormat Like "" Then
                                                            tb.Cells(NumRows, 19).TextString = ArrayExcel(15, i)
                                                        Else
                                                            Erase StrPKArray
                                                            StrPKArray = PiketKPKStrFormat.Replace("+", " ").Split
                                                            If IsArray(StrPKArray) = True Then
                                                                If StrPKArray.Count > 1 Then
                                                                    tb.Cells(NumRows, 19).TextString = StrPKArray(0)
                                                                    tb.Cells(NumRows, 20).TextString = StrPKArray(1)
                                                                Else
                                                                    tb.Cells(NumRows, 19).TextString = StrPKArray(0)
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        tb.Cells(NumRows, 19).TextString = "-"
                                                        tb.Cells(NumRows, 20).TextString = "-"
                                                    End If
                                                Catch
                                                End Try
                                            Else
                                                tb.Cells(NumRows, 17).TextString = "-"
                                                tb.Cells(NumRows, 18).TextString = "-"
                                                tb.Cells(NumRows, 19).TextString = "-"
                                                tb.Cells(NumRows, 20).TextString = "-"
                                            End If
                                            '=================================================================================================
                                            'записываем длины прямых участков
                                            If i <> 0 Then
                                                tb.Cells(NumRows - 1, 22).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(17, i - 1), RoundLineZn) 'Длина прямого участка
                                            End If
                                            'КООРДИНАТЫ ВЕРШИН УГЛОВ
                                            tb.Cells(NumRows, 24).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(20, i), FormRoadPlane.NumericUpDown6.Value) 'Y
                                            tb.Cells(NumRows, 25).TextString = FuncFormatZN.FuncFormatCoordinate(ArrayExcel(19, i), FormRoadPlane.NumericUpDown6.Value) 'X
                                            'рассчитываем рассояния между вершинами
                                            Dim vect As Vector2d = Nothing
                                            Dim XY1 As Point2d = Nothing
                                            Dim XY2 As Point2d = Nothing
                                            If i = 0 Then
                                                XY1 = New Point2d(AlignStartX, AlignStartY)
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            Else
                                                XY1 = New Point2d(ArrayExcel(19, i - 1), ArrayExcel(20, i - 1))
                                                XY2 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                            End If
                                            vect = XY1.GetVectorTo(XY2)
                                            tb.Cells(NumRows - 1, 21).TextString = FuncFormatZN.FuncFormatCoordinate(vect.Length, RoundLineZn) 'Расстояние между вершинами
                                            'записываем румб
                                            Dim DirectMath As Double = vect.Angle 'угол в радианах (румб)
                                            Dim DirectGeodesic As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath)
                                            Dim Grad As Double = (DirectGeodesic * 180) / Math.PI 'переводим в градусы
                                            tb.Cells(NumRows - 1, 23).TextString = FuncSurvey.FuncDirectToRumbDDMMSS(Grad, FormRoadPlane.NumericUpDown1.Value)
                                            NumRows += 2
                                        Next i
                                    End If
                                    '=================================================================================================================================
                                    'записываем конец трассы
                                    tb.InsertRows(NumRows, HRowsTbl, 2)
                                    tb.Rows(NumRows).Height = HRowsTbl
                                    tb.Rows(NumRows).ContentColor = Color.FromColorIndex(ColorMethod.ByAci, ColorTextTbl)
                                    tb.Rows(NumRows).TextHeight = HTextTbl
                                    tb.Cells(NumRows, 0).TextString = "КТ"
                                    Dim PKEndAcad As Double = PKEndAlign 'пикет начала трассы
                                    'преобразовываем его в форматированный пикет
                                    Dim PiketEndStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, PKEndAlign, RoundLineZn, indFormatPK)
                                    'вычленяем километр
                                    KMInt = Val(PiketEndStrFormat)
                                    KM1 = Math.Floor(KMInt) / 10
                                    tb.Cells(NumRows, 1).TextString = KM1  'значение километра
                                    'записываем пикет
                                    Erase StrPKArray
                                    StrPKArray = PiketEndStrFormat.Replace("+", " ").Split
                                    If IsArray(StrPKArray) = True Then
                                        If StrPKArray.Count > 1 Then
                                            tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                            tb.Cells(NumRows, 3).TextString = StrPKArray(1)
                                        Else
                                            tb.Cells(NumRows, 2).TextString = StrPKArray(0)
                                        End If
                                    End If
                                    'записываем угол начальный
                                    strAngle = FuncFormatZN.FuncConvertDegToDDMMSS(0, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    tb.Cells(NumRows, 4).TextString = strAngle
                                    tb.Cells(NumRows, 5).TextString = "-"
                                    '===================================================================================================================
                                    'записываем X и Y
                                    'записываем координаты вершины угла
                                    tb.Cells(NumRows, 24).TextString = FuncFormatZN.FuncFormatCoordinate(AlignEndY, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    tb.Cells(NumRows, 25).TextString = FuncFormatZN.FuncFormatCoordinate(AlignEndX, FormRoadPlane.NumericUpDown6.Value) 'записываем координаты вершины
                                    'рассчитываем расстояния между вершинами
                                    Dim XYk As Point2d = Nothing
                                    If IsArray(ArrayExcel) = True Then
                                        Dim PosC As Integer = ArrayExcel.GetUpperBound(1)
                                        XYk = New Point2d(ArrayExcel(19, PosC), ArrayExcel(20, PosC))
                                    Else
                                        XYk = New Point2d(AlignStartX, AlignStartY)
                                        tb.Cells(NumRows - 1, 6).TextString = "-" 'радиус
                                        tb.Cells(NumRows - 1, 7).TextString = "-" 'T1
                                        tb.Cells(NumRows - 1, 8).TextString = "-" 'T2
                                        tb.Cells(NumRows - 1, 9).TextString = "-" 'L1
                                        tb.Cells(NumRows - 1, 10).TextString = "-" 'L2
                                        tb.Cells(NumRows - 1, 11).TextString = "-" 'K
                                        tb.Cells(NumRows - 1, 12).TextString = "-" 'Б
                                        tb.Cells(NumRows - 1, 13).TextString = "-" 'НПК
                                        tb.Cells(NumRows - 1, 14).TextString = "-" 'НПК
                                        tb.Cells(NumRows - 1, 15).TextString = "-" 'НКК
                                        tb.Cells(NumRows - 1, 16).TextString = "-" 'НКК
                                        tb.Cells(NumRows - 1, 17).TextString = "-" 'ККК
                                        tb.Cells(NumRows - 1, 18).TextString = "-" 'ККК
                                        tb.Cells(NumRows - 1, 19).TextString = "-" 'КПК
                                        tb.Cells(NumRows - 1, 20).TextString = "-" 'КПК
                                    End If
                                    Dim XYk2 As Point2d = New Point2d(AlignEndX, AlignEndY)
                                    Dim vect1 As Vector2d = XYk.GetVectorTo(XYk2)
                                    tb.Cells(NumRows - 1, 21).TextString = FuncFormatZN.FuncFormatCoordinate(vect1.Length, FormRoadPlane.NumericUpDown6.Value) 'Расстояние между вершинами
                                    'записываем румб
                                    Dim DirectMath2 As Double = vect1.Angle 'угол в радианах (румб)
                                    Dim DirectGeodesic2 As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(DirectMath2)
                                    Dim Grad2 As Double = (DirectGeodesic2 * 180) / Math.PI 'переводим в градусы
                                    tb.Cells(NumRows - 1, 23).TextString = FuncSurvey.FuncDirectToRumbDDMMSS(Grad2, 1)
                                    'записываем длину конечного отрезка
                                    tb.Cells(NumRows - 1, 22).TextString = FuncFormatZN.FuncFormatCoordinate(TLenghtEnd, FormRoadPlane.NumericUpDown6.Value) 'длина конечной прямой
                                    'записываем длину прямой вставки начального отрезка
                                    tb.Cells(PosStartLg, 22).TextString = FuncFormatZN.FuncFormatCoordinate(TLenghtStart, FormRoadPlane.NumericUpDown6.Value) 'длина начальной прямой
                                End If
                            End If
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'отчет схема разбивки кривой в плане
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    ElseIf FormRoadPlane.RadioButton2.Checked = True Then
                        'формируем исходные данные для функции
                        'ОТЧЕТ В MS EXCEL
                        If FormRoadPlane.ComboBox10.SelectedIndex = 0 Then
                            'запускаем Excel
                            Dim xlApp As Excel.Application = Nothing
                            Dim xlBook As Excel.Workbook = Nothing
                            Dim xlSheets As Excel.Worksheet = Nothing
                            Dim cell1 As Excel.Range 'первая ячейка в строке
                            Dim cell2 As Excel.Range 'последняя ячейка в строке
                            Dim CellRng As Excel.Range 'массив ячеек
                            Dim Col As Integer = 23
                            xlApp = CreateObject("Excel.Application")
                            xlBook = xlApp.Workbooks.Add()
                            xlApp.Visible = True
                            'находим нужный лист(Ведомость объемов работ)
                            Dim SheetCountZu1 As Integer = 1
                            Dim xlsheet As Excel.Worksheet
                            xlsheet = xlBook.Worksheets.Item("Лист1") 'делаем ссылку на этот лист
                            xlsheet.Activate()
                            Dim PosInd As Integer = 1
                            Dim ArrayExcelRezult As Double(,) = Nothing 'результирующий массив для координат
                            Dim ArrayK As Double(,) = Nothing '(1 - НПК, 2 - НКК, 3 - ККК, 4 - КПК, 5 - СК, 6 - Рубл.пикет)
                            Dim CountArrayK As Integer = 0
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                Dim K As Double = ArrayExcel(8, i)
                                If K <> 0 Then
                                    'кривые существуют в данной конструкции модно начинать разбивку
                                    'заносим в массив основные параметры кривой
                                    Dim NPK As Double = -1
                                    Dim NPK2 As Double = -1
                                    Dim NKK As Double = -1
                                    Dim KKK As Double = -1
                                    Dim KPK As Double = -1
                                    If ArrayExcel(4, i) > 0 Then 'если существует первая L1 записываем НПК
                                        If ArrayExcel(9, i) = 0 Then
                                            NPK = ArrayExcel(12, i)
                                            KPK = ArrayExcel(15, i)
                                        Else
                                            NPK = ArrayExcel(12, i)
                                        End If
                                    End If
                                    If ArrayExcel(5, i) > 0 Then
                                        If ArrayExcel(9, i) = 0 Then
                                            NPK2 = ArrayExcel(13, i)
                                            KPK = ArrayExcel(15, i)
                                        Else
                                            KPK = ArrayExcel(15, i)
                                        End If
                                    End If
                                    If ArrayExcel(9, i) > 0 Then 'если есть Ksohr
                                        NKK = ArrayExcel(13, i)
                                        KKK = ArrayExcel(14, i)
                                    End If
                                    'записываем полученные результаты в массив
                                    If NPK > -1 Then
                                        ReDim Preserve ArrayK(2, CountArrayK)
                                        ArrayK(0, CountArrayK) = NPK 'НПК
                                        ArrayK(1, CountArrayK) = 1
                                        ArrayK(2, CountArrayK) = 0
                                        CountArrayK += 1
                                    End If
                                    If NPK2 > -1 Then
                                        ReDim Preserve ArrayK(2, CountArrayK)
                                        ArrayK(0, CountArrayK) = NPK2 'НПК
                                        ArrayK(1, CountArrayK) = 1
                                        ArrayK(2, CountArrayK) = 0
                                        CountArrayK += 1
                                    End If
                                    If NKK > -1 Then
                                        ReDim Preserve ArrayK(2, CountArrayK)
                                        ArrayK(0, CountArrayK) = NKK 'НПК
                                        ArrayK(1, CountArrayK) = 2
                                        ArrayK(2, CountArrayK) = 0
                                        CountArrayK += 1
                                    End If
                                    If KKK > -1 Then
                                        ReDim Preserve ArrayK(2, CountArrayK)
                                        ArrayK(0, CountArrayK) = KKK 'НПК
                                        ArrayK(1, CountArrayK) = 3
                                        ArrayK(2, CountArrayK) = 0
                                        CountArrayK += 1
                                    End If
                                    If KPK > -1 Then
                                        ReDim Preserve ArrayK(2, CountArrayK)
                                        ArrayK(0, CountArrayK) = KPK 'НПК
                                        ArrayK(1, CountArrayK) = 4
                                        ArrayK(2, CountArrayK) = 0
                                        CountArrayK += 1
                                    End If
                                    'записываем середину кривой в массив
                                    '====================================================================================================================
                                    'записываем пикет середины кривой
                                    ReDim Preserve ArrayK(2, CountArrayK)
                                    ArrayK(0, CountArrayK) = Math.Round(ArrayExcel(23, i), RoundLineZn) 'СК
                                    ArrayK(1, CountArrayK) = 5
                                    ArrayK(2, CountArrayK) = 0
                                    'находим начало и конец конструкции
                                    '=====================================================================================================================
                                    'определяем стартовый и конечный пикеты
                                    Dim StartPK As Double = ArrayExcel(12, i)
                                    If ArrayExcel(4, i) = 0 Then
                                        If ArrayExcel(9, i) > 0 Then
                                            StartPK = ArrayExcel(13, i)
                                        End If
                                    End If
                                    Dim EndPK As Double = ArrayExcel(15, i)
                                    If ArrayExcel(5, i) = 0 Then
                                        If ArrayExcel(9, i) > 0 Then
                                            EndPK = ArrayExcel(14, i)
                                        End If
                                    End If
                                    'записываем рубленные пикеты
                                    Dim eqPiketColl As StationEquationCollection = Align.StationEquations
                                    If eqPiketColl.Count > 0 Then
                                        For Each PKRubl As StationEquation In eqPiketColl
                                            Dim RAwPKRubl As Double = PKRubl.RawStationBack
                                            RAwPKRubl = Math.Round(RAwPKRubl, RoundLineZn)
                                            If RAwPKRubl >= StartPK And RAwPKRubl <= EndPK Then
                                                ReDim Preserve ArrayK(2, CountArrayK)
                                                ArrayK(0, CountArrayK) = RAwPKRubl 'RAW пикет
                                                ArrayK(1, CountArrayK) = 6 'описание
                                                ArrayK(2, CountArrayK) = 0 'флаг что элемент нуждается в записи
                                                CountArrayK += 1
                                            End If
                                        Next
                                    End If
                                    '=================================================================================================
                                    'сортируем массив
                                    Dim boolArrayK As Boolean = MathFunction.FuncSortDblArray(ArrayK, 0)
                                    '=====================================================================================================================
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'запускаем функцию по определению координат
                                    Dim ShagR As Integer = FormRoadPlane.NumericUpDown2.Value
                                    Dim VUCoord As Point2d = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                    Dim boolCoordPlane As Boolean = FuncAlignment.FuncSplitAlignCurvePlane(Align, ArrayExcelRezult, StartPK, EndPK, ArrayK, ShagR, VUCoord, ArrayExcel(23, i), RoundLineZn)
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                '==========================================================================================================================
                                'запись значений в Excel
                                If IsArray(ArrayExcelRezult) = True Then
                                    '============================================================================
                                    'формируем шапку таблицы
                                    cell1 = xlsheet.Cells(PosInd, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 8, 7) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.HorizontalAlignment = Excel.Constants.xlCenter
                                    CellRng.VerticalAlignment = Excel.Constants.xlCenter
                                    With CellRng.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With
                                    With CellRng.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With
                                    With CellRng.Borders(Excel.XlBordersIndex.xlEdgeRight)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With
                                    With CellRng.Borders(Excel.XlBordersIndex.xlEdgeTop)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With
                                    With CellRng.Borders(Excel.XlBordersIndex.xlInsideVertical)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With
                                    With CellRng.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                                        .LineStyle = Excel.XlLineStyle.xlContinuous
                                        .ColorIndex = 0
                                        .TintAndShade = 0
                                        .Weight = 2
                                    End With

                                    '===========================================================================
                                    'форматируем значения (число знаков после запятой
                                    Dim NumbFormat As String = "0.00"
                                    If RoundLineZn = 0 Then
                                        NumbFormat = "0"
                                    ElseIf RoundLineZn = 1 Then
                                        NumbFormat = "0.0"
                                    ElseIf RoundLineZn = 2 Then
                                        NumbFormat = "0.00"
                                    ElseIf RoundLineZn = 3 Then
                                        NumbFormat = "0.000"
                                    ElseIf RoundLineZn = 4 Then
                                        NumbFormat = "0.00000"
                                    End If
                                    cell1 = xlsheet.Cells(PosInd + 1, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 1, 7) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.NumberFormat = NumbFormat
                                    cell1 = xlsheet.Cells(PosInd + 3, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 3, 7) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.NumberFormat = NumbFormat
                                    xlsheet.Cells(PosInd, 1) = "Вершина"
                                    xlsheet.Cells(PosInd + 1, 1) = "ВУ" & ArrayExcel(0, i)
                                    xlsheet.Cells(PosInd, 2) = "X"
                                    xlsheet.Cells(PosInd + 1, 2) = ArrayExcel(20, i)
                                    xlsheet.Cells(PosInd, 3) = "Y"
                                    xlsheet.Cells(PosInd + 1, 3) = ArrayExcel(19, i)
                                    xlsheet.Cells(PosInd, 4) = "R"
                                    Dim rad1 As Double = ArrayExcel(3, i)
                                    Dim rad2 As Double = ArrayExcel(21, i)
                                    Dim rad3 As Double = ArrayExcel(22, i)
                                    Dim StrRAd As String = "" 'результирующая строка
                                    If rad1 > 0 Then
                                        Dim flrad As Double = Math.Abs(rad1 - Int(rad1))
                                        If Math.Abs(flrad) < 0.01 Then
                                            StrRAd = Int(rad1) & vbLf 'радиус
                                        Else
                                            StrRAd = FuncFormatZN.FuncFormatCoordinate(rad1, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                        End If
                                    End If
                                    If rad2 > 0 Then
                                        Dim flrad As Double = Math.Abs(rad2 - Int(rad2))
                                        If Math.Abs(flrad) < 0.01 Then
                                            StrRAd = StrRAd & Int(rad2) & vbLf 'радиус
                                        Else
                                            StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad2, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                        End If
                                    End If
                                    If rad3 > 0 Then
                                        Dim flrad As Double = Math.Abs(rad3 - Int(rad3))
                                        If Math.Abs(flrad) < 0.01 Then
                                            StrRAd = StrRAd & Int(rad3) & vbLf 'радиус
                                        Else
                                            StrRAd = StrRAd & FuncFormatZN.FuncFormatCoordinate(rad3, FormRoadPlane.NumericUpDown6.Value) & vbLf 'записываем радиус
                                        End If
                                    End If
                                    xlsheet.Cells(PosInd + 1, 4) = StrRAd.Trim
                                    xlsheet.Cells(PosInd, 5) = "L1"
                                    xlsheet.Cells(PosInd + 1, 5) = ArrayExcel(4, i)
                                    xlsheet.Cells(PosInd, 6) = "L2"
                                    xlsheet.Cells(PosInd + 1, 6) = ArrayExcel(5, i)
                                    xlsheet.Cells(PosInd + 2, 1) = "T1"
                                    xlsheet.Cells(PosInd + 3, 1) = ArrayExcel(6, i)
                                    xlsheet.Cells(PosInd + 2, 2) = "T2"
                                    xlsheet.Cells(PosInd + 3, 2) = ArrayExcel(7, i)
                                    xlsheet.Cells(PosInd + 2, 3) = "K"
                                    xlsheet.Cells(PosInd + 3, 3) = ArrayExcel(8, i)
                                    xlsheet.Cells(PosInd + 2, 4) = "Kсохр"
                                    xlsheet.Cells(PosInd + 3, 4) = ArrayExcel(9, i)
                                    xlsheet.Cells(PosInd + 2, 5) = "Б"
                                    xlsheet.Cells(PosInd + 3, 5) = ArrayExcel(10, i)
                                    xlsheet.Cells(PosInd + 2, 6) = "Д"
                                    xlsheet.Cells(PosInd + 3, 6) = ArrayExcel(11, i)
                                    xlsheet.Cells(PosInd + 4, 1) = "ПК ВУ"
                                    xlsheet.Cells(PosInd + 5, 1) = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(1, i), RoundLineZn, indFormatPK)
                                    xlsheet.Cells(PosInd + 4, 2) = "Угол"
                                    Dim Ugol As Double = ArrayExcel(2, i)
                                    Dim AbsUgol As Double = Math.Abs(Ugol)
                                    Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, FormRoadPlane.NumericUpDown1.Value)
                                    xlsheet.Cells(PosInd + 5, 2) = DAngle
                                    xlsheet.Cells(PosInd + 4, 3) = "НПК"
                                    xlsheet.Cells(PosInd + 5, 3) = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(12, i), RoundLineZn, indFormatPK)
                                    xlsheet.Cells(PosInd + 4, 4) = "НКК"
                                    xlsheet.Cells(PosInd + 5, 4) = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(13, i), RoundLineZn, indFormatPK)
                                    xlsheet.Cells(PosInd + 4, 5) = "ККК"
                                    xlsheet.Cells(PosInd + 5, 5) = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(14, i), RoundLineZn, indFormatPK)
                                    xlsheet.Cells(PosInd + 4, 6) = "КПК"
                                    xlsheet.Cells(PosInd + 5, 6) = FuncAlignment.FuncFormatRublPiket(Align, ArrayExcel(15, i), RoundLineZn, indFormatPK)
                                    cell1 = xlsheet.Cells(PosInd + 6, 2) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 6, 4) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    xlsheet.Cells(PosInd + 6, 2) = "Разбивочные данные (м.)"
                                    cell1 = xlsheet.Cells(PosInd + 6, 5) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 6, 6) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    xlsheet.Cells(PosInd + 6, 5) = "Координаты (м.)"
                                    xlsheet.Cells(PosInd + 7, 2) = "dS"
                                    xlsheet.Cells(PosInd + 7, 3) = "dX"
                                    xlsheet.Cells(PosInd + 7, 4) = "dY"
                                    xlsheet.Cells(PosInd + 7, 5) = "X"
                                    xlsheet.Cells(PosInd + 7, 6) = "Y"
                                    'объединяем ячейки для поля ПК
                                    cell1 = xlsheet.Cells(PosInd + 6, 1) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 7, 1) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    xlsheet.Cells(PosInd + 6, 1) = "ПК+"
                                    'объединяем ячейки для поля Примечание
                                    cell1 = xlsheet.Cells(PosInd + 6, 7) 'ссылка на первую ячейку
                                    cell2 = xlsheet.Cells(PosInd + 7, 7) 'ссылка на 6 ячейку в строке
                                    CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Merge(Type.Missing)
                                    xlsheet.Cells(PosInd + 6, 7) = "Примечание"
                                    PosInd = PosInd + 8
                                    For j2 As Integer = 0 To ArrayExcelRezult.GetUpperBound(1)
                                        cell1 = xlsheet.Cells(PosInd, 1) 'ссылка на первую ячейку
                                        cell2 = xlsheet.Cells(PosInd, 7) 'ссылка на 6 ячейку в строке
                                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                        With CellRng.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                                            .LineStyle = Excel.XlLineStyle.xlContinuous
                                            .ColorIndex = 0
                                            .TintAndShade = 0
                                            .Weight = 2
                                        End With
                                        With CellRng.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                                            .LineStyle = Excel.XlLineStyle.xlContinuous
                                            .ColorIndex = 0
                                            .TintAndShade = 0
                                            .Weight = 2
                                        End With
                                        With CellRng.Borders(Excel.XlBordersIndex.xlEdgeRight)
                                            .LineStyle = Excel.XlLineStyle.xlContinuous
                                            .ColorIndex = 0
                                            .TintAndShade = 0
                                            .Weight = 2
                                        End With
                                        With CellRng.Borders(Excel.XlBordersIndex.xlEdgeTop)
                                            .LineStyle = Excel.XlLineStyle.xlContinuous
                                            .ColorIndex = 0
                                            .TintAndShade = 0
                                            .Weight = 2
                                        End With
                                        With CellRng.Borders(Excel.XlBordersIndex.xlInsideVertical)
                                            .LineStyle = Excel.XlLineStyle.xlContinuous
                                            .ColorIndex = 0
                                            .TintAndShade = 0
                                            .Weight = 2
                                        End With
                                        CellRng.HorizontalAlignment = Excel.Constants.xlCenter
                                        CellRng.VerticalAlignment = Excel.Constants.xlCenter
                                        cell1 = xlsheet.Cells(PosInd, 2) 'ссылка на первую ячейку
                                        cell2 = xlsheet.Cells(PosInd, 6) 'ссылка на 6 ячейку в строке
                                        CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                                        CellRng.NumberFormat = NumbFormat
                                        Dim NPiket As String = ArrayExcelRezult(0, j2)
                                        'форматируем с учетом рубленности
                                        Dim PiketStartStrFormat As String = FuncAlignment.FuncFormatRublPiket(Align, NPiket, RoundLineZn, indFormatPK)
                                        Dim ds As Double = Val(ArrayExcelRezult(1, j2))
                                        Dim dx As Double = Val(ArrayExcelRezult(2, j2))
                                        Dim dy As Double = Val(ArrayExcelRezult(3, j2))
                                        Dim X As Double = Val(ArrayExcelRezult(4, j2))
                                        Dim Y As Double = Val(ArrayExcelRezult(5, j2))
                                        Dim Note As Double = ArrayExcelRezult(6, j2)
                                        Dim NoteStr As String = "-"
                                        If Note = 0 Then
                                            NoteStr = "-"
                                        ElseIf Note = 1 Then
                                            NoteStr = "НПК"
                                        ElseIf Note = 2 Then
                                            NoteStr = "НКК"
                                        ElseIf Note = 3 Then
                                            NoteStr = "ККК"
                                        ElseIf Note = 4 Then
                                            NoteStr = "КПК"
                                        ElseIf Note = 5 Then
                                            NoteStr = "СК"
                                        ElseIf Note = 6 Then
                                            NoteStr = "Руб.ПК"
                                        End If
                                        xlsheet.Cells(PosInd, 1) = PiketStartStrFormat
                                        xlsheet.Cells(PosInd, 2) = ds
                                        xlsheet.Cells(PosInd, 3) = dy
                                        xlsheet.Cells(PosInd, 4) = dx
                                        xlsheet.Cells(PosInd, 5) = Y
                                        xlsheet.Cells(PosInd, 6) = X
                                        xlsheet.Cells(PosInd, 7) = NoteStr
                                        PosInd += 1
                                        '=========================================================================================
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'РИСУЕМ ГРАФИКУ
                                        'If ArrayExcelRezult(7, j2) <> -1 And ArrayExcelRezult(8, j2) <> -1 Then
                                        '    Dim Pt1 As Point2d = New Point2d(ArrayExcelRezult(4, j2), ArrayExcelRezult(5, j2))
                                        '    Dim Pt2 As Point2d = New Point2d(ArrayExcelRezult(7, j2), ArrayExcelRezult(8, j2))
                                        '    'вставляем второй тангенс
                                        '    Dim acPolyT2 As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                                        '    acPolyT2.SetDatabaseDefaults()
                                        '    acPolyT2.AddVertexAt(0, Pt1, 0, 0, 0)
                                        '    acPolyT2.AddVertexAt(1, Pt2, 0, 0, 0)
                                        '    acPolyT2.Layer = NameLayerVS
                                        '    acBlkTblRec.AppendEntity(acPolyT2)
                                        '    tr.AddNewlyCreatedDBObject(acPolyT2, True)
                                        'End If
                                    Next j2
                                End If
                            Next i
                            xlBook.Close()
                            xlBook = Nothing
                            xlApp = Nothing
                        ElseIf FormRoadPlane.ComboBox10.SelectedIndex = 1 Then
                            'обработка таблицы автокад, НАДО ДОБАВИТЬ ВОЗМОЖНОСТЬ
                        End If
                    End If
                    'возвращаем системную переменную по точности пикетов в предыдущее состояние
                    If boolSist = True Then
                        CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value = SistPerem
                    End If
                    '===============================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выводим графику
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'массив для атрибутов номеров углов поворота, центров кривых
                    If FormRoadPlane.CheckBox1.Checked = True Then
                        'смещение мультивыноски
                        Dim OfsetLeaderX As Double = 4
                        Dim OfsetLeaderY As Double = 4
                        'смещение текстов
                        Dim OffsetText As Double = 1
                        'вставляем точку начала трассы
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'определяем точку начало трассы
                        Dim LayerBlkVUP As String = "0"
                        Dim LayerBlkColor As Integer = 256
                        Dim LayerBlk As Integer = 256
                        Dim pt1 As Point2d = New Point2d(AlignStartX, AlignStartY) 'координаты точки вставки блока
                        Dim PTXYZ1 As Point3d = New Point3d(AlignStartX, AlignStartY, 0)
                        Dim acBlockTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        If acBlockTbl.Has(NameBlkVUP) = True Then
                            'вставляем блок
                            Dim bol As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(PTXYZ1, ArrayPropBlkUPK, 0)
                        End If
                        '===============================================================================================
                        'вставляем мультивыноску с началом трассы
                        Dim PiketStartStr As String = PKStartAlign
                        Try
                            PiketStartStr = Align.GetStationStringWithEquations(PKStartAlign) 'преобразовываем RawPiket
                        Catch ex As System.ArgumentException
                        End Try
                        Dim strNT As String = "Начало трассы" & vbLf & PiketStartStr
                        Dim IDInsMleader1 As ObjectId = AcadFunction.FuncDrawMleader(pt1, strNT, OfsetLeaderX, OfsetLeaderY, LayerMLeaderCurve, 0, Nothing, Nothing)
                        '===================================================================================================
                        'вставляем мультивыноску с координатами
                        Dim CoordX As String = FuncFormatZN.FuncFormatCoordinate(pt1.X, RoundLineZn)
                        Dim CoordY As String = FuncFormatZN.FuncFormatCoordinate(pt1.Y, RoundLineZn)
                        Dim strCoord As String = "X=" & CoordY & vbLf & "Y=" & CoordX
                        Dim IDInsMleader As ObjectId = AcadFunction.FuncDrawMleader(pt1, strCoord, OfsetLeaderX, -4, LayerMLeaderCoord, 0, Nothing, Nothing)
                        '=============================================================================================================================
                        'записываем координаты всех элементов трассы
                        If IsArray(ArrayEntity) = True Then
                            For i As Integer = 1 To ArrayEntity.GetUpperBound(1)
                                Dim ptinsMl As Point2d = New Point2d(ArrayEntity(5, i), ArrayEntity(6, i))
                                CoordX = FuncFormatZN.FuncFormatCoordinate(ArrayEntity(5, i), RoundLineZn)
                                CoordY = FuncFormatZN.FuncFormatCoordinate(ArrayEntity(6, i), RoundLineZn)
                                strCoord = "X=" & CoordY & vbLf & "Y=" & CoordX
                                IDInsMleader = AcadFunction.FuncDrawMleader(ptinsMl, strCoord, OfsetLeaderX, -4, LayerMLeaderCoord, 0, Nothing, Nothing)
                            Next
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'идем циклом по всем вершинам
                        If IsArray(ArrayExcel) = True Then
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                'находим вспомогательные величины (заново)
                                Dim PKStartTan1 As Double = ArrayExcel(12, i) 'NPK
                                If ArrayExcel(4, i) = 0 Then
                                    If ArrayExcel(9, i) > 0 Then
                                        PKStartTan1 = ArrayExcel(13, i) 'NKK
                                    End If
                                End If
                                Dim PKEndTan2 As Double = ArrayExcel(15, i) 'KPK
                                If ArrayExcel(5, i) = 0 Then
                                    If ArrayExcel(9, i) > 0 Then
                                        PKEndTan2 = ArrayExcel(14, i) 'KKK
                                    End If
                                End If
                                'вершина угла поворота
                                pt1 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                Dim AttrPT3d As Point3d = New Point3d(pt1.X, pt1.Y, 0)
                                'вставляем вершину угла
                                '==========================================================================================================================
                                '1. Вставляем блок вершины угла
                                If acBlockTbl.Has(NameBlkVUP) = True Then
                                    'вставляем блок
                                    Dim bol As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(AttrPT3d, ArrayPropBlkUPK, 0)
                                End If
                                '=======================================================================================
                                '2. вставляем мультивыноску с координатами вершины угла
                                CoordX = FuncFormatZN.FuncFormatCoordinate(pt1.X, RoundLineZn)
                                CoordY = FuncFormatZN.FuncFormatCoordinate(pt1.Y, RoundLineZn)
                                strCoord = "X=" & CoordY & vbLf & "Y=" & CoordX
                                IDInsMleader = AcadFunction.FuncDrawMleader(pt1, strCoord, OfsetLeaderX, -4, LayerMLeaderCoord, 0, Nothing, Nothing)
                                '===============================================================================================
                                '3. вставляем характеристики вершины угла
                                Dim NumbVUStr As String = "ВУ " & ArrayExcel(0, i)
                                Dim Ugol As Double = ArrayExcel(2, i)
                                Dim AbsUgol As Double = Math.Abs(Ugol)
                                Dim DAngle As String = FuncFormatZN.FuncConvertDegToDDMMSS(AbsUgol, FormRoadPlane.ComboBox1.SelectedIndex, RoundAngleZn)
                                Dim rad1 As Double = ArrayExcel(3, i)
                                Dim rad2 As Double = ArrayExcel(21, i)
                                Dim rad3 As Double = ArrayExcel(22, i)
                                Dim R As Double = 0
                                If rad1 > 0 Then
                                    R = rad1
                                ElseIf rad2 > 0 Then
                                    R = rad2
                                ElseIf rad3 > 0 Then
                                    R = rad3
                                End If
                                Dim RStr As String = ""
                                If (R - Int(R)) = 0 Then
                                    RStr = MathFunction.FuncTrimDigitFloo(R, 0)
                                Else
                                    RStr = FuncFormatZN.FuncFormatCoordinate(R, RoundLineZn)
                                End If
                                Dim L1 As Double = ArrayExcel(4, i)
                                Dim L2 As Double = ArrayExcel(5, i)
                                Dim L1Str As String = ""
                                Dim L2Str As String = ""
                                If L1 > 0 Then
                                    If (L1 - Int(L1)) = 0 Then
                                        L1Str = MathFunction.FuncTrimDigitFloo(L1, 0)
                                    Else
                                        L1Str = FuncFormatZN.FuncFormatCoordinate(L1, RoundLineZn)
                                    End If
                                End If
                                If L2 > 0 Then
                                    If (L2 - Int(L2)) = 0 Then
                                        L2Str = MathFunction.FuncTrimDigitFloo(L2, 0)
                                    Else
                                        L2Str = FuncFormatZN.FuncFormatCoordinate(L2, RoundLineZn)
                                    End If
                                End If
                                Dim Kpoln As Double = ArrayExcel(8, i)
                                Dim KpolnStr As String = ""
                                If (Kpoln - Int(Kpoln)) = 0 Then
                                    KpolnStr = MathFunction.FuncTrimDigitFloo(Kpoln, 0)
                                Else
                                    KpolnStr = FuncFormatZN.FuncFormatCoordinate(Kpoln, RoundLineZn)
                                End If
                                Dim B As Double = ArrayExcel(10, i)
                                Dim Bstr As String = ""
                                If (B - Int(B)) = 0 Then
                                    Bstr = MathFunction.FuncTrimDigitFloo(B, 0)
                                Else
                                    Bstr = FuncFormatZN.FuncFormatCoordinate(B, RoundLineZn)
                                End If
                                Dim T1 As Double = ArrayExcel(6, i)
                                Dim T2 As Double = ArrayExcel(7, i)
                                Dim T1Str As String = ""
                                Dim T2Str As String = ""
                                If T1 > 0 Then
                                    If (T1 - Int(T1)) = 0 Then
                                        T1Str = MathFunction.FuncTrimDigitFloo(T1, 0)
                                    Else
                                        T1Str = FuncFormatZN.FuncFormatCoordinate(T1, RoundLineZn)
                                    End If
                                End If
                                If T2 > 0 Then
                                    If (T2 - Int(T2)) = 0 Then
                                        T2Str = MathFunction.FuncTrimDigitFloo(T2, 0)
                                    Else
                                        T2Str = FuncFormatZN.FuncFormatCoordinate(T2, RoundLineZn)
                                    End If
                                End If
                                '========================================================================================================
                                'находим пикет вершины угла
                                Dim PKVu As Double = ArrayExcel(1, i) 'пикет вершины угла
                                'преобразовываем число в пикет
                                Dim PiketStr As String = FuncFormatZN.FuncFormatPK(PKVu, RoundLineZn)
                                '=======================================================================================================
                                'выводим домер
                                Dim D As Double = ArrayExcel(11, i)
                                Dim Dstr As String = ""
                                If (D - Int(D)) = 0 Then
                                    Dstr = MathFunction.FuncTrimDigitFloo(D, 0)
                                Else
                                    Dstr = FuncFormatZN.FuncFormatCoordinate(D, RoundLineZn)
                                End If
                                '===========================================================================================================
                                'формируем строку
                                NumbVUStr = NumbVUStr & " \U+0394=" & DAngle & " R=" & RStr & " Б=" & Bstr & " Кполн=" & KpolnStr & vbLf
                                If T1Str Like T2Str Then
                                    NumbVUStr = NumbVUStr & "T1=T2=" & T1Str
                                Else
                                    If T1Str.Length > 0 Then
                                        NumbVUStr = NumbVUStr & " T1=" & T1Str
                                    End If
                                    If T2Str.Length > 0 Then
                                        NumbVUStr = NumbVUStr & " T2=" & T2Str
                                    End If
                                End If
                                If Not (L1Str Like "" And L2Str Like "") Then
                                    If L1Str Like L2Str Then
                                        NumbVUStr = NumbVUStr & " L1=L2=" & L1Str
                                    Else
                                        If L1Str.Length > 0 Then
                                            NumbVUStr = NumbVUStr & " L1=" & L1Str
                                        End If
                                        If L2Str.Length > 0 Then
                                            NumbVUStr = NumbVUStr & " L2=" & L2Str
                                        End If
                                    End If
                                End If
                                'дополняем данные домером и пикетом ВУ
                                NumbVUStr = NumbVUStr & " Д=" & Dstr & " ПК " & PiketStr
                                IDInsMleader = AcadFunction.FuncDrawMleader(pt1, NumbVUStr, OfsetLeaderX, OfsetLeaderY, LayerMLeaderCurve, 0, Nothing, Nothing)
                                '==========================================================================================================
                                '=======================================================================================
                                If ArrayExcel(3, i) > 0 Or ArrayExcel(21, i) > 0 Or ArrayExcel(22, i) Then
                                    'находим координаты точек NPK и KPK
                                    Dim StartT1X As Double = -1
                                    Dim StartT1Y As Double = -1
                                    Dim EndT2X As Double = -1
                                    Dim EndT2Y As Double = -1
                                    Try
                                        Align.PointLocation(PKStartTan1, 0, StartT1X, StartT1Y)
                                        Align.PointLocation(PKEndTan2, 0, EndT2X, EndT2Y)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        GoTo line1
                                    End Try
                                    'рисуем линии тангенсов
                                    Dim PTStartT1 As Point2d = New Point2d(StartT1X, StartT1Y)
                                    Dim PTEndT2 As Point2d = New Point2d(EndT2X, EndT2Y)
                                    'определяем направления тангенсов
                                    Dim AngleT1 As Double = PTStartT1.GetVectorTo(pt1).Angle 'направление первого тангенса
                                    Dim AngleT2 As Double = pt1.GetVectorTo(PTEndT2).Angle 'направление второго тангенса
                                    'удлинняем первый тангенс
                                    Dim VnT1X As Double = (40 + PTStartT1.GetVectorTo(pt1).Length) * Math.Cos(AngleT1) + StartT1X
                                    Dim VnT1Y As Double = (40 + PTStartT1.GetVectorTo(pt1).Length) * Math.Sin(AngleT1) + StartT1Y
                                    'вставляем первый тангенс
                                    Dim ArrayVertexTan1 As Double(,) = Nothing
                                    ReDim Preserve ArrayVertexTan1(1, 1)
                                    ArrayVertexTan1(0, 0) = StartT1X
                                    ArrayVertexTan1(1, 0) = StartT1Y
                                    ArrayVertexTan1(0, 1) = pt1.X
                                    ArrayVertexTan1(1, 1) = pt1.Y
                                    Dim acPolyT1 As Autodesk.AutoCAD.DatabaseServices.Polyline = AcadFunction.FuncDrawPolylineByArrayProperties(ArrayVertexTan1, ArrayPropLinTH)
                                    'вставляем дорезок
                                    Dim ArrayVertexD As Double(,) = Nothing
                                    ReDim Preserve ArrayVertexD(1, 1)
                                    ArrayVertexD(0, 0) = pt1.X
                                    ArrayVertexD(1, 0) = pt1.Y
                                    ArrayVertexD(0, 1) = VnT1X
                                    ArrayVertexD(1, 1) = VnT1Y
                                    Dim acPolyT1D As Autodesk.AutoCAD.DatabaseServices.Polyline = AcadFunction.FuncDrawPolylineByArrayProperties(ArrayVertexD, ArrayPropLinTH)
                                    Dim ArrayVertexTan2 As Double(,) = Nothing
                                    ReDim Preserve ArrayVertexTan2(1, 1)
                                    ArrayVertexTan2(0, 0) = pt1.X
                                    ArrayVertexTan2(1, 0) = pt1.Y
                                    ArrayVertexTan2(0, 1) = PTEndT2.X
                                    ArrayVertexTan2(1, 1) = PTEndT2.Y
                                    'вставляем второй тангенс
                                    Dim acPolyT2 As Autodesk.AutoCAD.DatabaseServices.Polyline = AcadFunction.FuncDrawPolylineByArrayProperties(ArrayVertexTan2, ArrayPropLinTH)
                                    '=========================================================================================================
                                    '3. Вставляем дельта-угол
                                    Dim VnT2X As Double = (40) * Math.Cos(AngleT2) + pt1.X 'вторая точка вставки для угла
                                    Dim VnT2Y As Double = (40) * Math.Sin(AngleT2) + pt1.Y
                                    'точка вставки текста
                                    Dim ArcPtX As Double = (VnT1X + VnT2X) / 2
                                    Dim ArcPtY As Double = (VnT1Y + VnT2Y) / 2
                                    'вставляем угол
                                    Dim acLinAngDim As Point3AngularDimension = New Point3AngularDimension
                                    acLinAngDim.SetDatabaseDefaults()
                                    acLinAngDim.XLine1Point = New Point3d(VnT1X, VnT1Y, 0)
                                    acLinAngDim.XLine2Point = New Point3d(VnT2X, VnT2Y, 0)
                                    acLinAngDim.CenterPoint = New Point3d(pt1.X, pt1.Y, 0)
                                    acLinAngDim.ArcPoint = New Point3d(ArcPtX, ArcPtY, 0)
                                    acLinAngDim.DimensionStyle = acCurDb.Dimstyle
                                    acLinAngDim.Dimaunit = 1
                                    acLinAngDim.Dimdec = 5
                                    acLinAngDim.Prefix = "\U+0394="
                                    acLinAngDim.DimensionStyle = IDDimtyleDraw
                                    acLinAngDim.Layer = LayerDim
                                    acBlkTblRec.AppendEntity(acLinAngDim)
                                    tr.AddNewlyCreatedDBObject(acLinAngDim, True)
                                    '========================================================================================================
                                    '4.расставляем значения тангенсов
                                    Dim TXsr1 As Double = (PTStartT1.X + pt1.X) / 2
                                    Dim TYsr1 As Double = (PTStartT1.Y + pt1.Y) / 2
                                    Dim TXsr2 As Double = (pt1.X + PTEndT2.X) / 2
                                    Dim TYsr2 As Double = (pt1.Y + PTEndT2.Y) / 2
                                    'вычисляем позиции
                                    Select Case AngleT1
                                        Case 0 To Math.PI / 2
                                            AngleT1 = AngleT1
                                            Dim xIns As Double = TXsr1 + OffsetText * (Math.Cos(AngleT1 + Math.PI / 2))
                                            Dim yIns As Double = TYsr1 + OffsetText * (Math.Sin(AngleT1 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T1=" & T1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT1, 22)
                                        Case Math.PI / 2 To Math.PI
                                            AngleT1 = AngleT1 + Math.PI
                                            Dim xIns As Double = TXsr1 + OffsetText * (Math.Cos(AngleT1 + Math.PI / 2))
                                            Dim yIns As Double = TYsr1 + OffsetText * (Math.Sin(AngleT1 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T1=" & T1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT1, 22)
                                        Case Math.PI To (3 * Math.PI) / 2
                                            AngleT1 = AngleT1 - Math.PI
                                            Dim xIns As Double = TXsr1 + OffsetText * (Math.Cos(AngleT1 + Math.PI / 2))
                                            Dim yIns As Double = TYsr1 + OffsetText * (Math.Sin(AngleT1 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T1=" & T1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT1, 22)
                                        Case (3 * Math.PI) / 2 To Math.PI * 2
                                            AngleT1 = AngleT1
                                            Dim xIns As Double = TXsr1 + OffsetText * (Math.Cos(AngleT1 + Math.PI / 2))
                                            Dim yIns As Double = TYsr1 + OffsetText * (Math.Sin(AngleT1 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T1=" & T1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT1, 22)
                                    End Select
                                    Select Case AngleT2
                                        Case 0 To Math.PI / 2
                                            AngleT2 = AngleT2
                                            Dim xIns As Double = TXsr2 + OffsetText * (Math.Cos(AngleT2 + Math.PI / 2))
                                            Dim yIns As Double = TYsr2 + OffsetText * (Math.Sin(AngleT2 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T2=" & T2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT2, 22)
                                        Case Math.PI / 2 To Math.PI
                                            AngleT2 = AngleT2 + Math.PI
                                            Dim xIns As Double = TXsr2 + OffsetText * (Math.Cos(AngleT2 + Math.PI / 2))
                                            Dim yIns As Double = TYsr2 + OffsetText * (Math.Sin(AngleT2 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T2=" & T2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT2, 22)
                                        Case Math.PI To (3 * Math.PI) / 2
                                            AngleT2 = AngleT2 - Math.PI
                                            Dim xIns As Double = TXsr2 + OffsetText * (Math.Cos(AngleT2 + Math.PI / 2))
                                            Dim yIns As Double = TYsr2 + OffsetText * (Math.Sin(AngleT2 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T2=" & T2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT2, 22)
                                        Case (3 * Math.PI) / 2 To Math.PI * 2
                                            AngleT2 = AngleT2
                                            Dim xIns As Double = TXsr2 + OffsetText * (Math.Cos(AngleT2 + Math.PI / 2))
                                            Dim yIns As Double = TYsr2 + OffsetText * (Math.Sin(AngleT2 + Math.PI / 2))
                                            Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "T2=" & T2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleT2, 22)
                                    End Select
line1:
                                    '5. вставляем значения L1 и L2
                                    Dim L1Pt1 As Point2d = Nothing
                                    'Dim L1 As Double = ArrayExcel(4, i)
                                    If L1 > 0 Then
                                        'Dim L1str As String = FuncFormatZN.FuncFormatCoordinate(L1, FormRoadPlane.NumericUpDown6.Value)
                                        'ищем середину участка
                                        'находим координаты точек NPK и KPK
                                        Dim PKStartL1 As Double = ArrayExcel(12, i) 'NPK
                                        Dim PKEndL1 As Double = ArrayExcel(13, i) 'NKK
                                        If PKStartL1 = PKEndL1 Then
                                            PKEndL1 = ArrayExcel(15, i)
                                        End If
                                        'находим середину переходной кривой
                                        Dim PKSrd As Double = (PKStartL1 + PKEndL1) / 2
                                        'ищем приближенную касательную к середине прямой
                                        Dim PKSrdAng As Double = PKSrd + 0.01
                                        'ищем координаты
                                        Dim L1X As Double = -1 'координаты середины участка
                                        Dim L1Y As Double = -1
                                        Dim L1Xang As Double = -1 'координаты точки касательной
                                        Dim L1Yang As Double = -1
                                        Try
                                            Align.PointLocation(PKSrd, 0, L1X, L1Y)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo LineL1
                                        End Try
                                        If L1X <> -1 And L1Y <> -1 Then
                                            Dim AngleL1 As Double = 0
                                            L1Pt1 = New Point2d(L1X, L1Y)
                                            'ищем координаты точки касательной
                                            Try
                                                Align.PointLocation(PKSrdAng, 0, L1Xang, L1Yang)
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                            Dim L1Pt1ang As Point2d = Nothing
                                            If L1Xang <> -1 And L1Yang <> -1 Then
                                                L1Pt1ang = New Point2d(L1Xang, L1Yang)
                                                AngleL1 = L1Pt1.GetVectorTo(L1Pt1ang).Angle
                                                Select Case AngleL1
                                                    Case 0 To Math.PI / 2
                                                        Dim xIns As Double = L1X + OffsetText * (Math.Cos(AngleL1 + Math.PI / 2))
                                                        Dim yIns As Double = L1Y + OffsetText * (Math.Sin(AngleL1 + Math.PI / 2))
                                                        Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L1=" & L1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL1, 22)
                                                    Case Math.PI / 2 To Math.PI
                                                        AngleL1 = AngleL1 + Math.PI
                                                        Dim xIns As Double = L1X + OffsetText * (Math.Cos(AngleL1 + Math.PI / 2))
                                                        Dim yIns As Double = L1Y + OffsetText * (Math.Sin(AngleL1 + Math.PI / 2))
                                                        Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L1=" & L1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL1, 22)
                                                    Case Math.PI To (3 * Math.PI) / 2
                                                        AngleL1 = AngleL1 - Math.PI
                                                        Dim xIns As Double = L1X + OffsetText * (Math.Cos(AngleL1 + Math.PI / 2))
                                                        Dim yIns As Double = L1Y + OffsetText * (Math.Sin(AngleL1 + Math.PI / 2))
                                                        Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L1=" & L1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL1, 22)
                                                    Case (3 * Math.PI) / 2 To Math.PI * 2
                                                        Dim xIns As Double = L1X + OffsetText * (Math.Cos(AngleL1 + Math.PI / 2))
                                                        Dim yIns As Double = L1Y + OffsetText * (Math.Sin(AngleL1 + Math.PI / 2))
                                                        Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L1=" & L1Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL1, 22)
                                                End Select
                                            End If
                                        End If
                                    End If
LineL1:
                                    Dim L2Pt1 As Point2d = Nothing
                                    If L2 > 0 Then
                                        'ищем середину участка
                                        'находим координаты точек NPK и KPK
                                        Dim PKStartL2 As Double = ArrayExcel(14, i) 'KKK
                                        Dim PKEndL2 As Double = ArrayExcel(15, i) 'KPK
                                        If PKStartL2 = PKEndL2 Then
                                            PKStartL2 = ArrayExcel(12, i)
                                        End If
                                        Dim StartL2X As Double = -1
                                        Dim StartL2Y As Double = -1
                                        Dim EndL2X As Double = -1
                                        Dim EndL2Y As Double = -1
                                        Try
                                            Align.PointLocation(PKStartL2, 0, StartL2X, StartL2Y)
                                            Align.PointLocation(PKEndL2, 0, EndL2X, EndL2Y)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo LineL2
                                        End Try
                                        Dim PKSrd As Double = (PKStartL2 + PKEndL2) / 2
                                        'ищем приближенную касательную к середине прямой
                                        Dim PKSrdAng As Double = PKSrd + 0.01
                                        'ищем координаты
                                        Dim L2X As Double = -1 'координаты середины участка
                                        Dim L2Y As Double = -1
                                        Dim L2Xang As Double = -1 'координаты точки касательной
                                        Dim L2Yang As Double = -1
                                        Try
                                            Align.PointLocation(PKSrd, 0, L2X, L2Y)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo LineL1
                                        End Try
                                        If L2X <> -1 And L2Y <> -1 Then
                                            Dim AngleL2 As Double = 0 'угол поворота текста
                                            L2Pt1 = New Point2d(L2X, L2Y)
                                            'ищем координаты точки касательной
                                            Try
                                                Align.PointLocation(PKSrdAng, 0, L2Xang, L2Yang)
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                            Dim L1Pt1ang As Point2d = Nothing
                                            If L2Xang <> -1 And L2Yang <> -1 Then 'если координаты найдены ищем угол поворота текста
                                                Dim L2Pt1ang As Point2d = New Point2d(L2Xang, L2Yang)
                                                AngleL2 = L2Pt1.GetVectorTo(L2Pt1ang).Angle
                                            End If
                                            '==============================================================
                                            'редактируем местоположение текста
                                            Select Case AngleL2
                                                Case 0 To Math.PI / 2
                                                    Dim xIns As Double = L2X + OffsetText * (Math.Cos(AngleL2 + Math.PI / 2))
                                                    Dim yIns As Double = L2Y + OffsetText * (Math.Sin(AngleL2 + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L2=" & L2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL2, 22)
                                                Case Math.PI / 2 To Math.PI
                                                    AngleL2 = AngleL2 + Math.PI
                                                    Dim xIns As Double = L2X + OffsetText * (Math.Cos(AngleL2 + Math.PI / 2))
                                                    Dim yIns As Double = L2Y + OffsetText * (Math.Sin(AngleL2 + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L2=" & L2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL2, 22)
                                                Case Math.PI To (3 * Math.PI) / 2
                                                    AngleL2 = AngleL2 - Math.PI
                                                    Dim xIns As Double = L2X + OffsetText * (Math.Cos(AngleL2 + Math.PI / 2))
                                                    Dim yIns As Double = L2Y + OffsetText * (Math.Sin(AngleL2 + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L2=" & L2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL2, 22)
                                                Case (3 * Math.PI) / 2 To Math.PI * 2
                                                    Dim xIns As Double = L2X + OffsetText * (Math.Cos(AngleL2 + Math.PI / 2))
                                                    Dim yIns As Double = L2Y + OffsetText * (Math.Sin(AngleL2 + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "L2=" & L2Str, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleL2, 22)
                                            End Select
                                        End If
                                    End If
LineL2:
                                    '==========================================================
                                    'вставляем значение К
                                    Dim Ksohr As Double = ArrayExcel(9, i)
                                    If Ksohr > 0 Then
                                        Dim AngleK As Double = 0
                                        Dim Kstr As String = FuncFormatZN.FuncFormatCoordinate(Ksohr, FormRoadPlane.NumericUpDown6.Value)
                                        'ищем середину участка
                                        Dim PKStartK As Double = ArrayExcel(13, i) 'NKK
                                        Dim PKEndK As Double = ArrayExcel(14, i) 'KKK
                                        Dim StartKX As Double = -1
                                        Dim StartKY As Double = -1
                                        Dim EndKX As Double = -1
                                        Dim EndKY As Double = -1
                                        Try
                                            Align.PointLocation(PKStartK, 0, StartKX, StartKY)
                                            Align.PointLocation(PKEndK, 0, EndKX, EndKY)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo LineK
                                        End Try
                                        'ищем среднюю точку
                                        Dim PKSrd As Double = (PKStartK + PKEndK) / 2
                                        'ищем приближенную касательную к середине прямой
                                        Dim PKSrdAng As Double = PKSrd + 0.01
                                        'ищем координаты
                                        Dim KX As Double = -1 'координаты середины участка
                                        Dim KY As Double = -1
                                        Dim KXang As Double = -1 'координаты точки касательной
                                        Dim KYang As Double = -1
                                        Try
                                            Align.PointLocation(PKSrd, 0, KX, KY)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            GoTo LineK
                                        End Try
                                        If KX <> -1 And KY <> -1 Then
                                            Dim KPt As Point2d = New Point2d(KX, KY)
                                            'ищем координаты точки касательной
                                            Try
                                                Align.PointLocation(PKSrdAng, 0, KXang, KYang)
                                                Dim KPTang As Point2d = New Point2d(KXang, KYang)
                                                AngleK = KPt.GetVectorTo(KPTang).Angle
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                            '==============================================================
                                            'редактируем местоположение текста
                                            Select Case AngleK
                                                Case 0 To Math.PI / 2
                                                    Dim xIns As Double = KX + OffsetText * (Math.Cos(AngleK + Math.PI / 2))
                                                    Dim yIns As Double = KY + OffsetText * (Math.Sin(AngleK + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "K=" & Kstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleK, 22)
                                                Case Math.PI / 2 To Math.PI
                                                    AngleK = AngleK + Math.PI
                                                    Dim xIns As Double = KX + OffsetText * (Math.Cos(AngleK + Math.PI / 2))
                                                    Dim yIns As Double = KY + OffsetText * (Math.Sin(AngleK + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "K=" & Kstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleK, 22)
                                                Case Math.PI To (3 * Math.PI) / 2
                                                    AngleK = AngleK - Math.PI
                                                    Dim xIns As Double = KX + OffsetText * (Math.Cos(AngleK + Math.PI / 2))
                                                    Dim yIns As Double = KY + OffsetText * (Math.Sin(AngleK + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "K=" & Kstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleK, 22)
                                                Case (3 * Math.PI) / 2 To Math.PI * 2
                                                    Dim xIns As Double = KX + OffsetText * (Math.Cos(AngleK + Math.PI / 2))
                                                    Dim yIns As Double = KY + OffsetText * (Math.Sin(AngleK + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "K=" & Kstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleK, 22)
                                            End Select
                                        End If
                                    End If
LineK:
                                    '============================================================================================
                                    '6. вставляем биссектису
                                    If StartT1X <> -1 And StartT1Y <> -1 And EndT2X <> -1 And EndT2Y <> -1 Then
                                        Dim Vect_T1 As Vector2d = PTStartT1.GetVectorTo(pt1)
                                        Dim Vect_T2 As Vector2d = pt1.GetVectorTo(PTEndT2)
                                        Dim LineTX1 As Double = 10 * Math.Cos(Vect_T1.Angle + Math.PI) + pt1.X
                                        Dim LineTY1 As Double = 10 * Math.Sin(Vect_T1.Angle + Math.PI) + pt1.Y
                                        Dim LineTX2 As Double = 10 * Math.Cos(Vect_T2.Angle) + pt1.X
                                        Dim LineTY2 As Double = 10 * Math.Sin(Vect_T2.Angle) + pt1.Y
                                        Dim LineT1XSr As Double = (LineTX1 + LineTX2) / 2
                                        Dim LineT1YSr As Double = (LineTY1 + LineTY2) / 2
                                        Dim ptLineT1 As Point3d = New Point3d(LineT1XSr, LineT1YSr, 0)
                                        Dim pt13d As Point3d = New Point3d(pt1.X, pt1.Y, 0)
                                        Dim LineB As Line = New Line(pt13d, ptLineT1)
                                        Dim PointCol As Point3dCollection = New Point3dCollection
                                        Align.IntersectWith(LineB, Intersect.ExtendBoth, PointCol, 1, 1)
                                        If PointCol.Count > 0 Then
                                            'начинаем перебирать коллекцию
                                            Dim PtB2d As Point2d = Nothing
                                            For atp As Integer = 0 To PointCol.Count - 1
                                                Dim PtB As Point3d = PointCol.Item(atp)
                                                PtB2d = New Point2d(PtB.X, PtB.Y)
                                            Next
                                            Dim ArrayVertexB As Double(,) = Nothing
                                            ReDim Preserve ArrayVertexB(1, 1)
                                            ArrayVertexB(0, 0) = pt1.X
                                            ArrayVertexB(1, 0) = pt1.Y
                                            ArrayVertexB(0, 1) = PtB2d.X
                                            ArrayVertexB(1, 1) = PtB2d.Y
                                            Dim acPoly3 As Autodesk.AutoCAD.DatabaseServices.Polyline = AcadFunction.FuncDrawPolylineByArrayProperties(ArrayVertexB, ArrayPropLinTH)
                                            Dim XB As Double = (pt1.X + PtB2d.X) / 2
                                            Dim YB As Double = (pt1.Y + PtB2d.Y) / 2
                                            Dim AngleB As Double = pt1.GetVectorTo(PtB2d).Angle
                                            'редактируем местоположение текста
                                            Select Case AngleB
                                                Case 0 To Math.PI / 2
                                                    Dim xIns As Double = XB + OffsetText * (Math.Cos(AngleB + Math.PI / 2))
                                                    Dim yIns As Double = YB + OffsetText * (Math.Sin(AngleB + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "B=" & Bstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleB, 22)
                                                Case Math.PI / 2 To Math.PI
                                                    AngleB = AngleB + Math.PI
                                                    Dim xIns As Double = XB + OffsetText * (Math.Cos(AngleB + Math.PI / 2))
                                                    Dim yIns As Double = YB + OffsetText * (Math.Sin(AngleB + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "B=" & Bstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleB, 22)
                                                Case Math.PI To (3 * Math.PI) / 2
                                                    AngleB = AngleB - Math.PI
                                                    Dim xIns As Double = XB + OffsetText * (Math.Cos(AngleB + Math.PI / 2))
                                                    Dim yIns As Double = YB + OffsetText * (Math.Sin(AngleB + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "B=" & Bstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleB, 22)
                                                Case (3 * Math.PI) / 2 To Math.PI * 2
                                                    Dim xIns As Double = XB + OffsetText * (Math.Cos(AngleB + Math.PI / 2))
                                                    Dim yIns As Double = YB + OffsetText * (Math.Sin(AngleB + Math.PI / 2))
                                                    Dim insTextT1 As Boolean = AcadFunction.FuncDrawMText(New Point3d(xIns, yIns, 0), "B=" & Bstr, HText, LayerMLeaderCurve, IDTextStyleDraw, AngleB, 22)
                                            End Select
                                        End If
                                    End If
                                    '=============================================================================================
                                    'записываем радиусы
                                    'вставляем значение радиуса
                                    'ищем координаты середины конструкции
                                    Dim PTSK2d As Point2d = Nothing
                                    Dim PTCenter As Point3d = Nothing
                                    Dim PTSK3d As Point3d = Nothing
                                    'вычисляем точку конца полки
                                    If ArrayExcel(9, i) > 0 Then 'окружность существует вставляем радиус в середину дуги
                                        'в конструкции есть дуга вставляем ее радиус
                                        'находим координата начала дуги
                                        Dim xNK As Double = -1 'координаты начала кривой
                                        Dim yNK As Double = -1
                                        Try
                                            Align.PointLocation(ArrayExcel(13, i), 0, xNK, yNK)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                        'находим координата конца дуги
                                        Dim xKK As Double = -1 'координаты начала кривой
                                        Dim yKK As Double = -1
                                        Try
                                            Align.PointLocation(ArrayExcel(14, i), 0, xKK, yKK)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                        Dim PKSK As Double = (ArrayExcel(13, i) + ArrayExcel(14, i)) / 2
                                        'ищем координаты середины кривой
                                        Dim XSK As Double = -1
                                        Dim YSK As Double = -1
                                        Try
                                            Align.PointLocation(PKSK, 0, XSK, YSK)
                                            PTSK2d = New Point2d(XSK, YSK) 'точка середины дуги
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        End Try
                                        If xNK <> -1 And yNK <> -1 And xKK <> -1 And yKK <> -1 And XSK <> -1 And YSK <> -1 Then
                                            'ищем центр дуги
                                            Dim CircPT As Point2d = MathFunction.FuncFindCenterCircleByThreePoint(New Point3d(xNK, yNK, 0), New Point3d(XSK, YSK, 0), New Point3d(xKK, yKK, 0))
                                            Dim AngleR As Double = PTSK2d.GetVectorTo(CircPT).Angle
                                            Dim DimStartX As Double = PTSK2d.X + 50 * Math.Cos(AngleR)
                                            Dim DimStarty As Double = PTSK2d.X + 50 * Math.Sin(AngleR)
                                            PTSK3d = New Point3d(PTSK2d.X, PTSK2d.Y, 0)
                                            PTCenter = New Point3d(CircPT.X, CircPT.Y, 0)
                                            '===========================================================================================
                                            'вставляем размер радиуса
                                            Dim acRadDim As RadialDimension = New RadialDimension()
                                            acRadDim.SetDatabaseDefaults()
                                            acRadDim.Center = PTCenter
                                            acRadDim.ChordPoint = PTSK3d
                                            acRadDim.LeaderLength = -50
                                            acRadDim.DimensionStyle = IDDimtyleDraw
                                            acRadDim.Layer = LayerDim
                                            acBlkTblRec.AppendEntity(acRadDim)
                                            tr.AddNewlyCreatedDBObject(acRadDim, True)
                                        End If
                                    Else
                                        'есть только переходные кривые
                                        If ArrayExcel(4, i) > 0 And ArrayExcel(5, i) > 0 Then
                                            'обе переходные кривые существуют
                                            Dim PKSK As Double = (ArrayExcel(12, i) + ArrayExcel(15, i)) / 2
                                            'ищем координаты середины кривой
                                            Dim XSK As Double = -1
                                            Dim YSK As Double = -1
                                            Try
                                                Align.PointLocation(PKSK, 0, XSK, YSK)
                                                PTSK2d = New Point2d(XSK, YSK) 'точка середины дуги
                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            End Try
                                            'ищем центр дуги
                                            Dim CPt1 As Point2d = PTStartT1
                                            Dim CPT2 As Point2d = PTSK2d
                                            Dim CPT3 As Point2d = PTEndT2
                                            Dim CircPT As Point2d = MathFunction.FuncFindCenterCircleByThreePoint(New Point3d(CPt1.X, CPt1.Y, 0), New Point3d(CPT2.X, CPT2.Y, 0), New Point3d(CPT3.X, CPT3.Y, 0))
                                            Dim AngleR As Double = PTSK2d.GetVectorTo(CircPT).Angle
                                            Dim DimStartX As Double = PTSK2d.X + 50 * Math.Cos(AngleR)
                                            Dim DimStarty As Double = PTSK2d.X + 50 * Math.Sin(AngleR)
                                            PTSK3d = New Point3d(PTSK2d.X, PTSK2d.Y, 0)
                                            PTCenter = New Point3d(CircPT.X, CircPT.Y, 0)
                                            '===========================================================================================
                                            'вставляем размер радиуса
                                            Dim acRadDim As RadialDimension = New RadialDimension()
                                            acRadDim.SetDatabaseDefaults()
                                            acRadDim.Center = PTCenter
                                            acRadDim.ChordPoint = PTSK3d
                                            acRadDim.LeaderLength = -50
                                            acRadDim.DimensionStyle = IDDimtyleDraw
                                            acRadDim.Layer = LayerDim
                                            acRadDim.DimensionText = "R " & RStr
                                            acBlkTblRec.AppendEntity(acRadDim)
                                            tr.AddNewlyCreatedDBObject(acRadDim, True)
                                        End If
                                    End If
                                    '===============================================================================================
LineB:
                                End If
                            Next
                        End If
                        '===================================================================================
                        'вставляем точку конца трассы
                        pt1 = New Point2d(AlignEndX, AlignEndY)
                        Dim AttrPT3d2 As Point3d = New Point3d(pt1.X, pt1.Y, 0)
                        If acBlockTbl.Has(NameBlkVUP) = True Then
                            'вставляем блок
                            Dim bol As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(AttrPT3d2, ArrayPropBlkUPK, 0)
                        End If
                        '===============================================================================================
                        'вставляем мультивыноску с началом трассы
                        Dim PiketEndStr As String = PKEndAlign
                        Try
                            PiketEndStr = Align.GetStationStringWithEquations(PKEndAlign) 'преобразовываем RawPiket
                        Catch ex As System.ArgumentException
                        End Try
                        Dim strKT As String = "Конец трассы" & vbLf & PiketEndStr
                        Dim IDInsMleader2 As ObjectId = AcadFunction.FuncDrawMleader(pt1, strKT, OfsetLeaderX, OfsetLeaderY, LayerMLeaderCurve, 0, Nothing, Nothing)
                        'вставляем мультивыноску с координатами
                        CoordX = FuncFormatZN.FuncFormatCoordinate(pt1.X, RoundLineZn)
                        CoordY = FuncFormatZN.FuncFormatCoordinate(pt1.Y, RoundLineZn)
                        strCoord = "X=" & CoordY & vbLf & "Y=" & CoordX
                        IDInsMleader = AcadFunction.FuncDrawMleader(pt1, strCoord, OfsetLeaderX, -4, LayerMLeaderCurve, 0, Nothing, Nothing)
                    Else
                        'рисуем тангенциальный ход
                        If IsArray(ArrayExcel) = True Then
                            Dim pt1 As Point2d = New Point2d(AlignStartX, AlignStartY) 'координаты начала хода
                            Dim acPolyTH As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline() 'полилиния
                            acPolyTH.SetDatabaseDefaults()
                            'добавляем первую вершину
                            acPolyTH.AddVertexAt(0, pt1, 0, 0, 0)
                            'идем циклом по всем вершинам
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                pt1 = New Point2d(ArrayExcel(19, i), ArrayExcel(20, i))
                                acPolyTH.AddVertexAt(i + 1, pt1, 0, 0, 0)
                            Next
                            'точка с координатами
                            pt1 = New Point2d(AlignEndX, AlignEndY)
                            Dim N As Integer = ArrayExcel.GetUpperBound(1) + 2
                            acPolyTH.AddVertexAt(N, pt1, 0, 0, 0)
                            acPolyTH.Layer = LayerMLeaderCoord
                            acBlkTblRec.AppendEntity(acPolyTH)
                            tr.AddNewlyCreatedDBObject(acPolyTH, True)
                        End If

                    End If
                    '===================================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'вставляем элементы кривых
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    '=================================================================================================================
                    'расставляем рубленные пикеты
                    '=================================================================================================================
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    If FormRoadPlane.CheckBox1.Checked = True Then
                        Dim CountInsRublPK As Integer = FuncAlignment.FuncInsertRublPK(Align, ArrayPropBlkRPK, RoundLineZn)
                    End If
                    tr.Commit()
                Catch ex As System.AccessViolationException
                Catch ex As System.Exception
                End Try
            End Using
lineEnd:
        End Using
    End Sub
End Class
