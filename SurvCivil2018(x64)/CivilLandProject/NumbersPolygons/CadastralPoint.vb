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
Imports CadastralTools
Imports NetTopologySuite
Imports NetTopologySuite.Geometries

Public Class CadastralPoint
    '===================================================================================================================
    'автонумерация полигонов
    '===================================================================================================================
    <CommandMethod("AutoNumberCadastralObject")>
    Public Sub AutoNumberCadastralObject()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        Dim formNumbersPoint As CadastralTools.FormCreatePointCadastral = New CadastralTools.FormCreatePointCadastral(MAnno)
        'вызываем диалоговое окно УСТАНОВКИ
        formNumbersPoint.ShowDialog()
        If formNumbersPoint.boolShow = False Then
            Exit Sub
        End If
        'путь к шаблону
        Dim templateDWG As String = formNumbersPoint.PutchTemlateDwg
        Dim templateXML As String = formNumbersPoint.PutchTemlateXML
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        Dim idStyle As ObjectId = Nothing 'текстовые стиль
        'префикс номера
        Dim prefNumber As String = formNumbersPoint.prefNumber
        'метод определения координат
        Dim MetodCoord As String = formNumbersPoint.metodCoord
        'способ закрепления
        Dim Zakrep As String = formNumbersPoint.zakrPoint
        'погрешность определения точки
        Dim Delta As String = formNumbersPoint.icnarisityPoint
        'поиск дубликатов
        Dim boolDoublePoint As Boolean = formNumbersPoint.boolDublePoint
        Dim tolerance As Double = formNumbersPoint.tolerance
        'округрение вершин
        Dim boolRoundVertex As Boolean = formNumbersPoint.boolRoundVertex
        'по часовой стрелке
        Dim boolClock As Boolean = formNumbersPoint.boolClockVertex
        'северо-западный угол
        Dim boolSVAngle As Boolean = formNumbersPoint.boolSEVertex
        '========================================================================================
        Dim M As Double = formNumbersPoint.scalePoint * 1000 'масштаб
        Dim HText As Double = (M * formNumbersPoint.hText) / 1000 ' высота подписи
        Dim DText As Double = (M * formNumbersPoint.offsetText) / 1000 'расстояние смещения
        '==================================================================================================
        'имя блока
        Dim nameBlk As String = ""
        Dim rotationBlk As Double = 0
        '===========================================================================================
        'характеристика атрибута
        Dim nameAttributeNumber As String = formNumbersPoint.nameAttribyteNumber
        Dim attributeHeight As Double = (HText * M) / 1000
        Dim arrayAttributes As String(,) = formNumbersPoint.arrayAttribytes
        '============================================================================================
        Dim categoryZN As String = formNumbersPoint.categoryZN
        Dim tableZn As String = formNumbersPoint.tableZN
        Dim typeObject As CadastralTools.StructureElement.typeGeometryObject = StructureElement.typeGeometryObject.block
        Dim typeLineObject As CadastralTools.StructureElement.typeGeometryObject = formNumbersPoint.typeGeometry
        Dim descrittionBlockStyle As String = formNumbersPoint.deskBlockStyle
        Dim descriptionObjectStyle As String = formNumbersPoint.deskObjectStyle
        Dim NamePSTable As String = "CADASTRAL_POINT"
        Dim idTablePs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(templateXML, NamePSTable)
        '===============================================================================================
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        Dim arrayPropertiesBlock As String(,) = {}
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем транзакцию
        Dim styleBlockReference As CadastralBlockStyle = New CadastralBlockStyle()
        Dim acadNetTopogy As FuncNetTopologySuite = New FuncNetTopologySuite()
        Dim cadastralNetTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using locker
            Using tr
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'копируем блок
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                styleBlockReference.DBDrawing = db
                styleBlockReference.TemplateDWG = templateDWG
                styleBlockReference.TemplateXML = templateXML
                Dim boolReadPropertiesBlock As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZN, tableZn, typeObject, descrittionBlockStyle, arrayPropertiesBlock)
                If arrayPropertiesBlock.Length > 0 Then
                    styleBlockReference.createBlockStyle(arrayPropertiesBlock)
                End If
                styleBlockReference.ScaleBlock = (styleBlockReference.ScaleBlock * M) / 1000
                nameBlk = styleBlockReference.NameBlock
                If IsNothing(nameBlk) = True Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                If acBlkTbl.Has(nameBlk) = False Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                '====================================================================================================
                'фильтуем точки чертежа в пределах видимости
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(db, nameBlk)
                Dim LargNumber As Integer = 1
                If listBlockReference.Count > 0 Then
                    'находим наибольший номер
                    LargNumber = CadastralTools.CadastralPoint.getLargNumber(listBlockReference, nameAttributeNumber)
                End If
                '====================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'начальный номер точки
                Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер точки: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowNone = True
                pStrOpts.AllowZero = False
                pStrOpts.DefaultValue = LargNumber
                Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts) 'запрашиваем номер
                If pStrRes.Status = PromptStatus.OK Then
                    Dim numberPoint As Integer = pStrRes.Value
                    Dim dictCadastralParcels As New Dictionary(Of Integer, String(,))
                    Try
                        'выбираем примитивы
                        Dim CountParcels As Integer = 0
                        Dim SelectParcel As PromptSelectionResult = ed.GetSelection()
                        If SelectParcel.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                            Dim acSSet1 As SelectionSet = SelectParcel.Value
                            For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                Dim entPLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(IDObj, OpenMode.ForRead, False)
                                If TypeOf entPLine Is Polyline Then
                                    Dim deskObject As String = ""
                                    Dim dataPolyline As StructureElement = Nothing
                                    Dim boolReadData As Boolean = FuncXRecords.getXRecords(entPLine, dataPolyline)
                                    If boolReadData = True Then
                                        Dim PSObject As String = dataPolyline.TablePS
                                        If PSObject Like "CADASTRAL_OSTROV" Or PSObject Like "CADASTRAL_OBJECT" Then
                                            Continue For
                                        End If
                                        deskObject = dataPolyline.Description
                                        Dim keyParam As String = dataPolyline.KeyParameter
                                        If deskObject Like descriptionObjectStyle Then
                                            Dim categoryEarth As String = ""
                                            Dim ConditionNumber As String = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "ConditionalNumber")
                                            If Delta <= 0 Then
                                                categoryEarth = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Category")
                                            End If
                                            '========================================================================================
                                            'читаем расширенные данные
                                            Dim LABEL As String = ""
                                            ReDim Preserve ArrayTablesXML(1, 0)
                                            ArrayTablesXML(0, 0) = dataPolyline.TablePS
                                            Dim ArrayData As String(,) = {}
                                            Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(entPLine, ArrayData, ArrayTablesXML)
                                            If ArrayData.Length > 0 Then
                                                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                    Dim NTbl As String = ArrayData(0, i)
                                                    Dim NField As String = ArrayData(1, i)
                                                    Dim Zn As String = ArrayData(2, i)
                                                    If NField Like "LABEL" Then
                                                        deskObject = Zn
                                                    ElseIf NField Like "ConditionalNumber" Then
                                                        Dim tempConditionalNumber As String = Zn.Trim
                                                        If Not (tempConditionalNumber Like ConditionNumber) Then
                                                            keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", tempConditionalNumber)
                                                            dataPolyline.KeyParameter = keyParam
                                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entPLine, dataPolyline)
                                                        End If
                                                        ConditionNumber = Zn
                                                    ElseIf NField Like "Category" Then
                                                        Dim tempCategory As String = Zn.Trim
                                                        If Not (tempCategory Like categoryEarth) Then
                                                            keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "Category", tempCategory)
                                                            dataPolyline.KeyParameter = keyParam
                                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entPLine, dataPolyline)
                                                        End If
                                                        categoryEarth = Zn
                                                    End If
                                                Next i
                                            Else
                                                Continue For
                                            End If
                                            If formNumbersPoint.boolIcnarisity = True Then
                                                If formNumbersPoint.icnarisityPoint > 0 Then
                                                    Delta = formNumbersPoint.icnarisityPoint
                                                End If
                                            End If
                                            If IsNothing(ConditionNumber) = False Then
                                                If ConditionNumber.Trim.Length > 0 Then
                                                    Dim numberParcel As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(ConditionNumber)
                                                    '=====================================================================================================
                                                    If dictCadastralParcels.ContainsKey(numberParcel) = False Then
                                                        ReDim Preserve ArrayPoligon(3, 0)
                                                        ArrayPoligon(0, 0) = numberParcel 'номер участка
                                                        ArrayPoligon(1, 0) = entPLine.Handle.ToString 'хангл полилинии
                                                        ArrayPoligon(2, 0) = dataPolyline.TablePS
                                                        If formNumbersPoint.boolIcnarisity = True Then
                                                            Delta = formNumbersPoint.icnarisityPoint
                                                        Else
                                                            If categoryEarth.Trim.Length > 0 Then
                                                                Delta = CadastralTools.CadastralInfo.getDeltaPositionPointByCategoryEarth(categoryEarth)
                                                            Else
                                                                Delta = formNumbersPoint.icnarisityPoint
                                                            End If
                                                        End If
                                                        ArrayPoligon(3, 0) = Delta
                                                        dictCadastralParcels.Add(numberParcel, ArrayPoligon)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            tr.Abort()
                            locker.Dispose()
                            MsgBox("Прервано пользователем.")
                            Exit Sub
                        End If
                        '==================================================================================================
                        If dictCadastralParcels.Count = 0 Then
                            tr.Abort()
                            locker.Dispose()
                            MsgBox("Массив с участками пуст!!!")
                            Exit Sub
                        Else
                            If dictCadastralParcels.Count > 1 Then
                                dictCadastralParcels = dictCadastralParcels.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
                            End If
                        End If
                        '====================================================================================================
                        'выбираем кадастровые полигоны
                        Dim dictCadastralObject As New Dictionary(Of String, List(Of Geometries.Geometry))
                        If typeLineObject = StructureElement.typeGeometryObject.polygon Then
                            dictCadastralObject = SelectCadastralObject.getCadastralPolygon(db, descriptionObjectStyle)
                        ElseIf typeLineObject = StructureElement.typeGeometryObject.polyline Then
                            dictCadastralObject = SelectCadastralObject.getCadastralPolyline(db, descriptionObjectStyle)
                        End If
                        If dictCadastralObject.Count = 0 Then
                            MsgBox("Не удалось найти объекты на чертеже типа " & typeLineObject.ToString)
                            tr.Abort()
                            locker.Dispose()
                            Exit Sub
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'нумеруем полигоны
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                        For i As Integer = 0 To dictCadastralParcels.Count - 1
                            ArrayPoligon = dictCadastralParcels.ElementAt(i).Value
                            'получаем доступ к полилинии для извлечения координат
                            Dim numberParcel As Integer = Val(ArrayPoligon(0, 0))
                            Dim handlePLine As String = ArrayPoligon(1, 0)
                            Dim tablePS As String = ArrayPoligon(2, 0)
                            Delta = ArrayPoligon(3, 0)
                            Dim listParcels As List(Of NetTopologySuite.Geometries.Geometry) = dictCadastralObject.Item(descriptionObjectStyle)
                            'забираем все контура для выбранного номера участка
                            Dim dictCounters As Dictionary(Of Integer, NetTopologySuite.Geometries.Geometry) = Nothing
                            If typeLineObject = StructureElement.typeGeometryObject.polygon Then
                                dictCounters = CadastralTools.CadastralPolygon.getCountersByNumber(numberParcel, listParcels)
                            ElseIf typeLineObject = StructureElement.typeGeometryObject.polyline Then
                                dictCounters = CadastralTools.CadastralLine.getCountersByNumber(numberParcel, listParcels)
                            End If
                            If IsNothing(dictCounters) = False Then
                                If dictCounters.Count > 0 Then
                                    For j As Integer = 0 To dictCounters.Count - 1
                                        Dim tempCounter As NetTopologySuite.Geometries.Geometry = dictCounters.ElementAt(j).Value
                                        Dim triplets As New List(Of (VertexIndex As Integer, Prev As Coordinate, Current As Coordinate, [Next] As Coordinate))
                                        Dim dataCounter As StructureElement = Nothing
                                        If TypeOf tempCounter Is NetTopologySuite.Geometries.Polygon Then
                                            Dim counter As NetTopologySuite.Geometries.Polygon = tempCounter
                                            dataCounter = counter.UserData
                                            counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                            If boolSVAngle = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.NorthWestVertexFirst(counter)
                                            End If
                                            If boolClock = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.ReversePolygonPreserveOrientation(counter)
                                            End If
                                            If boolRoundVertex = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.RoundPolygonVertices(counter, tolerance)
                                            End If
                                            Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromPolygon(counter)
                                            triplets = CadastralTools.CadastralPoint.GetPolygonVertexTriplets(counter)
                                        ElseIf TypeOf tempCounter Is NetTopologySuite.Geometries.LineString Then
                                            Dim counter As NetTopologySuite.Geometries.LineString = tempCounter
                                            dataCounter = counter.UserData
                                            counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                            If boolRoundVertex = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.RoundLineStringVertices(counter, tolerance)
                                            End If
                                            Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromLineString(counter)
                                            triplets = CadastralTools.CadastralPoint.GetLineStringVertexTriplets(counter)
                                        End If
                                        If triplets.Count > 0 Then
                                            For Each t In triplets
                                                Dim insAttr As Double() = {}
                                                Dim Vert0 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Prev.X, t.Prev.Y)
                                                Dim Vert1 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Current.X, t.Current.Y)
                                                Dim Vert2 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Next.X, t.Next.Y)
                                                'ищем уже существующую точку
                                                If boolDoublePoint = True Then
                                                    Dim boolFindPoint As Boolean = CadastralTools.CadastralPoint.HasPoint(Vert1, listBlockReference, "", tolerance)
                                                    If boolFindPoint = True Then
                                                        Continue For
                                                    End If
                                                End If
                                                Dim boolInsAttr As Boolean = CadastralTools.CadastralPoint.calculatePositionAttributeBlock(Vert0, Vert1, Vert2, insAttr, DText, True)
                                                'заполняем атрибуты
                                                Dim listAttributes As List(Of AttributeCadastralBlock) = New List(Of AttributeCadastralBlock)
                                                For k As Integer = 1 To arrayAttributes.GetUpperBound(0)
                                                    Dim attribyteNumberPoint As AttributeCadastralBlock = New AttributeCadastralBlock()
                                                    'номер
                                                    attribyteNumberPoint.TagAttribute = arrayAttributes(k, 0)
                                                    attribyteNumberPoint.DefinitionAttribute = arrayAttributes(k, 1)
                                                    If attribyteNumberPoint.TagAttribute Like nameAttributeNumber Then
                                                        attribyteNumberPoint.TextStringAttribute = prefNumber & numberPoint
                                                        attribyteNumberPoint.HeightAttribute = HText
                                                        attribyteNumberPoint.XAttribute = insAttr(0)
                                                        attribyteNumberPoint.YAttribute = insAttr(1)
                                                        attribyteNumberPoint.HorizontalAlignment = insAttr(2)
                                                        attribyteNumberPoint.VerticalAlignment = insAttr(3)
                                                    ElseIf attribyteNumberPoint.TagAttribute Like "DELTA" Then
                                                        attribyteNumberPoint.TextStringAttribute = Delta
                                                    ElseIf attribyteNumberPoint.TagAttribute Like "FASTDESK" Then
                                                        attribyteNumberPoint.TextStringAttribute = Zakrep
                                                    ElseIf attribyteNumberPoint.TagAttribute Like "METHODCOORDINATES" Then
                                                        attribyteNumberPoint.TextStringAttribute = MetodCoord
                                                    End If
                                                    listAttributes.Add(attribyteNumberPoint)
                                                Next k
                                                'вставляем блок
                                                Dim insertPointBlk As Point3d = New Point3d(t.Current.X, t.Current.Y, 0)
                                                Dim insertBlk As BlockReference = CreateDWGObject.CreateBlock(insertPointBlk, styleBlockReference, listAttributes)
                                                If IsNothing(insertBlk) = False Then
                                                    Dim dataPoint As CadastralTools.StructureElement = CadastralTools.CadastralPoint.createDataPoint(categoryZN, tableZn, descrittionBlockStyle, dataCounter.IdElement, numberPoint, prefNumber, Delta, Zakrep, MetodCoord, insertBlk.Handle.ToString)
                                                    Dim dictData As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(dataPoint.KeyParameter)
                                                    Dim recData As Boolean = FuncXRecords.writeXRecords(insertBlk, dataPoint)
                                                    If IsNothing(idTablePs) = False Then
                                                        Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(insertBlk, idTablePs)
                                                        If boolADDPSTBL = True Then
                                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(insertBlk, dataPoint)
                                                        End If
                                                    End If
                                                End If
                                                If boolDoublePoint = True Then
                                                    Dim newDictPoint As New Dictionary(Of String, String)
                                                    newDictPoint.Add(nameAttributeNumber, numberPoint)
                                                    Vert1.UserData = newDictPoint
                                                    listBlockReference.Add(insertBlk.Handle.ToString, Vert1)
                                                End If
                                                numberPoint += 1
                                            Next
                                        End If
                                    Next j
                                End If
                            End If
                        Next i
                    Catch ex As System.Exception
                        MsgBox("Ошибка!!! " & ex.ToString)
                        tr.Abort()
                        locker.Dispose()
                    End Try
                Else
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Прервано пользователем.")
                    Exit Sub
                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    '===================================================================================================================
    'ручная нумерация полигонов: точку вставки блока указывает пользователь
    'Установки — то же диалоговое окно и тот же стиль блока, что и в AutoNumberCadastralObject
    '===================================================================================================================
    <CommandMethod("CreateCadastralPoint")>
    Public Sub CreateCadastralPoint()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ — то же окно, что и в автонумерации
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        Dim formNumbersPoint As CadastralTools.FormCreatePointCadastral = New CadastralTools.FormCreatePointCadastral(MAnno)
        formNumbersPoint.ShowDialog()
        If formNumbersPoint.boolShow = False Then
            Exit Sub
        End If
        'путь к шаблону
        Dim templateDWG As String = formNumbersPoint.PutchTemlateDwg
        Dim templateXML As String = formNumbersPoint.PutchTemlateXML
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        'префикс номера
        Dim prefNumber As String = formNumbersPoint.prefNumber
        'метод определения координат
        Dim MetodCoord As String = formNumbersPoint.metodCoord
        'способ закрепления
        Dim Zakrep As String = formNumbersPoint.zakrPoint
        'погрешность определения точки
        Dim Delta As String = formNumbersPoint.icnarisityPoint
        'допуск поиска уже существующей точки в вершине
        Dim tolerance As Double = formNumbersPoint.tolerance
        'округрение вершин и порядок обхода в ручном режиме не используются:
        'место вставки блока задаёт пользователь
        '========================================================================================
        Dim M As Double = formNumbersPoint.scalePoint * 1000 'масштаб
        Dim HText As Double = (M * formNumbersPoint.hText) / 1000 ' высота подписи
        Dim DText As Double = (M * formNumbersPoint.offsetText) / 1000 'расстояние смещения
        '==================================================================================================
        'характеристика атрибута
        Dim nameAttributeNumber As String = formNumbersPoint.nameAttribyteNumber
        Dim arrayAttributes As String(,) = formNumbersPoint.arrayAttribytes
        '============================================================================================
        Dim categoryZN As String = formNumbersPoint.categoryZN
        Dim tableZn As String = formNumbersPoint.tableZN
        Dim descrittionBlockStyle As String = formNumbersPoint.deskBlockStyle
        Dim descriptionObjectStyle As String = formNumbersPoint.deskObjectStyle
        Dim NamePSTable As String = "CADASTRAL_POINT"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim styleBlockReference As CadastralBlockStyle = New CadastralBlockStyle()
        Dim arrayPropertiesBlock As String(,) = {}
        Dim CountPoint As Integer = 0 'сколько точек вставлено
        Dim numberPoint As Integer = 0 'текущий номер точки
        Dim nameBlk As String = "" 'имя блока кадастровой точки
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем блок — так же, как в автонумерации
            'Транзакция подготовки короткая: её обязательно нужно закрыть ДО цикла нумерации,
            'иначе вложенные транзакции вставки не попадают в чертёж до конца команды,
            'и поставленные точки не видны на экране
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                styleBlockReference.DBDrawing = db
                styleBlockReference.TemplateDWG = templateDWG
                styleBlockReference.TemplateXML = templateXML
                Dim boolReadPropertiesBlock As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZN, tableZn, StructureElement.typeGeometryObject.block, descrittionBlockStyle, arrayPropertiesBlock)
                If arrayPropertiesBlock.Length > 0 Then
                    styleBlockReference.createBlockStyle(arrayPropertiesBlock)
                End If
                styleBlockReference.ScaleBlock = (styleBlockReference.ScaleBlock * M) / 1000
                nameBlk = styleBlockReference.NameBlock
                If IsNothing(nameBlk) = True Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                If acBlkTbl.Has(nameBlk) = False Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                tr.Commit()
            End Using
            '====================================================================================================
            'фильтруем существующие точки чертежа и находим наибольший номер
            Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(db, nameBlk)
            Dim LargNumber As Integer = 1
            If listBlockReference.Count > 0 Then
                LargNumber = CadastralTools.CadastralPoint.getLargNumber(listBlockReference, nameAttributeNumber)
            End If
            '====================================================================================================
            'начальный номер точки
            Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер точки: ")
            pStrOpts.AllowNegative = False
            pStrOpts.AllowNone = True
            pStrOpts.AllowZero = False
            pStrOpts.DefaultValue = LargNumber
            Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts) 'запрашиваем номер
            If pStrRes.Status <> PromptStatus.OK Then
                locker.Dispose()
                MsgBox("Прервано пользователем.")
                Exit Sub
            End If
            numberPoint = pStrRes.Value
            '====================================================================================================
            'таблица свойств для точек
            Dim idTablePs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(templateXML, NamePSTable)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ручная нумерация: пользователь сам указывает точку вставки блока
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ed.WriteMessage(vbLf & "Начальный номер точки: " & prefNumber & numberPoint)
            Dim peo As PromptEntityOptions = New PromptEntityOptions(vbLf & "Укажите точку вставки блока (Enter — завершить нумерацию): ")
            peo.SetRejectMessage(vbLf & "Выберите полилинию (контур участка).")
            peo.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), False)
            peo.AllowNone = True
            Try
                Do
                    Dim per As PromptEntityResult = ed.GetEntity(peo)
                    If per.Status = PromptStatus.None Then
                        Exit Do
                    ElseIf per.Status <> PromptStatus.OK Then
                        ed.WriteMessage(vbLf & "Нумерация прервана пользователем.")
                        Exit Do
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'каждая точка — в своей транзакции: сразу после указания места она фиксируется
                    'и становится видна в чертеже
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Using trPoint As Transaction = db.TransactionManager.StartTransaction()
                        '--- выбранная полилиния ---
                        Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(trPoint.GetObject(per.ObjectId, OpenMode.ForRead, False), Autodesk.AutoCAD.DatabaseServices.Polyline)
                        If IsNothing(Pline) = True Then
                            trPoint.Abort()
                            ed.WriteMessage(vbLf & "Выбранный объект не является полилинией.")
                            Continue Do
                        End If
                        '--- данные кадастрового объекта ---
                        Dim dataPolyline As StructureElement = Nothing
                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(Pline, dataPolyline)
                        If boolReadData = False OrElse IsNothing(dataPolyline) = True Then
                            trPoint.Abort()
                            ed.WriteMessage(vbLf & "У выбранной полилинии нет данных кадастрового объекта.")
                            Continue Do
                        End If
                        If dataPolyline.TablePS Like "CADASTRAL_OBJECT" Then
                            trPoint.Abort()
                            ed.WriteMessage(vbLf & "Этот контур не является участком (TablePS = " & dataPolyline.TablePS & ").")
                            Continue Do
                        End If
                        'внутренние островки нумеруем независимо от настроенного стиля объекта
                        If Not (dataPolyline.Description Like descriptionObjectStyle) Then
                            If Not (dataPolyline.Description Like "Внутренний островок") Then
                                trPoint.Abort()
                                ed.WriteMessage(vbLf & "Стиль объекта не совпадает с настройкой """ & descriptionObjectStyle & """.")
                                Continue Do
                            End If
                        End If
                        '--- погрешность по категории земель (как в автонумерации) ---
                        Dim DeltaPoint As String = Delta
                        Dim keyParam As String = dataPolyline.KeyParameter
                        Dim categoryEarth As String = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Category")
                        If IsNothing(categoryEarth) = False Then
                            If categoryEarth.Trim.Length > 0 Then
                                DeltaPoint = CadastralTools.CadastralInfo.getDeltaPositionPointByCategoryEarth(categoryEarth)
                            End If
                        End If
                        '--- точка вставки: ближайшая к указанному месту вершина полилинии ---
                        Dim insertPoint As Point3d = GetNearestVertexPoint(Pline, per.PickedPoint)
                        '--- если в этой вершине точка уже есть, вторую не ставим ---
                        Dim vertexNtsPoint As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(insertPoint.X, insertPoint.Y)
                        Dim boolFindPoint As Boolean = CadastralTools.CadastralPoint.HasPoint(vertexNtsPoint, listBlockReference, "", tolerance)
                        If boolFindPoint = True Then
                            trPoint.Abort()
                            ed.WriteMessage(vbLf & "В этой вершине уже есть точка.")
                            Continue Do
                        End If
                        '--- положение номера: направление как в автонумерации, отсчёт от вершины ---
                        Dim insAttr As Double() = {}
                        Dim boolInsAttr As Boolean = GetPointAttributePosition(Pline, insertPoint, DText, insAttr)
                        If boolInsAttr = False Then
                            trPoint.Abort()
                            ed.WriteMessage(vbLf & "Не удалось рассчитать положение номера точки.")
                            Continue Do
                        End If
                        '--- заполняем атрибуты ---
                        Dim listAttributes As List(Of AttributeCadastralBlock) = New List(Of AttributeCadastralBlock)
                        If IsNothing(arrayAttributes) = False Then
                            If arrayAttributes.Length > 0 Then
                                For k As Integer = 1 To arrayAttributes.GetUpperBound(0)
                                    Dim attribyteNumberPoint As AttributeCadastralBlock = New AttributeCadastralBlock()
                                    'номер
                                    attribyteNumberPoint.TagAttribute = arrayAttributes(k, 0)
                                    attribyteNumberPoint.DefinitionAttribute = arrayAttributes(k, 1)
                                    If attribyteNumberPoint.TagAttribute Like nameAttributeNumber Then
                                        attribyteNumberPoint.TextStringAttribute = prefNumber & numberPoint
                                        attribyteNumberPoint.HeightAttribute = HText
                                        attribyteNumberPoint.XAttribute = insAttr(0)
                                        attribyteNumberPoint.YAttribute = insAttr(1)
                                        attribyteNumberPoint.HorizontalAlignment = insAttr(2)
                                        attribyteNumberPoint.VerticalAlignment = insAttr(3)
                                    ElseIf attribyteNumberPoint.TagAttribute Like "DELTA" Then
                                        attribyteNumberPoint.TextStringAttribute = DeltaPoint
                                    ElseIf attribyteNumberPoint.TagAttribute Like "FASTDESK" Then
                                        attribyteNumberPoint.TextStringAttribute = Zakrep
                                    ElseIf attribyteNumberPoint.TagAttribute Like "METHODCOORDINATES" Then
                                        attribyteNumberPoint.TextStringAttribute = MetodCoord
                                    End If
                                    listAttributes.Add(attribyteNumberPoint)
                                Next k
                            End If
                        End If
                        '--- вставляем блок в указанную точку ---
                        Dim insertBlk As BlockReference = CreateDWGObject.CreateBlock(insertPoint, styleBlockReference, listAttributes)
                        If IsNothing(insertBlk) = False Then
                            Dim idElement As String = ""
                            If IsNothing(dataPolyline.IdElement) = False Then
                                idElement = dataPolyline.IdElement
                            End If
                            Dim dataPoint As CadastralTools.StructureElement = CadastralTools.CadastralPoint.createDataPoint(categoryZN, tableZn, descrittionBlockStyle, idElement, numberPoint, prefNumber, Val(DeltaPoint), Zakrep, MetodCoord, insertBlk.Handle.ToString)
                            Dim recData As Boolean = FuncXRecords.writeXRecords(insertBlk, dataPoint)
                            If IsNothing(idTablePs) = False Then
                                Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(insertBlk, idTablePs)
                                If boolADDPSTBL = True Then
                                    Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(insertBlk, dataPoint)
                                End If
                            End If
                            ed.WriteMessage(vbLf & "Вставлена точка " & prefNumber & numberPoint & ".")
                            'запоминаем вставленную точку, чтобы повторно её не поставить
                            listBlockReference.Add(insertBlk.Handle.ToString, vertexNtsPoint)
                            numberPoint += 1
                            CountPoint += 1
                        Else
                            ed.WriteMessage(vbLf & "Не удалось вставить блок. Проверьте имя блока в чертеже.")
                        End If
                        'фиксируем вставку сразу — точка появляется в чертеже до конца команды
                        trPoint.Commit()
                    End Using
                    'обновляем изображение
                    Application.UpdateScreen()
                Loop
            Catch ex As System.Exception
                MsgBox("Ошибка!!! " & ex.ToString)
                locker.Dispose()
                Exit Sub
            End Try
        End Using
        If CountPoint > 0 Then
            ed.WriteMessage(vbLf & "Ручная нумерация завершена. Вставлено точек: " & CountPoint & ". Последний номер: " & prefNumber & (numberPoint - 1))
        Else
            ed.WriteMessage(vbLf & "Точки не вставлены.")
        End If
    End Sub
    '===================================================================================================================
    'положение номера точки для ручной нумерации
    'Смещение отсчитывается от указанной точки (обычно это ближайшая вершина полилинии),
    'а всю геометрию считает CadastralTools.CadastralPoint.GetPositionAttributeByLine —
    'она не зависит от CAD и одинаково работает с замкнутыми и разомкнутыми линиями.
    'Возвращает массив: X, Y, ГорВыр(1-лево 2-право), ВертВыр(1-низ 2-верх)
    '===================================================================================================================
    Private Shared Function GetPointAttributePosition(ByVal Pline As Autodesk.AutoCAD.DatabaseServices.Polyline, ByVal insertPoint As Point3d, ByVal Dist As Double, ByRef insAttr As Double()) As Boolean
        GetPointAttributePosition = False
        If IsNothing(Pline) = True Then Return False
        Dim count As Integer = Pline.NumberOfVertices
        If count < 2 Then Return False

        'линия NTS из вершин полилинии.
        'У замкнутой полилинии AutoCAD последняя вершина НЕ дублирует первую, поэтому
        'замыкаем список координат сами — тогда NTS считает линию замкнутой и обход
        'вершин идёт по кругу, как и положено
        Dim pointCount As Integer = count
        If Pline.Closed = True Then pointCount = count + 1
        Dim userCoords(pointCount - 1) As NetTopologySuite.Geometries.Coordinate
        For i As Integer = 0 To count - 1
            Dim pt As Point2d = Pline.GetPoint2dAt(i)
            userCoords(i) = New NetTopologySuite.Geometries.Coordinate(pt.X, pt.Y)
        Next i
        If Pline.Closed = True Then
            userCoords(count) = New NetTopologySuite.Geometries.Coordinate(userCoords(0).X, userCoords(0).Y)
        End If

        Dim userLine As NetTopologySuite.Geometries.LineString = Nothing
        Try
            userLine = New NetTopologySuite.Geometries.GeometryFactory().CreateLineString(userCoords)
        Catch ex As System.Exception
            Return False
        End Try
        If IsNothing(userLine) = True Then Return False

        'смещение отсчитываем от указанной точки; замкнутость NTS определит по IsClosed
        Dim userPoint As NetTopologySuite.Geometries.Coordinate = New NetTopologySuite.Geometries.Coordinate(insertPoint.X, insertPoint.Y)
        GetPointAttributePosition = CadastralTools.CadastralPoint.GetPositionAttributeByLine(userLine, Dist, insAttr, userPoint, True, False)
        Return GetPointAttributePosition
    End Function
    '===================================================================================================================
    'индекс ближайшей к указанной точке вершины полилинии
    '===================================================================================================================
    Private Shared Function GetNearestVertexIndex(ByVal Pline As Autodesk.AutoCAD.DatabaseServices.Polyline, ByVal pickedPoint As Point3d) As Integer
        GetNearestVertexIndex = -1
        If IsNothing(Pline) = True Then Return -1
        Dim count As Integer = Pline.NumberOfVertices
        If count < 1 Then Return -1

        Dim userPoint As Point2d = New Point2d(pickedPoint.X, pickedPoint.Y)
        Dim minDistance As Double = Double.MaxValue
        For i As Integer = 0 To count - 1
            Dim distance As Double = Pline.GetPoint2dAt(i).GetDistanceTo(userPoint)
            If distance < minDistance Then
                minDistance = distance
                GetNearestVertexIndex = i
            End If
        Next i
        Return GetNearestVertexIndex
    End Function
    '===================================================================================================================
    'ближайшая к указанному месту вершина полилинии — в неё вставляется блок
    '===================================================================================================================
    Private Shared Function GetNearestVertexPoint(ByVal Pline As Autodesk.AutoCAD.DatabaseServices.Polyline, ByVal pickedPoint As Point3d) As Point3d
        Dim numberVertex As Integer = GetNearestVertexIndex(Pline, pickedPoint)
        If numberVertex < 0 Then Return pickedPoint
        Dim pt As Point2d = Pline.GetPoint2dAt(numberVertex)
        Return New Point3d(pt.X, pt.Y, 0) 'Z=0: атрибут номера CreateBlock ставит в эту же плоскость
    End Function
    '===================================================================================================================
    'перенумерация полигонов (команда переименована: имя ReNumberCadastralObject дублировало RenumberCadastralObject из UpdateCadastralObject.vb и вызывала eDuplicateKey при загрузке)
    '===================================================================================================================
    <CommandMethod("ReNumberCadastralPolygon")>
    Public Sub ReNumberCadastralObject()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем настройки для диалогового окна УСТАНОВКИ
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        Dim formNumbersPoint As CadastralTools.FormCreatePointCadastral = New CadastralTools.FormCreatePointCadastral(MAnno)
        'вызываем диалоговое окно УСТАНОВКИ
        formNumbersPoint.ShowDialog()
        If formNumbersPoint.boolShow = False Then
            Exit Sub
        End If
        'путь к шаблону
        Dim templateDWG As String = formNumbersPoint.PutchTemlateDwg
        Dim templateXML As String = formNumbersPoint.PutchTemlateXML
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'настройки
        Dim idStyle As ObjectId = Nothing 'текстовые стиль
        'префикс номера
        Dim prefNumber As String = formNumbersPoint.prefNumber
        'метод определения координат
        Dim MetodCoord As String = formNumbersPoint.metodCoord
        'способ закрепления
        Dim Zakrep As String = formNumbersPoint.zakrPoint
        'погрешность определения точки
        Dim Delta As String = formNumbersPoint.icnarisityPoint
        'поиск дубликатов
        Dim boolDoublePoint As Boolean = formNumbersPoint.boolDublePoint
        Dim tolerance As Double = formNumbersPoint.tolerance
        'округрение вершин
        Dim boolRoundVertex As Boolean = formNumbersPoint.boolRoundVertex
        'по часовой стрелке
        Dim boolClock As Boolean = formNumbersPoint.boolClockVertex
        'северо-западный угол
        Dim boolSVAngle As Boolean = formNumbersPoint.boolSEVertex
        '========================================================================================
        Dim M As Double = formNumbersPoint.scalePoint * 1000 'масштаб
        Dim HText As Double = (M * formNumbersPoint.hText) / 1000 ' высота подписи
        Dim DText As Double = (M * formNumbersPoint.offsetText) / 1000 'расстояние смещения
        '==================================================================================================
        'размер блока
        Dim nameBlk As String = ""
        '===========================================================================================
        'характеристика атрибута
        Dim nameAttributeNumber As String = formNumbersPoint.nameAttribyteNumber
        Dim arrayAttributes As String(,) = formNumbersPoint.arrayAttribytes
        '============================================================================================
        Dim categoryZN As String = formNumbersPoint.categoryZN
        Dim tableZn As String = formNumbersPoint.tableZN
        Dim typeObject As CadastralTools.StructureElement.typeGeometryObject = StructureElement.typeGeometryObject.block
        Dim typeLineObject As CadastralTools.StructureElement.typeGeometryObject = formNumbersPoint.typeGeometry
        Dim descrittionBlockStyle As String = formNumbersPoint.deskBlockStyle
        Dim descriptionObjectStyle As String = formNumbersPoint.deskObjectStyle
        Dim NamePSTable As String = "CADASTRAL_POINT"
        Dim idTablePs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(templateXML, NamePSTable)
        '===============================================================================================
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        Dim arrayPropertiesBlock As String(,) = {}
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim styleBlockReference As CadastralBlockStyle = New CadastralBlockStyle()
        Dim acadNetTopogy As FuncNetTopologySuite = New FuncNetTopologySuite()
        Dim cadastralNetTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using locker
            Using tr
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'копируем блок
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                styleBlockReference.DBDrawing = db
                styleBlockReference.TemplateDWG = templateDWG
                styleBlockReference.TemplateXML = templateXML
                Dim boolReadPropertiesBlock As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZN, tableZn, typeObject, descrittionBlockStyle, arrayPropertiesBlock)
                If arrayPropertiesBlock.Length > 0 Then
                    styleBlockReference.createBlockStyle(arrayPropertiesBlock)
                End If
                styleBlockReference.ScaleBlock = (styleBlockReference.ScaleBlock * M) / 1000
                nameBlk = styleBlockReference.NameBlock
                If IsNothing(nameBlk) = True Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                If acBlkTbl.Has(nameBlk) = False Then
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Имя блока не найдено в чертеже. Загрузка из шаблона также не удалась. Попробуйте добавить его в ручную.")
                    Exit Sub
                End If
                '====================================================================================================
                'фильтуем точки чертежа в пределах видимости
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(db, nameBlk)
                Dim LargNumber As Integer = 1
                Dim oldPoint As New List(Of String)
                oldPoint = listBlockReference.Keys.ToList
                Dim newListPoint As New List(Of Point)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'начальный номер точки
                Dim pStrOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер точки: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowNone = True
                pStrOpts.AllowZero = False
                pStrOpts.DefaultValue = LargNumber
                Dim pStrRes As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOpts) 'запрашиваем номер
                If pStrRes.Status = PromptStatus.OK Then
                    Dim numberPoint As Integer = pStrRes.Value
                    Dim dictCadastralParcels As New Dictionary(Of Integer, String(,))
                    Try
                        'выбираем примитивы
                        Dim CountParcels As Integer = 0
                        Dim SelectParcel As PromptSelectionResult = ed.GetSelection()
                        If SelectParcel.Status = PromptStatus.OK Then 'если выбор состоялся тогда продолжаем
                            Dim acSSet1 As SelectionSet = SelectParcel.Value
                            For Each IDObj As ObjectId In acSSet1.GetObjectIds
                                Dim entPLine As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(IDObj, OpenMode.ForRead, False)
                                If TypeOf entPLine Is Polyline Then
                                    Dim deskObject As String = ""
                                    Dim dataPolyline As StructureElement = Nothing
                                    Dim boolReadData As Boolean = FuncXRecords.getXRecords(entPLine, dataPolyline)
                                    If boolReadData = True Then
                                        Dim PSObject As String = dataPolyline.TablePS
                                        If PSObject Like "CADASTRAL_OSTROV" Or PSObject Like "CADASTRAL_OBJECT" Then
                                            Continue For
                                        End If
                                        deskObject = dataPolyline.Description
                                        Dim keyParam As String = dataPolyline.KeyParameter
                                        If deskObject Like descriptionObjectStyle Then
                                            Dim categoryEarth As String = ""
                                            Dim ConditionNumber As String = CadastralTools.JsonValidator.GetValueFromJson(dataPolyline.KeyParameter, "ConditionalNumber")
                                            If Delta <= 0 Then
                                                categoryEarth = CadastralTools.JsonValidator.GetValueFromJson(dataPolyline.KeyParameter, "Category")
                                            End If
                                            '========================================================================================
                                            'читаем расширенные данные
                                            Dim LABEL As String = ""
                                            ReDim Preserve ArrayTablesXML(1, 0)
                                            ArrayTablesXML(0, 0) = dataPolyline.TablePS
                                            Dim ArrayData As String(,) = {}
                                            Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(entPLine, ArrayData, ArrayTablesXML)
                                            If ArrayData.Length > 0 Then
                                                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                                    Dim NTbl As String = ArrayData(0, i)
                                                    Dim NField As String = ArrayData(1, i)
                                                    Dim Zn As String = ArrayData(2, i)
                                                    If NField Like "LABEL" Then
                                                        deskObject = Zn
                                                    ElseIf NField Like "ConditionalNumber" Then
                                                        Dim tempConditionalNumber As String = Zn.Trim
                                                        If Not (tempConditionalNumber Like ConditionNumber) Then
                                                            keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", tempConditionalNumber)
                                                            dataPolyline.KeyParameter = keyParam
                                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entPLine, dataPolyline)
                                                        End If
                                                        ConditionNumber = Zn
                                                    ElseIf NField Like "Category" Then
                                                        Dim tempCategory As String = Zn.Trim
                                                        If Not (tempCategory Like categoryEarth) Then
                                                            keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "Category", tempCategory)
                                                            dataPolyline.KeyParameter = keyParam
                                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entPLine, dataPolyline)
                                                        End If
                                                        categoryEarth = Zn
                                                    End If
                                                Next i
                                            Else
                                                Continue For
                                            End If

                                            If IsNothing(ConditionNumber) = False Then
                                                If ConditionNumber.Trim.Length > 0 Then
                                                    Dim numberParcel As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(ConditionNumber)
                                                    '=====================================================================================================
                                                    If dictCadastralParcels.ContainsKey(numberParcel) = False Then
                                                        ReDim Preserve ArrayPoligon(3, 0)
                                                        ArrayPoligon(0, 0) = numberParcel 'номер участка
                                                        ArrayPoligon(1, 0) = entPLine.Handle.ToString 'хангл полилинии
                                                        ArrayPoligon(2, 0) = dataPolyline.TablePS
                                                        If formNumbersPoint.boolIcnarisity = True Then
                                                            Delta = formNumbersPoint.icnarisityPoint
                                                        Else
                                                            If categoryEarth.Trim.Length > 0 Then
                                                                Delta = CadastralTools.CadastralInfo.getDeltaPositionPointByCategoryEarth(categoryEarth)
                                                            Else
                                                                Delta = formNumbersPoint.icnarisityPoint
                                                            End If
                                                        End If
                                                        ArrayPoligon(3, 0) = Delta
                                                        dictCadastralParcels.Add(numberParcel, ArrayPoligon)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            tr.Abort()
                            locker.Dispose()
                            MsgBox("Прервано пользователем.")
                            Exit Sub
                        End If
                        '==================================================================================================
                        If dictCadastralParcels.Count = 0 Then
                            tr.Abort()
                            locker.Dispose()
                            MsgBox("Массив с участками пуст!!!")
                            Exit Sub
                        Else
                            If dictCadastralParcels.Count > 1 Then
                                dictCadastralParcels = dictCadastralParcels.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
                            End If
                        End If
                        '====================================================================================================
                        'выбираем кадастровые полигоны
                        Dim dictCadastralObject As New Dictionary(Of String, List(Of Geometries.Geometry))
                        If typeLineObject = StructureElement.typeGeometryObject.polygon Then
                            dictCadastralObject = SelectCadastralObject.getCadastralPolygon(db, descriptionObjectStyle)
                        ElseIf typeLineObject = StructureElement.typeGeometryObject.polyline Then
                            dictCadastralObject = SelectCadastralObject.getCadastralPolyline(db, descriptionObjectStyle)
                        End If
                        If dictCadastralObject.Count = 0 Then
                            MsgBox("Не удалось найти объекты на чертеже типа " & typeLineObject.ToString)
                            tr.Abort()
                            locker.Dispose()
                            Exit Sub
                        End If
                        '====================================================================================================
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'массив с участками
                        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                        For i As Integer = 0 To dictCadastralParcels.Count - 1
                            ArrayPoligon = dictCadastralParcels.ElementAt(i).Value
                            'получаем доступ к полилинии для извлечения координат
                            Dim numberParcel As Integer = Val(ArrayPoligon(0, 0))
                            Dim handlePLine As String = ArrayPoligon(1, 0)
                            Dim tablePS As String = ArrayPoligon(2, 0)
                            Delta = ArrayPoligon(3, 0)
                            Dim listParcels As List(Of NetTopologySuite.Geometries.Geometry) = dictCadastralObject.Item(descriptionObjectStyle)
                            'забираем все контура для выбранного номера участка
                            Dim dictCounters As Dictionary(Of Integer, NetTopologySuite.Geometries.Geometry) = Nothing
                            If typeLineObject = StructureElement.typeGeometryObject.polygon Then
                                dictCounters = CadastralTools.CadastralPolygon.getCountersByNumber(numberParcel, listParcels)
                            ElseIf typeLineObject = StructureElement.typeGeometryObject.polyline Then
                                dictCounters = CadastralTools.CadastralLine.getCountersByNumber(numberParcel, listParcels)
                            End If
                            If IsNothing(dictCounters) = False Then
                                If dictCounters.Count > 0 Then
                                    For j As Integer = 0 To dictCounters.Count - 1
                                        Dim tempCounter As NetTopologySuite.Geometries.Geometry = dictCounters.ElementAt(j).Value
                                        Dim triplets As New List(Of (VertexIndex As Integer, Prev As Coordinate, Current As Coordinate, [Next] As Coordinate))
                                        Dim dataCounter As StructureElement = Nothing
                                        If TypeOf tempCounter Is NetTopologySuite.Geometries.Polygon Then
                                            Dim counter As NetTopologySuite.Geometries.Polygon = tempCounter
                                            dataCounter = counter.UserData
                                            counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                            If boolSVAngle = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.NorthWestVertexFirst(counter)
                                            End If
                                            If boolClock = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.ReversePolygonPreserveOrientation(counter)
                                            End If
                                            If boolRoundVertex = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.RoundPolygonVertices(counter, tolerance)
                                            End If
                                            Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromPolygon(counter)
                                            triplets = CadastralTools.CadastralPoint.GetPolygonVertexTriplets(counter)
                                        ElseIf TypeOf tempCounter Is NetTopologySuite.Geometries.LineString Then
                                            Dim counter As NetTopologySuite.Geometries.LineString = tempCounter
                                            dataCounter = counter.UserData
                                            counter = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(counter)
                                            If boolRoundVertex = True Then
                                                counter = CadastralTools.FuncNetTopologySuite.RoundLineStringVertices(counter, tolerance)
                                            End If
                                            Dim boolUpdateVertexPolyline As Boolean = classDrawObject.UpdatePolylineFromLineString(counter)
                                            triplets = CadastralTools.CadastralPoint.GetLineStringVertexTriplets(counter)
                                        End If
                                        If triplets.Count > 0 Then
                                            For Each t In triplets
                                                Dim insAttr As Double() = {}
                                                Dim Vert0 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Prev.X, t.Prev.Y)
                                                Dim Vert1 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Current.X, t.Current.Y)
                                                Dim Vert2 As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(t.Next.X, t.Next.Y)
                                                If newListPoint.Contains(Vert1) Then
                                                    Continue For
                                                End If
                                                'ищем все точки в узле
                                                Dim findAttributeVertex As Boolean = False
                                                Dim dictPointInVertex As Dictionary(Of String, NetTopologySuite.Geometries.Point) = CadastralTools.CadastralPoint.getPointsByCoordinates(Vert1, listBlockReference)
                                                If dictPointInVertex.Count > 0 Then
                                                    For k As Integer = 0 To dictPointInVertex.Count - 1
                                                        Dim userPoint As NetTopologySuite.Geometries.Point = dictPointInVertex.ElementAt(k).Value
                                                        If IsNothing(userPoint.UserData) = False Then
                                                            If TypeOf userPoint.UserData Is Dictionary(Of String, String) Then
                                                                Dim dictDataPoint As Dictionary(Of String, String) = userPoint.UserData
                                                                If dictDataPoint.ContainsKey("IDRNG") = True Then
                                                                    Dim idStructures As String = dictDataPoint.Item("IDRNG")
                                                                    If boolDoublePoint = True Then
                                                                        idStructures = dataCounter.IdElement
                                                                    End If
                                                                    If idStructures Like dataCounter.IdElement Then
                                                                        'перезаписываем характеристики атрибута
                                                                        Dim dictAttributes As New Dictionary(Of String, String)
                                                                        dictAttributes.Add(nameAttributeNumber, prefNumber & numberPoint)
                                                                        dictAttributes.Add("DELTA", Delta)
                                                                        dictAttributes.Add("FASTDESK", Zakrep)
                                                                        dictAttributes.Add("METHODCOORDINATES", MetodCoord)
                                                                        Dim handleBlk As String = dictPointInVertex.ElementAt(k).Key
                                                                        Dim idBlk As ObjectId = AcadFunction.FuncReturnObjectID(handleBlk)
                                                                        Dim updateBlk As BlockReference = tr.GetObject(idBlk, OpenMode.ForWrite)
                                                                        Dim updateAttrib As Boolean = AttributeCadastralBlock.UpdateAttributesInBlockReference(db, updateBlk, dictAttributes)
                                                                        findAttributeVertex = True
                                                                        If oldPoint.Count > 0 Then
                                                                            oldPoint.Remove(handleBlk)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Next k
                                                End If
                                                'вставляем новую точку
                                                If findAttributeVertex = False Then
                                                    Dim boolInsAttr As Boolean = CadastralTools.CadastralPoint.calculatePositionAttributeBlock(Vert0, Vert1, Vert2, insAttr, DText, True)
                                                    'заполняем атрибуты
                                                    Dim listAttributes As List(Of AttributeCadastralBlock) = New List(Of AttributeCadastralBlock)
                                                    For k As Integer = 1 To arrayAttributes.GetUpperBound(0)
                                                        Dim attribyteNumberPoint As AttributeCadastralBlock = New AttributeCadastralBlock()
                                                        'номер
                                                        attribyteNumberPoint.TagAttribute = arrayAttributes(k, 0)
                                                        attribyteNumberPoint.DefinitionAttribute = arrayAttributes(k, 1)
                                                        If attribyteNumberPoint.TagAttribute Like nameAttributeNumber Then
                                                            attribyteNumberPoint.TextStringAttribute = prefNumber & numberPoint
                                                            attribyteNumberPoint.HeightAttribute = HText
                                                            attribyteNumberPoint.XAttribute = insAttr(0)
                                                            attribyteNumberPoint.YAttribute = insAttr(1)
                                                            attribyteNumberPoint.HorizontalAlignment = insAttr(2)
                                                            attribyteNumberPoint.VerticalAlignment = insAttr(3)
                                                        ElseIf attribyteNumberPoint.TagAttribute Like "DELTA" Then
                                                            attribyteNumberPoint.TextStringAttribute = Delta
                                                        ElseIf attribyteNumberPoint.TagAttribute Like "FASTDESK" Then
                                                            attribyteNumberPoint.TextStringAttribute = Zakrep
                                                        ElseIf attribyteNumberPoint.TagAttribute Like "METHODCOORDINATES" Then
                                                            attribyteNumberPoint.TextStringAttribute = MetodCoord
                                                        End If
                                                        listAttributes.Add(attribyteNumberPoint)
                                                    Next k
                                                    'вставляем блок
                                                    Dim insertPointBlk As Point3d = New Point3d(t.Current.X, t.Current.Y, 0)
                                                    Dim insertBlk As BlockReference = CreateDWGObject.CreateBlock(insertPointBlk, styleBlockReference, listAttributes)
                                                    If IsNothing(insertBlk) = False Then
                                                        Dim dataPoint As CadastralTools.StructureElement = CadastralTools.CadastralPoint.createDataPoint(categoryZN, tableZn, descrittionBlockStyle, dataCounter.IdElement, numberPoint, prefNumber, Delta, Zakrep, MetodCoord, insertBlk.Handle.ToString)
                                                        Dim dictData As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(dataPoint.KeyParameter)
                                                        Dim recData As Boolean = FuncXRecords.writeXRecords(insertBlk, dataPoint)
                                                        If IsNothing(idTablePs) = False Then
                                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(insertBlk, idTablePs)
                                                            If boolADDPSTBL = True Then
                                                                Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(insertBlk, dataPoint)
                                                            End If
                                                        End If
                                                    End If
                                                    If boolDoublePoint = True Then
                                                        Dim newDictPoint As New Dictionary(Of String, String)
                                                        newDictPoint.Add(nameAttributeNumber, numberPoint)
                                                        Vert1.UserData = newDictPoint
                                                        listBlockReference.Add(insertBlk.Handle.ToString, Vert1)
                                                    End If
                                                End If
                                                newListPoint.Add(Vert1)
                                                numberPoint += 1
                                            Next
                                        End If
                                    Next j
                                End If
                            End If
                        Next i
                    Catch ex As System.Exception
                        MsgBox("Ошибка!!! " & ex.ToString)
                        tr.Abort()
                        locker.Dispose()
                    End Try
                Else
                    tr.Abort()
                    locker.Dispose()
                    MsgBox("Прервано пользователем.")
                    Exit Sub
                End If
                If oldPoint.Count > 0 Then
                    Dim mgBox As MsgBoxResult = MsgBox("Удалить лишние точки?", MsgBoxStyle.OkCancel)
                    If mgBox = MsgBoxResult.Ok Then
                        For i As Integer = 0 To oldPoint.Count - 1
                            Dim hgBlk As String = oldPoint.Item(i)
                            If hgBlk.Trim.Length > 1 Then
                                Dim idPoint As ObjectId = AcadFunction.FuncReturnObjectID(hgBlk)
                                Dim entPoint As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idPoint, OpenMode.ForWrite)
                                entPoint.Erase()  ' <-- Этого не хватало для удаления
                            End If
                        Next i
                    End If
                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    '===================================================================================================================
    'ОБЩИЕ ПОМОЩНИКИ КОМАНД РАБОТЫ С КАДАСТРОВЫМИ ТОЧКАМИ
    '===================================================================================================================
    'сохраняем активный документ
    Private Shared Sub SaveActiveDocument(ByVal acDoc As Document)
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
    End Sub
    '===================================================================================================================
    'выбор блока-образца; Nothing — пользователь отказался или выбран не блок
    '===================================================================================================================
    Private Shared Function PickSampleBlock(ByVal tr As Transaction, ByVal ed As Editor) As BlockReference
        Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите один блок в качестве образца:")
        If per.Status <> PromptStatus.OK Then
            ed.WriteMessage(vbLf & "Блок не выбран.")
            Return Nothing
        End If
        Dim blk As BlockReference = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead, False), BlockReference)
        If IsNothing(blk) = True Then ed.WriteMessage(vbLf & "Выбранный объект не является блоком.")
        Return blk
    End Function
    '===================================================================================================================
    'имена видимых атрибутов блока (без повторов)
    '===================================================================================================================
    Private Shared Function GetVisibleAttributeTags(ByVal tr As Transaction, ByVal blk As BlockReference) As List(Of String)
        Dim result As New List(Of String)
        If IsNothing(blk) = True Then Return result
        Dim attCol As AttributeCollection = blk.AttributeCollection
        For Each attId As ObjectId In attCol
            Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
            If IsNothing(attRef) = True Then Continue For
            If attRef.Visible = False Then Continue For
            Dim tagName As String = attRef.Tag.Trim
            If tagName.Length = 0 Then Continue For
            If result.Contains(tagName) = False Then result.Add(tagName)
        Next
        Return result
    End Function
    '===================================================================================================================
    'выбор атрибута из списка видимых атрибутов образца;
    'если атрибут один — диалог не нужен, пустая строка — прервано
    '===================================================================================================================
    Private Shared Function GetAttributeName(ByVal ed As Editor, ByVal userKeywords As List(Of String), Optional ByVal message As String = "Выберите атрибут:") As String
        If userKeywords Is Nothing OrElse userKeywords.Count = 0 Then Return ""
        If userKeywords.Count = 1 Then
            ed.WriteMessage(vbLf & "Атрибут для работы: " & userKeywords(0))
            Return userKeywords(0)
        End If
        'имя атрибута становится ключевым словом, поэтому пробелы в нём недопустимы
        Dim listKeyword As New List(Of String)
        For i As Integer = 0 To userKeywords.Count - 1
            If userKeywords(i).Contains(" ") = True Then Continue For
            listKeyword.Add(userKeywords(i))
        Next
        If listKeyword.Count = 0 Then Return ""
        If listKeyword.Count = 1 Then
            ed.WriteMessage(vbLf & "Атрибут для работы: " & listKeyword(0))
            Return listKeyword(0)
        End If
        Return GetKeywordFromList(ed, message, listKeyword.ToArray(), listKeyword(0))
    End Function
    '===================================================================================================================
    'отбор блоков по образцу: имя блока берём у образца, значения атрибутов — в UserData точки
    '===================================================================================================================
    Private Shared Function GetSamplePointDictionary(ByVal db As Database, ByVal blkSample As BlockReference) As Dictionary(Of String, NetTopologySuite.Geometries.Point)
        Dim result As New Dictionary(Of String, NetTopologySuite.Geometries.Point)
        If IsNothing(blkSample) = True Then Return result
        Return SelectCadastralObject.getPointByListBlockReference(db, blkSample.Name)
    End Function
    '===================================================================================================================
    'ObjectId по хендлу; ObjectId.Null — если хендл пустой или объект не найден
    '===================================================================================================================
    Private Shared Function GetObjectIdByHandle(ByVal handle As String) As ObjectId
        If IsNothing(handle) = True Then Return ObjectId.Null
        If handle.Trim.Length = 0 Then Return ObjectId.Null
        If handle Like "0" Then Return ObjectId.Null
        Dim idPoint As ObjectId = AcadFunction.FuncReturnObjectID(handle)
        If idPoint.IsNull = True Then Return ObjectId.Null
        Return idPoint
    End Function
    '===================================================================================================================
    'хендл точки из списка по координатам; пустая строка — точки в узле нет
    '===================================================================================================================
    Private Shared Function GetHandleByPoint(ByVal listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point), ByVal pt As Point2d, ByVal tolerance As Double) As String
        If IsNothing(listBlockReference) = True Then Return ""
        If listBlockReference.Count = 0 Then Return ""
        Dim userPoint As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(pt.X, pt.Y)
        Dim handlePoint As String = CadastralTools.CadastralPoint.getIDPointByCoordinates(userPoint, listBlockReference, "", tolerance)
        If IsNothing(handlePoint) = True Then Return ""
        If handlePoint Like "0" Then Return ""
        Return handlePoint
    End Function
    '===================================================================================================================
    'идентификаторы всех точек из списка
    '===================================================================================================================
    Private Shared Function GetIdsByDictionary(ByVal listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point)) As ObjectIdCollection
        Dim result As ObjectIdCollection = New ObjectIdCollection()
        If IsNothing(listBlockReference) = True Then Return result
        For Each kvp As KeyValuePair(Of String, NetTopologySuite.Geometries.Point) In listBlockReference
            Dim idPoint As ObjectId = GetObjectIdByHandle(kvp.Key)
            If idPoint.IsNull = False Then
                If result.Contains(idPoint) = False Then result.Add(idPoint)
            End If
        Next
        Return result
    End Function
    '===================================================================================================================
    'значение выбранного атрибута точки из списка; пустая строка — атрибута нет
    '===================================================================================================================
    Private Shared Function GetAttributeValue(ByVal userPoint As NetTopologySuite.Geometries.Point, ByVal nameAttribute As String) As String
        If IsNothing(userPoint) = True Then Return ""
        If TypeOf userPoint.UserData IsNot Dictionary(Of String, String) Then Return ""
        Dim dictAttributes As Dictionary(Of String, String) = userPoint.UserData
        If dictAttributes.ContainsKey(nameAttribute) = False Then Return ""
        Return dictAttributes.Item(nameAttribute)
    End Function
    '===================================================================================================================
    'подходит ли точка под статус: Новые — есть "н", Существующие — только цифры,
    'сТарые — есть "%"; "Все" и пустая строка — подходит любая точка
    '===================================================================================================================
    Private Shared Function IsPointByStatus(ByVal strNumber As String, ByVal status As String) As Boolean
        If IsNothing(strNumber) = True Then strNumber = ""
        Dim tempStr As String = strNumber.Trim
        If status.Trim.Length = 0 Then Return True
        If status Like "Все" Then Return True
        If status Like "Новые" Then
            If tempStr.IndexOf("н") <> -1 OrElse tempStr.IndexOf("Н") <> -1 Then Return True
            Return False
        End If
        If status Like "Существующие" Then Return IsNumeric(tempStr)
        If status Like "сТарые" Then
            If tempStr.IndexOf("%") <> -1 Then Return True
            Return False
        End If
        Return False
    End Function
    '===================================================================================================================
    'удаление точки; False — объект не найден или уже удалён
    '===================================================================================================================
    Private Shared Function EraseGeoPoint(ByVal tr As Transaction, ByVal idPoint As ObjectId) As Boolean
        If idPoint.IsNull = True Then Return False
        Dim entPoint As Entity = TryCast(tr.GetObject(idPoint, OpenMode.ForWrite, False), Entity)
        If IsNothing(entPoint) = True Then Return False
        If entPoint.IsErased = True Then Return False
        entPoint.Erase()
        Return True
    End Function
    '===================================================================================================================
    'показать блок точки: подсветить и наехать видом
    '===================================================================================================================
    Private Shared Sub ShowPointBlock(ByVal ed As Editor, ByVal blk As BlockReference)
        If IsNothing(blk) = True Then Return
        ed.SetImpliedSelection(New ObjectId() {blk.ObjectId})
        Dim centrePoint As Point3d = blk.Position
        SubZoom.Zoom(New Point3d(centrePoint.X - 10, centrePoint.Y - 10, 0),
                     New Point3d(centrePoint.X + 10, centrePoint.Y + 10, 0), New Point3d(), 1)
    End Sub
    '===================================================================================================================
    'выбор режима обработки по ключевому слову; пустая строка — пользователь прервал команду
    '===================================================================================================================
    Private Shared Function GetKeywordMode(ByVal ed As Editor, ByVal message As String, ByVal userKeywords As String(), ByVal defaultKeyword As String) As String
        Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pIntOpts.Message = vbCrLf & message
        For i As Integer = 0 To userKeywords.Length - 1
            pIntOpts.Keywords.Add(userKeywords(i))
        Next i
        pIntOpts.Keywords.Default = defaultKeyword
        pIntOpts.AllowNone = True
        Dim pIntRes As PromptResult = ed.GetKeywords(pIntOpts)
        If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then Return ""
        Return pIntRes.StringResult
    End Function
    '===================================================================================================================
    'имя блока-образца, выбранного пользователем
    '===================================================================================================================
    Private Shared Function GetPickedBlockName(ByVal tr As Transaction, ByVal ed As Editor) As String
        Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите один блок, для образца:")
        If per.Status <> PromptStatus.OK Then Return ""
        Dim blk As BlockReference = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead, False), BlockReference)
        If IsNothing(blk) = True Then Return ""
        Return blk.Name
    End Function
    '===================================================================================================================
    'точки, стоящие в узлах полилинии (диапазон вершин включительно); без диапазона — все вершины
    'Точки ищутся в отобранном по образцу списке по координатам узлов
    '===================================================================================================================
    Private Shared Function GetGeoPointIdsInPolyline(ByVal Pline As Autodesk.AutoCAD.DatabaseServices.Polyline,
                                                     ByVal listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point),
                                                     Optional ByVal numberVertexStart As Integer = -1,
                                                     Optional ByVal numberVertexEnd As Integer = -1,
                                                     Optional ByVal tolerance As Double = 0.01) As ObjectIdCollection
        Dim result As ObjectIdCollection = New ObjectIdCollection()
        If IsNothing(Pline) = True Then Return result
        If IsNothing(listBlockReference) = True Then Return result
        If listBlockReference.Count = 0 Then Return result
        Dim count As Integer = Pline.NumberOfVertices
        If count < 1 Then Return result
        If numberVertexStart < 0 Then numberVertexStart = 0
        If numberVertexEnd < 0 OrElse numberVertexEnd > count - 1 Then numberVertexEnd = count - 1
        If numberVertexStart > numberVertexEnd Then Return result

        For i As Integer = numberVertexStart To numberVertexEnd
            Dim handlePoint As String = GetHandleByPoint(listBlockReference, Pline.GetPoint2dAt(i), tolerance)
            If handlePoint.Trim.Length = 0 Then Continue For
            Dim idPoint As ObjectId = GetObjectIdByHandle(handlePoint)
            If idPoint.IsNull = True Then Continue For
            If result.Contains(idPoint) = False Then result.Add(idPoint)
        Next i
        Return result
    End Function
    '===================================================================================================================
    'выбор диапазона вершин на одной полилинии (двумя щелчками)
    '===================================================================================================================
    Private Shared Function PickVertexRange(ByVal tr As Transaction,
                                            ByVal ed As Editor,
                                            ByRef Pline As Autodesk.AutoCAD.DatabaseServices.Polyline,
                                            ByRef numberVertexStart As Integer,
                                            ByRef numberVertexEnd As Integer) As Boolean
        PickVertexRange = False
        Pline = Nothing
        numberVertexStart = -1
        numberVertexEnd = -1

        Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите первую точку ближайшую к узлу полилинии:")
        If per.Status = PromptStatus.OK Then
            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead, False), Autodesk.AutoCAD.DatabaseServices.Polyline)
            If IsNothing(entObj) = False Then
                Pline = entObj
                numberVertexStart = GetNearestVertexIndex(Pline, per.PickedPoint)
            End If
        End If

        per = ed.GetEntity(vbLf & "Выберите последнюю точку ближайшую к узлу полилинии:")
        If per.Status = PromptStatus.OK Then
            Dim entObj As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead, False), Autodesk.AutoCAD.DatabaseServices.Polyline)
            If IsNothing(entObj) = False Then
                If IsNothing(Pline) = False Then
                    If Pline.Id <> per.ObjectId Then
                        MsgBox("Обработать точки можно только если они находятся на одной и той же полилинии. Макро прервано!!!")
                        Return False
                    End If
                Else
                    Pline = entObj
                End If
                numberVertexEnd = GetNearestVertexIndex(Pline, per.PickedPoint)
            End If
        End If

        If IsNothing(Pline) = False Then
            If numberVertexStart = -1 Then numberVertexStart = 0
            If numberVertexEnd = -1 Then numberVertexEnd = Pline.NumberOfVertices - 1
            'если конец указан раньше начала — меняем местами
            If numberVertexStart > numberVertexEnd Then
                Dim tempVertex As Integer = numberVertexStart
                numberVertexStart = numberVertexEnd
                numberVertexEnd = tempVertex
            End If
            PickVertexRange = True
        End If
        Return PickVertexRange
    End Function
    '===================================================================================================================
    'сбор точек по режиму обработки:
    '  Все      — все отобранные по образцу точки,
    '  Полигон  — точки в узлах выбранных полилиний,
    '  Диапазон — точки в узлах диапазона вершин одной полилинии
    '===================================================================================================================
    Private Shared Function GetTargetGeoPointIds(ByVal tr As Transaction,
                                                 ByVal ed As Editor,
                                                 ByVal mode As String,
                                                 ByVal listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point)) As ObjectIdCollection
        Dim result As ObjectIdCollection = New ObjectIdCollection()
        Select Case mode
            Case "Все"
                result = GetIdsByDictionary(listBlockReference)
            Case "Полигон"
                Dim acSSPrompt As PromptSelectionResult = ed.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    For Each IDObj As ObjectId In acSSPrompt.Value.GetObjectIds
                        Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(tr.GetObject(IDObj, OpenMode.ForRead, False), Autodesk.AutoCAD.DatabaseServices.Polyline)
                        If IsNothing(Pline) = False Then
                            Dim idsPoint As ObjectIdCollection = GetGeoPointIdsInPolyline(Pline, listBlockReference)
                            For Each idPoint As ObjectId In idsPoint
                                If result.Contains(idPoint) = False Then result.Add(idPoint)
                            Next
                        End If
                    Next
                End If
            Case "Диапазон"
                Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                Dim numberVertexStart As Integer = -1
                Dim numberVertexEnd As Integer = -1
                If PickVertexRange(tr, ed, Pline, numberVertexStart, numberVertexEnd) = True Then
                    result = GetGeoPointIdsInPolyline(Pline, listBlockReference, numberVertexStart, numberVertexEnd)
                End If
        End Select
        Return result
    End Function
    '===================================================================================================================
    'выбор одного значения из списка (например, имени атрибута); пустая строка — прервано
    '===================================================================================================================
    Private Shared Function GetKeywordFromList(ByVal ed As Editor, ByVal message As String, ByVal userKeywords As String(), ByVal defaultKeyword As String) As String
        GetKeywordFromList = ""
        If userKeywords Is Nothing OrElse userKeywords.Length = 0 Then Return ""
        Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pIntOpts.Message = vbCrLf & message
        For i As Integer = 0 To userKeywords.Length - 1
            pIntOpts.Keywords.Add(userKeywords(i))
        Next i
        pIntOpts.Keywords.Default = defaultKeyword
        pIntOpts.AllowNone = False
        Dim pIntRes As PromptResult = ed.GetKeywords(pIntOpts)
        If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then Return ""
        If pIntRes.Status = PromptStatus.None Then Return defaultKeyword
        If pIntRes.StringResult.Trim.Length = 0 Then Return defaultKeyword
        Return pIntRes.StringResult
    End Function
    '===================================================================================================================
    'найти точку с наибольшим номером
    'Пользователь выбирает блок-образец, затем атрибут, из которого читается номер,
    'после чего по чертежу отбираются блоки того же имени и показывается блок
    'с наибольшим номером
    '===================================================================================================================
    <CommandMethod("FindGeopointToLargNumber")>
    Public Sub FindGeopointToLargNumber()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите один блок в качестве образца:")
                If per.Status <> PromptStatus.OK Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Блок не выбран.")
                    Exit Sub
                End If
                Dim blkSample As BlockReference = TryCast(tr.GetObject(per.ObjectId, OpenMode.ForRead, False), BlockReference)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Выбранный объект не является блоком.")
                    Exit Sub
                End If

                '2. видимые атрибуты образца — из них собираем список для выбора
                Dim userKeywords As New List(Of String)
                Dim countSkipTag As Integer = 0
                Dim attCol As AttributeCollection = blkSample.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                    If IsNothing(attRef) = True Then Continue For
                    If attRef.Visible = False Then Continue For
                    Dim tagName As String = attRef.Tag.Trim
                    If tagName.Length = 0 Then Continue For
                    'имя атрибута становится ключевым словом, поэтому пробелы в нём недопустимы
                    If tagName.Contains(" ") = True Then
                        countSkipTag += 1
                        Continue For
                    End If
                    If userKeywords.Contains(tagName) = False Then userKeywords.Add(tagName)
                Next
                If userKeywords.Count = 0 Then
                    tr.Commit()
                    MsgBox("У выбранного блока нет видимых атрибутов, пригодных для выбора.")
                    Exit Sub
                End If
                If countSkipTag > 0 Then
                    ed.WriteMessage(vbLf & "Атрибутов с пробелом в имени пропущено: " & countSkipTag)
                End If

                '3. выбор атрибута, из которого читаем номер
                Dim nameAttribute As String = GetKeywordFromList(ed, "Выберите атрибут, из которого читать номер:", userKeywords.ToArray(), userKeywords(0))
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '4. отбираем блоки того же имени (в словаре - точка с атрибутами в UserData)
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = SelectCadastralObject.getPointByListBlockReference(db, blkSample.Name)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '5. ищем блок с наибольшим номером
                Dim handleRez As String = ""
                Dim strNumberRez As String = ""
                Dim largNumber As Integer = Integer.MinValue
                For Each kvp As KeyValuePair(Of String, NetTopologySuite.Geometries.Point) In listBlockReference
                    Dim userPoint As NetTopologySuite.Geometries.Point = kvp.Value
                    If IsNothing(userPoint) = True Then Continue For
                    If TypeOf userPoint.UserData IsNot Dictionary(Of String, String) Then Continue For
                    Dim dictAttributes As Dictionary(Of String, String) = userPoint.UserData
                    If dictAttributes.ContainsKey(nameAttribute) = False Then Continue For
                    Dim strNumber As String = dictAttributes.Item(nameAttribute)
                    Dim numberPoint As Integer = CInt(CadastralTools.MathFunction.ExtractNumber(strNumber))
                    If numberPoint > largNumber Then
                        largNumber = numberPoint
                        strNumberRez = strNumber
                        handleRez = kvp.Key
                    End If
                Next
                If handleRez.Trim.Length = 0 Then
                    tr.Commit()
                    MsgBox("Не удалось прочитать номер из атрибута " & nameAttribute & " ни из одного блока.")
                    Exit Sub
                End If

                'показываем найденный блок
                Dim idRez As ObjectId = AcadFunction.FuncReturnObjectID(handleRez)
                Dim blkRez As BlockReference = Nothing
                If idRez.IsNull = False Then blkRez = TryCast(tr.GetObject(idRez, OpenMode.ForRead, False), BlockReference)
                If IsNothing(blkRez) = False Then
                    ed.SetImpliedSelection(New ObjectId() {idRez})
                    Dim centrePoint As Point3d = blkRez.Position
                    Dim minPoint As Point3d = New Point3d(centrePoint.X - 10, centrePoint.Y - 10, 0)
                    Dim maxPoint As Point3d = New Point3d(centrePoint.X + 10, centrePoint.Y + 10, 0)
                    SubZoom.Zoom(minPoint, maxPoint, New Point3d(), 1)
                    ed.WriteMessage(vbLf & "Наибольший номер: " & strNumberRez & " (атрибут " & nameAttribute & ", хендл " & handleRez & ").")
                Else
                    MsgBox("Блок с наибольшим номером не найден в чертеже.")
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось найти точку с наибольшим номером.")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
    End Sub
    '===================================================================================================================
    'удалить кадастровые точки
    'Образец блока задаёт имя блока, выбранный атрибут — по нему определяется номер и статус точки
    '===================================================================================================================
    <CommandMethod("DeleteCadastralPoint")>
    Public Sub DeleteCadastralPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim countDeletePoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If

                '2. атрибут, из которого читается номер точки
                Dim nameAttribute As String = GetAttributeName(ed, GetVisibleAttributeTags(tr, blkSample), "Выберите атрибут, из которого читать номер:")
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '3. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '4. режим удаления
                Dim mode As String = GetKeywordMode(ed, "Удалить точки?", New String() {"Все", "Новые", "Существующие", "сТарые", "Полигон", "Диапазон"}, "Полигон")
                If mode.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Прервано пользователем.")
                    Exit Sub
                End If

                '5. удаляем
                If mode Like "Полигон" OrElse mode Like "Диапазон" Then
                    'точки в узлах выбранных полилиний или диапазона вершин
                    Dim pointIDColl As ObjectIdCollection = GetTargetGeoPointIds(tr, ed, mode, listBlockReference)
                    For Each idPoint As ObjectId In pointIDColl
                        If EraseGeoPoint(tr, idPoint) = True Then countDeletePoint += 1
                    Next
                Else
                    'по статусу, который читается из выбранного атрибута
                    For Each kvp As KeyValuePair(Of String, NetTopologySuite.Geometries.Point) In listBlockReference
                        Dim strNumber As String = GetAttributeValue(kvp.Value, nameAttribute)
                        If IsPointByStatus(strNumber, mode) = False Then Continue For
                        Dim idPoint As ObjectId = GetObjectIdByHandle(kvp.Key)
                        If EraseGeoPoint(tr, idPoint) = True Then countDeletePoint += 1
                    Next
                End If
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось удалить все точки.")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        MsgBox("Всего удалено " & countDeletePoint & " точек.")
    End Sub
    '===================================================================================================================
    'найти точку по номеру
    'Образец блока задаёт имя блока, выбранный атрибут — из него читается номер
    '===================================================================================================================
    <CommandMethod("FindCadactralPointByNumber")>
    Public Sub FindCadactralPointByNumber()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim boolFindNumber As Boolean = False
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If

                '2. атрибут, из которого читается номер
                Dim nameAttribute As String = GetAttributeName(ed, GetVisibleAttributeTags(tr, blkSample), "Выберите атрибут, из которого читать номер:")
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '3. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '4. номер для поиска
                Dim pIntOpts As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите номер точки: ")
                pIntOpts.AllowNegative = False
                pIntOpts.AllowZero = False
                pIntOpts.AllowNone = False
                pIntOpts.DefaultValue = 1
                Dim pIntRes As PromptIntegerResult = ed.GetInteger(pIntOpts)
                If pIntRes.Status <> PromptStatus.OK Then
                    tr.Commit()
                    Exit Sub
                End If
                Dim numFindPoint As Integer = pIntRes.Value

                '5. ищем точки с таким номером и показываем их
                For Each kvp As KeyValuePair(Of String, NetTopologySuite.Geometries.Point) In listBlockReference
                    Dim strNumber As String = GetAttributeValue(kvp.Value, nameAttribute)
                    If CInt(CadastralTools.MathFunction.ExtractNumber(strNumber)) <> numFindPoint Then Continue For
                    Dim idPoint As ObjectId = GetObjectIdByHandle(kvp.Key)
                    If idPoint.IsNull = True Then Continue For
                    Dim blk As BlockReference = TryCast(tr.GetObject(idPoint, OpenMode.ForRead, False), BlockReference)
                    If IsNothing(blk) = True Then Continue For
                    ShowPointBlock(ed, blk)
                    boolFindNumber = True
                    If GetKeywordMode(ed, "Искать еще точки с таким номером?", New String() {"Да", "Нет"}, "Да") Like "нет" Then Exit For
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Ошибка. Макро прервано")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        If boolFindNumber = False Then MsgBox("Точки с таким номером нет на чертеже!!!")
    End Sub
    '===================================================================================================================
    'изменить масштаб кадастровых точек
    'Образец блока задаёт имя блока; выбранный атрибут должен быть заполнен у точки
    '===================================================================================================================
    <CommandMethod("ScaleCadastralPoint")>
    Public Sub ScaleCadastralPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim CountPoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If

                '2. атрибут, по которому отбираются точки
                Dim nameAttribute As String = GetAttributeName(ed, GetVisibleAttributeTags(tr, blkSample), "Выберите атрибут, по которому отбирать точки:")
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '3. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '4. режим обработки
                Dim mode As String = GetKeywordMode(ed, "Изменить масштаб точкам?", New String() {"Все", "Полигон", "Диапазон"}, "Полигон")
                If mode.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Прервано пользователем.")
                    Exit Sub
                End If

                '5. новый масштаб
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите новый масштаб блока (точки): ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                pStrOpts.DefaultValue = 1
                Dim pStrRes As PromptDoubleResult = ed.GetDouble(pStrOpts)
                If pStrRes.Status <> PromptStatus.OK Then
                    tr.Commit()
                    Exit Sub
                End If
                If pStrRes.Value <= 0 Then
                    tr.Commit()
                    Exit Sub
                End If
                Dim MBlk As Double = pStrRes.Value

                '6. масштабируем отобранные точки
                Dim pointIDColl As ObjectIdCollection = GetTargetGeoPointIds(tr, ed, mode, listBlockReference)
                For Each idBlk As ObjectId In pointIDColl
                    'точку берём, только если выбранный атрибут заполнен
                    Dim handlePoint As String = idBlk.Handle.ToString
                    If listBlockReference.ContainsKey(handlePoint) = False Then Continue For
                    Dim strValue As String = GetAttributeValue(listBlockReference.Item(handlePoint), nameAttribute)
                    If strValue.Trim.Length = 0 Then Continue For
                    Dim blk As BlockReference = TryCast(tr.GetObject(idBlk, OpenMode.ForWrite), BlockReference)
                    If IsNothing(blk) = True Then Continue For
                    If blk.ScaleFactors.X = 0 Then Continue For
                    Dim newScale As Double = MBlk / blk.ScaleFactors.X
                    blk.TransformBy(Matrix3d.Scaling(newScale, blk.Position))
                    CountPoint += 1
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить масштаб точкам! Макро прервано")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        MsgBox("Всего преобразовано " & CountPoint & " точек.")
    End Sub
    '===================================================================================================================
    'изменить высоту одного видимого атрибута блока
    '===================================================================================================================
    Private Shared Function SetAttributeHeight(ByVal tr As Transaction, ByVal idBlk As ObjectId, ByVal nameAttribute As String, ByVal heightBlk As Double) As Boolean
        SetAttributeHeight = False
        Dim blk As BlockReference = TryCast(tr.GetObject(idBlk, OpenMode.ForWrite), BlockReference)
        If IsNothing(blk) = True Then Return False
        Dim attCol As AttributeCollection = blk.AttributeCollection
        For Each attId As ObjectId In attCol
            Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
            If IsNothing(attRef) = True Then Continue For
            If attRef.Visible = False Then Continue For
            If attRef.Tag Like nameAttribute = False Then Continue For
            attRef.Height = heightBlk
            SetAttributeHeight = True
        Next
        Return SetAttributeHeight
    End Function
    '===================================================================================================================
    'изменить размер атрибута кадастровых точек
    'Образец блока задаёт имя блока, выбранный атрибут — тот, у которого меняется размер
    '===================================================================================================================
    <CommandMethod("heightAttributeCadastralPoint")>
    Public Sub HeightAttributeCadastralPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim CountPoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If

                '2. атрибут, у которого меняется размер
                Dim nameAttribute As String = GetAttributeName(ed, GetVisibleAttributeTags(tr, blkSample), "Выберите атрибут, размер которого изменить:")
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '3. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '4. режим обработки
                Dim mode As String = GetKeywordMode(ed, "Изменить размер атрибута для блоков: ", New String() {"Все", "Полигон", "Диапазон"}, "Полигон")
                If mode.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Прервано пользователем.")
                    Exit Sub
                End If

                '5. новый размер атрибута
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите новый размер атрибута: ")
                pStrOpts.AllowNegative = False
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                pStrOpts.DefaultValue = 1
                Dim pStrRes As PromptDoubleResult = ed.GetDouble(pStrOpts)
                If pStrRes.Status <> PromptStatus.OK Then
                    tr.Commit()
                    Exit Sub
                End If
                If pStrRes.Value <= 0 Then
                    tr.Commit()
                    Exit Sub
                End If
                Dim heightBlk As Double = pStrRes.Value

                '6. меняем размер выбранного атрибута у отобранных точек
                Dim pointIDColl As ObjectIdCollection = GetTargetGeoPointIds(tr, ed, mode, listBlockReference)
                For Each idBlk As ObjectId In pointIDColl
                    If SetAttributeHeight(tr, idBlk, nameAttribute, heightBlk) = True Then CountPoint += 1
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить размер атрибута! Макро прервано")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        MsgBox("Всего преобразовано " & CountPoint & " точек.")
    End Sub
    '===================================================================================================================
    'изменить цвет видимых атрибутов кадастровых точек
    '===================================================================================================================
    <CommandMethod("colorAttributeCadastralPoint")>
    Public Sub ColorAttributeCadastralPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim CountPoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If
                '2. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If
                'режим обработки
                Dim mode As String = GetKeywordMode(ed, "Изменить цвет атрибутов для блоков: ", New String() {"Все", "Полигон", "Диапазон"}, "Полигон")
                If mode.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Прервано пользователем.")
                    Exit Sub
                End If
                'новый цвет: 1-255 — индекс цвета, 0 — ПоБлоку, 256 — ПоСлою,
                'ключевое слово "Диалог" — графический выбор цвета
                Dim pIntOptsColor As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите индекс цвета атрибутов (1-255, 0-ПоБлоку, 256-ПоСлою) или ""Диалог"": ")
                pIntOptsColor.AllowNegative = False
                pIntOptsColor.AllowZero = True
                pIntOptsColor.AllowNone = False
                pIntOptsColor.DefaultValue = 256
                pIntOptsColor.Keywords.Add("Диалог")
                Dim pIntResColor As PromptIntegerResult = ed.GetInteger(pIntOptsColor)
                Dim newColorIndex As Integer = 256
                Dim boolUseColorDialog As Boolean = False
                Dim colorDialog As Autodesk.AutoCAD.Windows.ColorDialog = Nothing
                If pIntResColor.Status = PromptStatus.Keyword Then
                    'графический выбор цвета
                    colorDialog = New Autodesk.AutoCAD.Windows.ColorDialog()
                    If colorDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                        tr.Commit()
                        ed.WriteMessage(vbLf & "Цвет не выбран.")
                        Exit Sub
                    End If
                    boolUseColorDialog = True
                ElseIf pIntResColor.Status = PromptStatus.OK Then
                    newColorIndex = pIntResColor.Value
                Else
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Цвет не выбран.")
                    Exit Sub
                End If
                Dim pointIDColl As ObjectIdCollection = GetTargetGeoPointIds(tr, ed, mode, listBlockReference)
                For Each idBlk As ObjectId In pointIDColl
                    Dim blk As BlockReference = TryCast(tr.GetObject(idBlk, OpenMode.ForWrite), BlockReference)
                    If IsNothing(blk) = True Then Continue For
                    'меняем цвет только у видимых атрибутов
                    Dim boolIzm As Boolean = False
                    Dim attCol As AttributeCollection = blk.AttributeCollection
                    For Each attId As ObjectId In attCol
                        Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                        If IsNothing(attRef) = True Then Continue For
                        If attRef.Visible = True Then
                            'цвет из диалога ставим как есть (в том числе TrueColor), иначе — по индексу АЦИ
                            If boolUseColorDialog = True Then
                                attRef.Color = colorDialog.Color
                            Else
                                attRef.ColorIndex = newColorIndex
                            End If
                            boolIzm = True
                        End If
                    Next
                    If boolIzm = True Then CountPoint += 1
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить цвет атрибутов! Макро прервано")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        MsgBox("Всего преобразовано " & CountPoint & " точек.")
    End Sub
    '===================================================================================================================
    'сдвинуть выбранный видимый атрибут блока на заданное смещение от точки вставки блока
    'Смещение передаётся числами: раньше оно шло через строковый массив и Val(), из-за чего
    'дробная часть терялась (Val("1,5") = 1)
    '===================================================================================================================
    Private Shared Function SetAttributeShift(ByVal tr As Transaction,
                                              ByVal idBlk As ObjectId,
                                              ByVal nameAttribute As String,
                                              ByVal shiftX As Double,
                                              ByVal shiftY As Double) As Boolean
        SetAttributeShift = False
        Dim blk As BlockReference = TryCast(tr.GetObject(idBlk, OpenMode.ForWrite), BlockReference)
        If IsNothing(blk) = True Then Return False
        Dim attCol As AttributeCollection = blk.AttributeCollection
        For Each attId As ObjectId In attCol
            Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
            If IsNothing(attRef) = True Then Continue For
            If attRef.Visible = False Then Continue For
            If attRef.Tag Like nameAttribute = False Then Continue For
            Dim tempPos As Point3d = attRef.Position
            attRef.Position = New Point3d(blk.Position.X + shiftX, blk.Position.Y + shiftY, tempPos.Z)
            SetAttributeShift = True
        Next
        Return SetAttributeShift
    End Function
    '===================================================================================================================
    'изменить местоположение атрибутов кадастровых точек
    'Образец задаёт имя блока и направление смещения атрибута, расстояние до атрибута задаёт пользователь
    '===================================================================================================================
    <CommandMethod("PositionAttributePoint")>
    Public Sub PositionAttributePoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        SaveActiveDocument(acDoc)
        Dim CountPoint As Integer = 0
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                '1. блок-образец
                Dim blkSample As BlockReference = PickSampleBlock(tr, ed)
                If IsNothing(blkSample) = True Then
                    tr.Commit()
                    Exit Sub
                End If

                '2. атрибут, положение которого переносим
                Dim nameAttribute As String = GetAttributeName(ed, GetVisibleAttributeTags(tr, blkSample), "Выберите атрибут, положение которого изменить:")
                If nameAttribute.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Атрибут не выбран.")
                    Exit Sub
                End If

                '3. отбираем блоки этого имени
                Dim listBlockReference As Dictionary(Of String, NetTopologySuite.Geometries.Point) = GetSamplePointDictionary(db, blkSample)
                If listBlockReference.Count = 0 Then
                    tr.Commit()
                    MsgBox("Блоки с именем """ & blkSample.Name & """ не найдены.")
                    Exit Sub
                End If

                '4. смещение выбранного атрибута у образца — из него берём направление
                'и расстояние по умолчанию
                Dim shiftX As Double = 0
                Dim shiftY As Double = 0
                Dim boolFindAttribute As Boolean = False
                Dim attCol As AttributeCollection = blkSample.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                    If IsNothing(attRef) = True Then Continue For
                    If attRef.Visible = False Then Continue For
                    If attRef.Tag Like nameAttribute = False Then Continue For
                    shiftX = attRef.Position.X - blkSample.Position.X
                    shiftY = attRef.Position.Y - blkSample.Position.Y
                    boolFindAttribute = True
                    Exit For
                Next
                If boolFindAttribute = False Then
                    tr.Commit()
                    MsgBox("У выбранного блока нет видимого атрибута " & nameAttribute & ".")
                    Exit Sub
                End If
                Dim distSample As Double = Math.Sqrt(shiftX * shiftX + shiftY * shiftY)

                '5. расстояние от точки до атрибута (по умолчанию — как у образца)
                Dim pDblOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите расстояние от точки до атрибута: ")
                pDblOpts.AllowNegative = False
                pDblOpts.AllowZero = True
                pDblOpts.AllowNone = False
                pDblOpts.DefaultValue = distSample
                Dim pDblRes As PromptDoubleResult = ed.GetDouble(pDblOpts)
                If pDblRes.Status <> PromptStatus.OK Then
                    tr.Commit()
                    Exit Sub
                End If
                Dim distNew As Double = pDblRes.Value

                'новое смещение: направление — от образца, длина — введённая
                Dim shiftNewX As Double = 0
                Dim shiftNewY As Double = 0
                If distNew > 0 Then
                    If distSample <= 0 Then
                        tr.Commit()
                        MsgBox("У образца атрибут стоит в самой точке блока, направление смещения определить нельзя.")
                        Exit Sub
                    End If
                    shiftNewX = shiftX / distSample * distNew
                    shiftNewY = shiftY / distSample * distNew
                End If

                '6. режим обработки
                Dim mode As String = GetKeywordMode(ed, "Изменить положение атрибутов для блоков: ", New String() {"Все", "Диапазон"}, "Диапазон")
                If mode.Trim.Length = 0 Then
                    tr.Commit()
                    ed.WriteMessage(vbLf & "Прервано пользователем.")
                    Exit Sub
                End If

                '7. отбор блоков: Все — по образцу, Диапазон — выбор блоков на чертеже
                Dim pointIDColl As ObjectIdCollection = New ObjectIdCollection()
                If mode Like "Все" Then
                    pointIDColl = GetTargetGeoPointIds(tr, ed, mode, listBlockReference)
                Else
                    Dim acSSPrompt As PromptSelectionResult = ed.GetSelection()
                    If acSSPrompt.Status = PromptStatus.OK Then
                        For Each IDObj As ObjectId In acSSPrompt.Value.GetObjectIds
                            Dim blk As BlockReference = TryCast(tr.GetObject(IDObj, OpenMode.ForRead, False), BlockReference)
                            If IsNothing(blk) = False Then
                                If pointIDColl.Contains(blk.ObjectId) = False Then pointIDColl.Add(blk.ObjectId)
                            End If
                        Next
                    End If
                End If

                '8. отодвигаем атрибут от точки вставки блока
                For Each idBlk As ObjectId In pointIDColl
                    If SetAttributeShift(tr, idBlk, nameAttribute, shiftNewX, shiftNewY) = True Then CountPoint += 1
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Не удалось изменить положение атрибутов! Макро прервано")
                tr.Abort()
                Exit Sub
            End Try
            tr.Commit()
        End Using
        MsgBox("Всего преобразовано " & CountPoint & " точек.")
    End Sub
End Class
