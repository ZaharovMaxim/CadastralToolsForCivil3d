Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports CadastralTools
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.LinearReferencing

Public Class CreateCadastralObject
    'создать проектируемый участок (новая версия)
    <CommandMethod("ManualCreateCadastralObject")>
    Public Sub ManualCreateCadastralObject()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создаем класс с новой формой
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
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
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вызываем диалоговое окно с первоначальными настройками
        Dim generalDidectory As String = "k:\Civil3DToolsUtility"
        Dim formCreateParcels As CadastralTools.FormCreateCadastralObject = New CadastralTools.FormCreateCadastralObject(generalDidectory)
        formCreateParcels.ShowDialog()
        If formCreateParcels.boolShow = False Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем префикс участка
        Dim numberParcel As Integer = formCreateParcels.numberParcel
        Dim numberCounter As Integer = formCreateParcels.numberCounter
        Dim fileTemlateXML1 As String = formCreateParcels.fileTemlateXML
        Dim fileTemplateDWG As String = fileTemlateXML1.Replace(".xml", ".dwg")
        Dim dictionaryPropertiesObject As Dictionary(Of String, String) = formCreateParcels.resultDictionary
        Dim arrayPropertyObject As String(,) = formCreateParcels.arrayPropertiesZn
        Dim arrayPropertiesLeader As String(,) = formCreateParcels.arrayPropertiesLeader
        Dim nameObject As CadastralTools.StructureElement.nameObject = formCreateParcels.nameTypeObject
        Dim typeGeometryObject As StructureElement.typeGeometryObject = formCreateParcels.typeGeometry
        Dim deskObject As String = formCreateParcels.nameObject
        Dim deskLeader As String = formCreateParcels.deskLeader
        Dim nameTableObject As String = formCreateParcels.nameTableZn
        Dim PS As String = formCreateParcels.PS
        Dim globalID As String = ""
        Dim boolCreateMleader As Boolean = formCreateParcels.boolCreateMleader
        If dictionaryPropertiesObject.ContainsKey("GLOBALID") = True Then
            globalID = dictionaryPropertiesObject.Item("GLOBALID")
        End If
        If dictionaryPropertiesObject.Count > 0 Then
            If arrayPropertyObject.Length > 0 Then
                '=========================================================================================================================
                'назначаем стиль объекту
                Dim styleObject As CadastralPolylineStyle = New CadastralPolylineStyle()
                styleObject.DBDrawing = db
                styleObject.TemplateXML = fileTemlateXML1
                styleObject.TemplateDWG = fileTemplateDWG
                Dim arrayProperties As String(,) = {}
                styleObject.createLineStyle(arrayPropertyObject)
                '=========================================================================================================================
                'PS
                Dim arrayPropertiesPs As String(,) = {}
                Dim IDPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(PS)
                If IsNothing(IDPSTable) = True Then
                    Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, PS, arrayPropertiesPs)
                    If arrayPropertiesPs.Length > 0 Then
                        IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
                    End If
                ElseIf IDPSTable.IsNull = True Then
                    Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, PS, arrayPropertiesPs)
                    If arrayPropertiesPs.Length > 0 Then
                        IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
                    End If
                End If
                '===========================================================================================================================
                Dim formatLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "NOTE")
                Dim cadastralNetSuit As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                Dim localClassTopology As FuncNetTopologySuite = New FuncNetTopologySuite()
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем объекты
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    Try
                        If typeGeometryObject = StructureElement.typeGeometryObject.polygon Or typeGeometryObject = StructureElement.typeGeometryObject.polyline Then
                            Dim acSSPrompt As PromptSelectionResult
                            Dim acSSet1 As SelectionSet = Nothing
                            acSSPrompt = ed.GetSelection() 'выбираем контура
                            If acSSPrompt.Status = PromptStatus.OK Then
                                acSSet1 = acSSPrompt.Value
                            Else
                                Exit Sub
                            End If
                            For Each objId1 As ObjectId In acSSet1.GetObjectIds()
                                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(objId1, OpenMode.ForWrite)
                                If TypeOf ent Is Polyline Then
                                    Dim plineParcels As Polyline = ent
                                    'точка вставки мультивыноски
                                    Dim innerPoint As NetTopologySuite.Geometries.Point = Nothing
                                    'обрабатываем полигон
                                    If typeGeometryObject = StructureElement.typeGeometryObject.polygon Then
                                        If plineParcels.NumberOfVertices > 2 Then
                                            Dim polygon As NetTopologySuite.Geometries.Polygon = localClassTopology.CreatePolygon(plineParcels)
                                            If polygon IsNot Nothing AndAlso Not polygon.IsEmpty AndAlso polygon.Area > 0 Then
                                                'ищем внутреннюю точку
                                                innerPoint = cadastralNetSuit.GetInteriorPoint(polygon)
                                                If dictionaryPropertiesObject.ContainsKey("Area") = True Then
                                                    dictionaryPropertiesObject.Item("Area") = Math.Round(plineParcels.Area, 0)
                                                ElseIf dictionaryPropertiesObject.ContainsKey("KeyParameter") = True Then
                                                    dictionaryPropertiesObject.Item("KeyParameter") = Math.Round(plineParcels.Area, 0)
                                                End If
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            Continue For
                                        End If
                                        'обрабатываем не замкнутую полилинию
                                    ElseIf typeGeometryObject = StructureElement.typeGeometryObject.polyline Then
                                        If plineParcels IsNot Nothing AndAlso plineParcels.NumberOfVertices > 1 Then
                                            Dim cadastralLine As NetTopologySuite.Geometries.LineString = localClassTopology.CreateLineString(plineParcels)
                                            If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                                ' Создаем полилинию (если нужно)
                                                cadastralLine = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(cadastralLine, 0.01)
                                                If cadastralLine.Count > 1 Then
                                                    If boolCreateMleader = True Then
                                                        ' Ищем внутреннюю точку на ГРАНИЦЕ (используем cadastralLine, а не plineParcels)
                                                        Dim locationIndexedLine As New LocationIndexedLine(cadastralLine)
                                                        ' Вычисляем половину длины от cadastralLine
                                                        Dim totalLength As Double = cadastralLine.Length / 2.0
                                                        ' Создаем "карту" для преобразования расстояния в LinearLocation
                                                        Dim lengthLocationMap As New LengthLocationMap(cadastralLine)
                                                        ' Преобразуем длину в LinearLocation
                                                        Dim targetLocation As LinearLocation = lengthLocationMap.GetLocation(totalLength)
                                                        ' Извлекаем координату
                                                        Dim midCoordinate As Coordinate = locationIndexedLine.ExtractPoint(targetLocation)
                                                        innerPoint = New NetTopologySuite.Geometries.Point(midCoordinate)
                                                    End If
                                                    If dictionaryPropertiesObject.ContainsKey("Lenght") = True Then
                                                        dictionaryPropertiesObject.Item("Lenght") = Math.Round(plineParcels.Length, 2)
                                                    ElseIf dictionaryPropertiesObject.ContainsKey("KeyParameter") = True Then
                                                        dictionaryPropertiesObject.Item("KeyParameter") = Math.Round(plineParcels.Length, 2)
                                                    End If
                                                Else
                                                    MsgBox("В указанной полилинии менее 2 вершин. Преобразование не завершено.")
                                                    Continue For
                                                End If
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            Continue For
                                        End If
                                    Else
                                        Continue For
                                    End If
                                    'обновляем свойства
                                    Dim conditionalParcel As String = numberParcel
                                    If numberCounter > 0 Then
                                        conditionalParcel = numberParcel & "(" & numberCounter & ")"
                                        numberCounter += 1
                                    Else
                                        numberParcel += 1
                                    End If
                                    'обновляем условный номер и локальный ID
                                    If dictionaryPropertiesObject.ContainsKey("ConditionalNumber") = True Then
                                        dictionaryPropertiesObject.Item("ConditionalNumber") = conditionalParcel
                                    End If
                                    If dictionaryPropertiesObject.ContainsKey("LOCALID") = True Then
                                        dictionaryPropertiesObject.Item("LOCALID") = Guid.NewGuid.ToString
                                    End If
                                    'формируем свойства объекта
                                    Dim keyParamObject As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictionaryPropertiesObject)
                                    Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                    dataCadastralObject.Label = formCreateParcels.nameCategory
                                    dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.TERRITORIAL_PLANNING
                                    dataCadastralObject.TypeGeometry = formCreateParcels.typeGeometry
                                    dataCadastralObject.NameTypeObject = nameObject
                                    dataCadastralObject.TableObject = nameTableObject
                                    dataCadastralObject.Description = deskObject
                                    dataCadastralObject.KeyParameter = keyParamObject
                                    dataCadastralObject.TablePS = formCreateParcels.PS
                                    dataCadastralObject.IdStructure = globalID
                                    dataCadastralObject.IdElement = Guid.NewGuid.ToString
                                    dataCadastralObject.IdObject = plineParcels.Handle.ToString
                                    Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(plineParcels, dataCadastralObject)
                                    If IsNothing(IDPSTable) = False Then
                                        If IDPSTable.IsNull = False Then
                                            Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(plineParcels, IDPSTable)
                                            If boolPSObject = True Then
                                                Dim boolRec As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(plineParcels, dataCadastralObject)
                                            End If
                                        End If
                                    End If
                                    styleObject.setStyle(plineParcels)
                                    '=========================================================================================================================
                                    'рисуем выноску
                                    '=========================================================================================================================
                                    If arrayPropertiesLeader.Length > 0 Then
                                        'назначаем стиль выноске
                                        Dim styleLeaderObject As CadastralLeaderStyle = New CadastralLeaderStyle()
                                        styleLeaderObject.TemplateXML = fileTemlateXML1
                                        styleLeaderObject.TemplateDWG = fileTemplateDWG
                                        styleLeaderObject.DBDrawing = db
                                        'PS
                                        Dim arrayPropertiesLabelPS As String(,) = {}
                                        Dim boolFindLabelPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, "PROJECT_LABEL", arrayPropertiesLabelPS)
                                        Dim IDLabelPSTable As ObjectId = Nothing
                                        If arrayPropertiesLabelPS.Length > 0 Then
                                            IDLabelPSTable = FuncPropertySet.FuncCreatedSetProperties("PROJECT_LABEL", arrayPropertiesLabelPS)
                                        End If
                                        '=======================================================================================================================
                                        Dim tempLayerLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Layer")
                                        If Mid(tempLayerLeader.Trim, 1, 1) Like "_" Then
                                            Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader, plineParcels.Layer)
                                        Else
                                            Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader)
                                        End If
                                        '===============================================================
                                        'вторая точка
                                        Dim pointBound As Extents3d = plineParcels.Bounds
                                        Dim maskLabel As String = ""
                                        Dim strNoteLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Note")
                                        If strNoteLeader.Trim.Length > 0 Then
                                            conditionalParcel = CadastralTools.DataFormatting.getTextLeader(strNoteLeader, dictionaryPropertiesObject, maskLabel)
                                            If conditionalParcel.Trim.Length > 0 Then
                                                Dim leaderObject As MLeader = classDrawObject.createLeader(New Point3d(innerPoint.X, innerPoint.Y, 0), conditionalParcel, pointBound.MaxPoint)
                                                If IsNothing(leaderObject) = False Then
                                                    styleLeaderObject.setLeaderStyle(leaderObject)
                                                    Dim createXdata As StructureElement = CadastralTools.CadastralLeader.getLeaderXData(dataCadastralObject, deskLeader, maskLabel, fileTemlateXML1)
                                                    Dim boolWriteDataLeader As Boolean = FuncXRecords.setXRecords(leaderObject, createXdata)
                                                    If IsNothing(IDLabelPSTable) = False Then
                                                        'формируем новый массив для Xrecords
                                                        Dim ArrayDataObj As String(,) = Nothing
                                                        ReDim ArrayDataObj(2, 4)
                                                        'условный номер
                                                        ArrayDataObj(0, 0) = "PROJECT_LABEL"
                                                        ArrayDataObj(1, 0) = "LABEL"
                                                        ArrayDataObj(2, 0) = deskLeader
                                                        'условный номер
                                                        ArrayDataObj(0, 1) = "PROJECT_LABEL"
                                                        ArrayDataObj(1, 1) = "PARENTID"
                                                        ArrayDataObj(2, 1) = dataCadastralObject.IdElement
                                                        'id контура
                                                        ArrayDataObj(0, 2) = "PROJECT_LABEL"
                                                        ArrayDataObj(1, 2) = "MASK"
                                                        ArrayDataObj(2, 2) = maskLabel
                                                        'локальный id
                                                        ArrayDataObj(0, 3) = "PROJECT_LABEL"
                                                        ArrayDataObj(1, 3) = "LOCALID"
                                                        ArrayDataObj(2, 3) = Guid.NewGuid.ToString
                                                        'примечание
                                                        ArrayDataObj(0, 4) = "PROJECT_LABEL"
                                                        ArrayDataObj(1, 4) = "NOTE"
                                                        ArrayDataObj(2, 4) = ""
                                                        If IsNothing(IDLabelPSTable) = False Then
                                                            If IDLabelPSTable.IsNull = False Then
                                                                Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(leaderObject, IDLabelPSTable)
                                                                If boolADDPSTBL = True Then
                                                                    Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(leaderObject, ArrayDataObj)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        ElseIf typeGeometryObject = StructureElement.typeGeometryObject.block Then
                            'загружаем в чертеж блок
                            'CadCursors.GetPoint(userCadView,, "Укажите точку вставки блока: ")
                            'userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgInsert, "Выборите блоки для преобразования: ")
                        End If
                    Catch ex As System.Exception
                        MsgBox(ex.Message.ToString)
                        tr.Abort()
                    End Try
                    tr.Commit()
                End Using
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'добавить островок
    <CommandMethod("AddCadastralObjectRing")>
    Public Sub CreateUserOstrov()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim generalDir As String = ""
        Dim templateXML As String = ""
        Dim templateDWG As String = ""
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
        If dictionaryTemplate.Count > 0 Then
            templateXML = dictionaryTemplate.ElementAt(0).Value
        End If
        If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
        '=========================================================================================================================
        'PS
        Dim arrayPropertiesPs As String(,) = {}
        Dim PS As String = "CADASTRAL_OSTROV"
        Dim IDPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(PS)
        If IsNothing(IDPSTable) = True Then
            Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, PS, arrayPropertiesPs)
            If arrayPropertiesPs.Length > 0 Then
                IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
            End If
        ElseIf IDPSTable.IsNull = True Then
            Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, PS, arrayPropertiesPs)
            If arrayPropertiesPs.Length > 0 Then
                IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
            End If
        End If
        If IsNothing(IDPSTable) = False Then
            If IDPSTable.IsNull = False Then
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim per As PromptEntityResult = ed.GetEntity("Выберите контур к которому будет добавлен островок: ")
                    'выбор полигона которому принадлежит этот островок
                    Dim selectParentObject As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект
                    If TypeOf selectParentObject Is Polyline Then
                        Dim parentPolyline As Polyline = selectParentObject
                        Dim dataParent As StructureElement = Nothing
                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentPolyline, dataParent)
                        If IsNothing(dataParent) = True Then
                            MsgBox("Выберите объект с присвоенными кадастрвыми данными")
                            Exit Sub
                        End If
                        'выбираем объекты внутри выбранного контура
                        Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = SelectCadastralObject.selectAllObject(db)
                        Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                        Dim acadNetTopoSuite As FuncNetTopologySuite = New FuncNetTopologySuite()
                        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                        Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentPolyline)
                        lineStyle.DBDrawing = db
                        If IsNothing(dataParent) = False And IsNothing(parentPolyline) = False Then
                            Dim parentID As String = dataParent.IdElement
                            Dim keyParam As String = dataParent.KeyParameter
                            Dim parentPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(parentPolyline)
                            Dim arrayPropertiesCadastralOstrov As String(,) = {}
                            Dim boolFindPropertiesOstrov As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, "CADASTRAL_OSTROV", arrayPropertiesCadastralOstrov)
                            Dim numberOstrov As Integer = CadastralTools.CadastralOstrov.getLastNumber(parentPolygon, dictionaryCadastralObject)
                            numberOstrov += 1
                            Dim listPolygon As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                            Dim listOstrov As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                            If dictionaryCadastralObject.ContainsKey(dataParent.NameTypeObject) = True Then
                                listPolygon = dictionaryCadastralObject.Item(dataParent.NameTypeObject)
                            End If
                            If dictionaryCadastralObject.ContainsKey(StructureElement.nameObject.ProjectOstrov) = True Then
                                listOstrov = dictionaryCadastralObject.Item(StructureElement.nameObject.ProjectOstrov)
                            End If
                            parentPolygon = CadastralTools.CadastralPolygon.getCounter(parentPolygon, listPolygon, listOstrov)
                            Dim areaParcel As Double = parentPolygon.Area
                            If arrayPropertiesCadastralOstrov.Length > 0 Then
                                'список полигонов островков для их упорядочивания
                                Dim listPolygonRing As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'выбираем островки
                                Dim acSSPrompt As PromptSelectionResult
                                Dim acSSet1 As SelectionSet = Nothing
                                acSSPrompt = ed.GetSelection() 'выбираем контура
                                If acSSPrompt.Status = PromptStatus.OK Then
                                    acSSet1 = acSSPrompt.Value
                                Else
                                    MsgBox("Прервано пользователем.")
                                    Exit Sub
                                End If
                                For Each acSSObj As ObjectId In acSSet1.GetObjectIds()
                                    If Not IsDBNull(acSSObj) Then
                                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj, OpenMode.ForWrite)
                                        If TypeOf (acEnt) Is Polyline Then
                                            Dim plineOstrov As Polyline = acEnt
                                            If plineOstrov.NumberOfVertices > 2 Then
                                                If plineOstrov.Area > 0 Then
                                                    Dim dataObject As StructureElement = Nothing
                                                    Dim boolReadDataObject As Boolean = FuncXRecords.getXRecords(plineOstrov, dataObject)
                                                    If boolReadDataObject = True Then
                                                        plineOstrov = plineOstrov.Clone
                                                        acBlkTblRec.AppendEntity(plineOstrov)
                                                        tr.AddNewlyCreatedDBObject(plineOstrov, True)
                                                    End If
                                                End If
                                                Dim ostrovPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(plineOstrov)
                                                ostrovPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(ostrovPolygon, 0.01)
                                                ostrovPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(ostrovPolygon)
                                                If ostrovPolygon.Coordinates.Count > 2 Then
                                                    If ostrovPolygon.Area > 0 Then
                                                        ostrovPolygon.UserData = plineOstrov.ObjectId
                                                        listPolygonRing.Add(ostrovPolygon)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                                If listPolygonRing.Count > 2 Then
                                    'упорядочиваем полигоны
                                    listPolygonRing = cadastralNetTopoSuite.SortByNearestNeighbor(listPolygonRing)
                                End If
                                For Each userPolygon As NetTopologySuite.Geometries.Polygon In listPolygonRing
                                    Dim idPolyline As ObjectId = userPolygon.UserData
                                    Dim polylineRing As Polyline = tr.GetObject(idPolyline, OpenMode.ForWrite)
                                    If IsNothing(polylineRing) = False Then
                                        polylineRing = classDrawObject.createPolylineByPolygon(userPolygon, polylineRing)
                                        'формируем свойства объекта
                                        Dim dataCadastralObject As CadastralTools.StructureElement = CadastralTools.CadastralOstrov.getProjectOstrovXData(dataParent, numberOstrov, arrayPropertiesCadastralOstrov)
                                        dataCadastralObject.IdObject = polylineRing.Handle.ToString
                                        Dim boolWriteDataOstriv As Boolean = FuncXRecords.writeXRecords(polylineRing, dataCadastralObject, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                        lineStyle.setStyle(polylineRing)
                                        parentPolygon = cadastralNetTopoSuite.AddPolygonAsHole(parentPolygon, userPolygon)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'формируем новый массив для Xrecords
                                        Dim ArrayDataObj As String(,) = Nothing
                                        ReDim ArrayDataObj(2, 6)
                                        'условный номер
                                        ArrayDataObj(0, 0) = PS
                                        ArrayDataObj(1, 0) = "LABEL"
                                        ArrayDataObj(2, 0) = dataCadastralObject.Description
                                        'условный номер
                                        ArrayDataObj(0, 1) = PS
                                        ArrayDataObj(1, 1) = "ConditionalNumber"
                                        ArrayDataObj(2, 1) = numberOstrov
                                        'условный номер
                                        ArrayDataObj(0, 2) = PS
                                        ArrayDataObj(1, 2) = "ParentNameType"
                                        ArrayDataObj(2, 2) = CadastralTools.StructureElement.GetEnumNameByValue(GetType(CadastralTools.StructureElement.nameObject), dataParent.NameTypeObject)
                                        'условный номер родительского контура
                                        ArrayDataObj(0, 3) = PS
                                        ArrayDataObj(1, 3) = "ParentNumber"
                                        ArrayDataObj(2, 3) = CadastralTools.JsonValidator.GetValueFromJson(dataParent.KeyParameter, "ConditionalNumber")
                                        'id контура
                                        ArrayDataObj(0, 4) = PS
                                        ArrayDataObj(1, 4) = "IDKOUNTER"
                                        ArrayDataObj(2, 4) = dataCadastralObject.IdStructure
                                        'локальный id
                                        ArrayDataObj(0, 5) = PS
                                        ArrayDataObj(1, 5) = "LOCALID"
                                        ArrayDataObj(2, 5) = dataCadastralObject.IdElement
                                        'примечание
                                        ArrayDataObj(0, 6) = PS
                                        ArrayDataObj(1, 6) = "NOTE"
                                        ArrayDataObj(2, 6) = dataCadastralObject.Note
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'находим таблицу
                                        Dim boolPSTables As Boolean = FuncPropertySet.FuncFindPSetTable(polylineRing, PS)
                                        If boolPSTables = True Then
                                            Dim boolADDDATA As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(polylineRing, ArrayDataObj)
                                        Else
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(polylineRing, IDPSTable)
                                            Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(polylineRing, ArrayDataObj)
                                        End If
                                        numberOstrov += 1
                                    End If
                                Next
                                Dim rezArea As Double = Math.Round(parentPolygon.Area, 0)
                                keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "Area", rezArea)
                                dataParent.KeyParameter = keyParam
                                Dim boolNewData As Boolean = FuncXRecords.writeXRecords(parentPolyline, dataParent, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES)
                            End If
                        End If
                    End If
                    tr.Commit()
                End Using
            End If
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'обновляем метку площадь
        'Dim acMleader As MLeader = Nothing
        '                Dim ContextParcel As String = ""
        '                acMleader = FuncCadastr.FuncFindMLeaderByLocalId(ParentLocalID)
        '                If IsNothing(acMleader) = False Then
        '                    If acMleader.IsWriteEnabled = False Then
        '                        acMleader.UpgradeOpen()
        '                    End If
        '                    Dim acMtext As String = acMleader.MText.Contents
        '                    '============================================================================================
        '                    'разбираем строку
        '                    If acMleader.HasFields = True Then
        '                        Dim id As ObjectId = acMleader.GetField() 'смотрим старый атрибут
        '                        Dim fd As Field = tr.GetObject(id, OpenMode.ForWrite) 'читаем его
        '                        Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
        '                        Dim pos2 As Integer = CodStrFiled.IndexOf("Area")
        '                        If pos2 > -1 Then
        '                            Dim fstr As String = FuncCadastr.FuncCreateFieldSQRTPolyline(Entobj.Id, idOstrov)
        '                            Dim pos3 As Integer = CodStrFiled.IndexOf("Площадь")
        '                            Dim pos4 As Integer = CodStrFiled.IndexOf("кв")
        '                            If pos3 > -1 And pos4 > 1 Then
        '                                Dim str1 As String = Mid(CodStrFiled, 1, pos3 + 7)
        '                                Dim str2 As String = Mid(CodStrFiled, pos4 + 1)
        '                                fd.SetFieldCode(str1 & fstr & str2)
        '                                fd.Evaluate()
        '                            End If
        '                        End If
        '                    End If
        '                End If
        '            End If
        '        End If
        '    Catch ex As System.Exception
        '        MsgBox(ex.ToString & vbLf & "Ошибка ввода данных. Макро прервано.")
        '        tr.Commit()
        '        locker.Dispose()
        '        Exit Sub
        '    End Try
        '    tr.Commit()
        '    locker.Dispose()
        'End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'добавить контур к участку
    <CommandMethod("AddCadastralObjectCounter")>
    Public Sub AddUserKounterObject()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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

        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim generalDir As String = ""
        Dim templateXML As String = ""
        Dim templateDWG As String = ""
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        'Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
        'If dictionaryTemplate.Count > 0 Then
        'templateXML = dictionaryTemplate.ElementAt(0).Value
        'End If
        'If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
        'MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
        'Exit Sub
        'End If
        'сброс выделения
        Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = SelectCadastralObject.selectAllObject(db)
        Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
        Dim acadNetTopoSuite As FuncNetTopologySuite = New FuncNetTopologySuite()
        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim per As PromptEntityResult = ed.GetEntity("Выберите объект к которому будеь добавлен контур: ")
                If per.Status = PromptStatus.OK Then
                    'выбор полигона которому принадлежит этот островок
                    Dim selectParentObject As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект
                    If TypeOf selectParentObject Is Polyline Then
                        Dim parentPolyline As Polyline = selectParentObject
                        Dim dataParentParcel As StructureElement = Nothing
                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentPolyline, dataParentParcel)
                        If IsNothing(dataParentParcel) = True Then
                            tr.Abort()
                            locker.Dispose()
                            MsgBox("Выберите объект с присвоенными кадастрвыми данными")
                            Exit Sub
                        End If
                        'выбираем подобные объекты в чертеже
                        If IsNothing(dataParentParcel) = False And IsNothing(parentPolyline) = False Then
                            Dim deskSelectObject As String = dataParentParcel.Description
                            Dim idStructure As String = dataParentParcel.IdStructure
                            Dim keyParamSelectObject As String = dataParentParcel.KeyParameter
                            Dim NamePS As String = dataParentParcel.TablePS
                            Dim IDLabelPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NamePS)
                            Dim conditionalParentParcel As String = ""
                            Dim numberParentParcel As Integer = 0
                            Dim numberParentCounterParcel As Integer = 0
                            Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentPolyline)
                            lineStyle.DBDrawing = db
                            Dim dictionaryDataParent As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(keyParamSelectObject)
                            If IsNothing(dictionaryDataParent) = False Then
                                If dictionaryDataParent.Count > 0 Then
                                    If dictionaryDataParent.ContainsKey("ConditionalNumber") = True Then
                                        conditionalParentParcel = dictionaryDataParent.Item("ConditionalNumber")
                                        numberParentParcel = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalParentParcel)
                                        numberParentCounterParcel = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalParentParcel)
                                    End If
                                End If
                            End If
                            'читаем PS это приоритет
                            Dim tempConditionalParentParcel As String = FuncPropertySet.FuncReadValueSetProperties(parentPolyline, "ConditionalNumber")
                            Dim tempNumberParentParcel As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(tempConditionalParentParcel)
                            Dim tempNumberParentCounterParcel As Integer = CadastralTools.DataFormatting.GetNumberInParenthesis(tempConditionalParentParcel)
                            If tempConditionalParentParcel.Trim.Length > 0 Then
                                conditionalParentParcel = tempConditionalParentParcel
                            End If
                            If tempConditionalParentParcel > 0 Then
                                numberParentParcel = tempNumberParentParcel
                            End If
                            If tempNumberParentCounterParcel > 0 Then
                                numberParentCounterParcel = tempNumberParentCounterParcel
                            End If
                            'перезаписываем исходный ЗУ
                            If numberParentParcel = 0 Then numberParentParcel = 1
                            If numberParentCounterParcel = 0 Then numberParentCounterParcel = 1
                            Dim newConditionatParentParcel As String = numberParentParcel & "(" & numberParentCounterParcel & ")"
                            If Not (newConditionatParentParcel Like conditionalParentParcel) Then
                                'перезаписываем свойства основного участка
                                keyParamSelectObject = CadastralTools.JsonValidator.SetJsonFieldValue(keyParamSelectObject, "ConditionalNumber", newConditionatParentParcel)
                                dataParentParcel.KeyParameter = keyParamSelectObject
                                If dataParentParcel.IdStructure.Length < 2 Then
                                    dataParentParcel.IdStructure = Guid.NewGuid.ToString
                                End If
                                Dim boolRecData As Boolean = FuncXRecords.writeXRecords(parentPolyline, dataParentParcel, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                boolReadData = FuncPropertySet.FuncWriteDataSetByStructureElement(parentPolyline, dataParentParcel)
                            End If
                            'увеличиваем порядковый номер для контура

                            numberParentCounterParcel += 1
                            'список полигонов островков для их упорядочивания
                            Dim listCounter As List(Of NetTopologySuite.Geometries.Geometry) = New List(Of NetTopologySuite.Geometries.Geometry)
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'выбираем контура
                            Dim acSSPrompt As PromptSelectionResult
                            Dim acSSet1 As SelectionSet = Nothing
                            acSSPrompt = ed.GetSelection() 'выбираем контура
                            If acSSPrompt.Status = PromptStatus.OK Then
                                acSSet1 = acSSPrompt.Value
                            Else
                                MsgBox("Прервано пользователем.")
                                Exit Sub
                            End If
                            For Each acSSObj As ObjectId In acSSet1.GetObjectIds()
                                If Not IsDBNull(acSSObj) Then
                                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj, OpenMode.ForWrite)
                                    If TypeOf (acEnt) Is Polyline Then
                                        Dim plineCounter As Polyline = acEnt
                                        If dataParentParcel.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                            If plineCounter.NumberOfVertices > 2 Then
                                                If plineCounter.Area > 0 Then
                                                    Dim counterPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(plineCounter)
                                                    counterPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(counterPolygon, 0.01)
                                                    counterPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(counterPolygon)
                                                    If counterPolygon.Coordinates.Count > 2 Then
                                                        If counterPolygon.Area > 0 Then
                                                            counterPolygon.UserData = plineCounter.ObjectId
                                                            listCounter.Add(counterPolygon)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        ElseIf dataParentParcel.TypeGeometry = StructureElement.typeGeometryObject.line Then
                                            Dim cadastralLine As NetTopologySuite.Geometries.LineString = acadNetTopoSuite.CreateLineString(plineCounter)
                                            If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                                cadastralLine = cadastralNetTopoSuite.RemoveDuplicateVertices(cadastralLine, 0.01)
                                                If cadastralLine.Coordinates.Count > 2 Then
                                                    If cadastralLine.Area > 0 Then
                                                        cadastralLine.UserData = plineCounter.ObjectId
                                                        listCounter.Add(cadastralLine)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            If listCounter.Count > 2 Then
                                'упорядочиваем полигоны
                                listCounter = cadastralNetTopoSuite.SortByNearestNeighbor(listCounter)
                            End If
                            '======================================================================================================================
                            'создаем контура
                            For Each topoObject As NetTopologySuite.Geometries.Geometry In listCounter
                                Dim idPolyline As ObjectId = topoObject.UserData
                                Dim polylineCounter As Polyline = tr.GetObject(idPolyline, OpenMode.ForWrite)
                                If polylineCounter.ObjectId.IsNull = False Then
                                    If TypeOf topoObject Is NetTopologySuite.Geometries.Polygon Then
                                        polylineCounter = classDrawObject.createPolylineByPolygon(topoObject, polylineCounter)
                                    ElseIf TypeOf topoObject Is LineString Then
                                        polylineCounter = classDrawObject.createPolylineByLineString(topoObject, polylineCounter)
                                    End If
                                    Dim newDataCounter As StructureElement = New StructureElement
                                    newDataCounter = dataParentParcel
                                    Dim keyParamCounter As String = newDataCounter.KeyParameter
                                    If dictionaryDataParent.ContainsKey("ConditionalNumber") = True Then
                                        keyParamCounter = CadastralTools.JsonValidator.SetJsonFieldValue(keyParamCounter, "ConditionalNumber", numberParentParcel & "(" & numberParentCounterParcel & ")")
                                    End If
                                    If dictionaryDataParent.ContainsKey("Area") = True Then
                                        keyParamCounter = CadastralTools.JsonValidator.SetJsonFieldValue(keyParamCounter, "Area", Math.Round(polylineCounter.Area, 0))
                                    End If
                                    If dictionaryDataParent.ContainsKey("Lenght") = True Then
                                        keyParamCounter = CadastralTools.JsonValidator.SetJsonFieldValue(keyParamCounter, "Lenght", Math.Round(polylineCounter.Area, 0))
                                    End If
                                    newDataCounter.KeyParameter = keyParamCounter
                                    newDataCounter.IdElement = Guid.NewGuid.ToString
                                    newDataCounter.IdStructure = idStructure
                                    Dim boolWriteDataOstriv As Boolean = FuncXRecords.writeXRecords(polylineCounter, newDataCounter, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                    lineStyle.setStyle(polylineCounter)
                                    If IsNothing(IDLabelPSTable) = False Then
                                        If IDLabelPSTable.IsNull = False Then
                                            Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(polylineCounter, IDLabelPSTable)
                                            If boolADDPSTBL = True Then
                                                Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(polylineCounter, newDataCounter)
                                            End If
                                        End If
                                    End If
                                    numberParentCounterParcel += 1
                                End If
                            Next
                        End If
                    End If
                End If
                tr.Commit()
            End Using
        End Using

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'меняем метку
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'формируем строку подписи
        '===============================================================================================================
        'If boolDrawMLeader = True Then
        '                                    '===============================================================================================================
        '                                    'формируем строку подписы
        '                                    Dim ParcelNumerator As String = ""
        '                                    Dim ParcelDenumerator As String = ""
        '                                    If numerator.IndexOf("ConditionalNumber") > -1 Then
        '                                        ParcelNumerator = numerator.Replace("<ConditionalNumber>", oboznZU)
        '                                    End If
        '                                    If denumerator.IndexOf("ConditionalNumber") > -1 Then
        '                                        ParcelDenumerator = denumerator.Replace("<ConditionalNumber>", oboznZU)
        '                                    End If

        '                                    If numerator.IndexOf("Name") > -1 Then
        '                                        ParcelNumerator = numerator.Replace("<Description>", DeskDercel)
        '                                    End If
        '                                    If denumerator.IndexOf("Name") > -1 Then
        '                                        ParcelDenumerator = denumerator.Replace("<Description>", DeskDercel)
        '                                    End If

        '                                    If numerator.IndexOf("KeyParameter") > -1 Then
        '                                        ParcelNumerator = numerator.Replace("<KeyParameter>", keyParam)
        '                                    End If
        '                                    If denumerator.IndexOf("KeyParameter") > -1 Then
        '                                        ParcelDenumerator = denumerator.Replace("<KeyParameter>", keyParam)
        '                                    End If

        '                                    If numerator.IndexOf("TypeWork") > -1 Then
        '                                        ParcelNumerator = numerator.Replace("<TypeWork>", workType)
        '                                    End If
        '                                    If denumerator.IndexOf("TypeWork") > -1 Then
        '                                        ParcelDenumerator = denumerator.Replace("<TypeWork>", workType)
        '                                    End If

        '                                    If numerator.IndexOf("Area") > -1 Then
        '                                        Dim StrId As Long = PlineCounter.Id.OldIdPtr
        '                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%"
        '                                        ParcelNumerator = numerator.Replace("<Area>", AreaStr)
        '                                    End If
        '                                    If denumerator.IndexOf("Area") > -1 Then
        '                                        Dim StrId As Long = PlineCounter.Id.OldIdPtr
        '                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%"
        '                                        ParcelDenumerator = denumerator.Replace("<Area>", AreaStr)
        '                                    End If

        '                                    If numerator.IndexOf("Lenght") > -1 Then
        '                                        Dim StrId As Long = PlineCounter.Id.OldIdPtr
        '                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Length \f " & """" & "%lu2%pr0" & """" & ">%"
        '                                        ParcelNumerator = numerator.Replace("<Lenght>", AreaStr)
        '                                    End If
        '                                    If denumerator.IndexOf("Lenght") > -1 Then
        '                                        Dim StrId As Long = PlineCounter.Id.OldIdPtr
        '                                        Dim AreaStr As String = "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Length \f " & """" & "%lu2%pr0" & """" & ">%"
        '                                        ParcelDenumerator = denumerator.Replace("<Lenght>", AreaStr)
        '                                    End If

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование объектов по аналогу
    <CommandMethod("ReplaceCadastralObject")>
    Public Sub ReplaceCadastralObject()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim strDWGName As String = acDoc.Name
        Try
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
        End Try
        Dim stylePolyline As CadastralPolylineStyle = Nothing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'получаем доступ к объекту
            Dim parentObjectSelectResult As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект в качестве образца (Полилинию): ")
            If parentObjectSelectResult.Status = PromptStatus.Cancel Then
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            'ссылка на объект-аналог
            Dim parentObject As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(parentObjectSelectResult.ObjectId, OpenMode.ForRead)
            Dim dataObject As StructureElement = Nothing
            Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentObject, dataObject)
            Dim tablePS As String = ""
            If IsNothing(dataObject) = False Then
                tablePS = dataObject.TablePS
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем расширенные данные
            Dim arrayDataPs As String(,) = {}
            Dim boolPSReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(parentObject, arrayDataPs)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'считываем стиль
            stylePolyline = New CadastralPolylineStyle(parentObject)
            stylePolyline.DBDrawing = acCurDb
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем новый объект
            Dim SelectNewObject As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Полилинию: ")
            If SelectNewObject.Status = PromptStatus.OK Then
                Dim entityObject As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectNewObject.ObjectId, OpenMode.ForWrite)
                If TypeOf entityObject Is Polyline Then
                    stylePolyline.setStyle(entityObject)
                    'удаляем все PS
                    Dim ArrayPSTableObj As String() = Nothing
                    Dim boolReadTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(entityObject, ArrayPSTableObj)
                    'удаляем лишние таблицы
                    If IsArray(ArrayPSTableObj) = True Then
                        For i As Integer = 0 To ArrayPSTableObj.Length - 1
                            Dim NameTlb As String = ArrayPSTableObj(i)
                            If IsNothing(NameTlb) Then
                                Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables(entityObject, NameTlb)
                            End If
                        Next
                    End If
                    Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(tablePS)
                    If CreateTBLPSID.IsNull = False Then
                        Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(entityObject, CreateTBLPSID)
                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(entityObject, arrayDataPs)
                    End If
                    Dim boolXRec As Boolean = FuncXRecords.writeXRecords(entityObject, dataObject)
                    If boolXRec = True Then
                        If parentObject.IsWriteEnabled = True Then
                            parentObject.Erase()
                        End If
                    End If
                Else
                    MsgBox("Необходимо выдрать примитив - Полилинию")
                End If
            End If
            acTrans.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'автоматическое создание земельных участков
    <CommandMethod("AutomaticCreateCadastralObject")>
    Public Sub AutomaticCreateCadastralObject()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim generalDidectory As String = "k:\Civil3DToolsUtility"
        Dim formCreateParcels As CadastralTools.FormCreateCadastralObject = New CadastralTools.FormCreateCadastralObject(generalDidectory)
        formCreateParcels.ShowDialog()
        If formCreateParcels.boolShow = False Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем префикс участка
        Dim numberParcel As Integer = formCreateParcels.numberParcel
        Dim numberCounter As Integer = formCreateParcels.numberCounter
        Dim fileTemlateXML1 As String = formCreateParcels.fileTemlateXML
        Dim fileTemplateDWG As String = fileTemlateXML1.Replace(".xml", ".dwg")
        Dim dictionaryPropertiesObject As Dictionary(Of String, String) = formCreateParcels.resultDictionary
        Dim arrayPropertyObject As String(,) = formCreateParcels.arrayPropertiesZn
        Dim arrayPropertiesLeader As String(,) = formCreateParcels.arrayPropertiesLeader
        Dim nameObject As CadastralTools.StructureElement.nameObject = formCreateParcels.nameTypeObject
        Dim typeGeometryObject As StructureElement.typeGeometryObject = formCreateParcels.typeGeometry
        Dim deskObject As String = formCreateParcels.nameObject
        Dim deskLeader As String = formCreateParcels.deskLeader
        Dim nameTableObject As String = formCreateParcels.nameTableZn
        Dim PS As String = formCreateParcels.PS
        Dim globalID As String = ""
        Dim boolCreateMleader As Boolean = formCreateParcels.boolCreateMleader
        If dictionaryPropertiesObject.ContainsKey("GLOBALID") = True Then
            globalID = dictionaryPropertiesObject.Item("GLOBALID")
        End If
        If dictionaryPropertiesObject.Count > 0 Then
            If arrayPropertyObject.Length > 0 Then
                Dim classTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite()
                Dim acadTopologySuite As FuncNetTopologySuite = New FuncNetTopologySuite
                Dim processor As New CadastralTools.PolygonIntersectionInnerProcessor
                Dim classDrawingObject As CreateDWGObject = New CreateDWGObject(db)
                'свойства линии
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'блокируем чертеж
                Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using locker
                    Dim tr As Transaction = db.TransactionManager.StartTransaction()
                    Using tr
                        '=========================================================================================================================
                        'назначаем стиль объекту
                        Dim styleObject As CadastralPolylineStyle = New CadastralPolylineStyle()
                        styleObject.DBDrawing = db
                        styleObject.TemplateXML = fileTemlateXML1
                        styleObject.TemplateDWG = fileTemplateDWG
                        Dim arrayProperties As String(,) = {}
                        styleObject.createLineStyle(arrayPropertyObject)
                        '=========================================================================================================================
                        'PS
                        Dim arrayPropertiesPs As String(,) = {}
                        Dim IDPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(PS)
                        If IsNothing(IDPSTable) = True Then
                            Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, PS, arrayPropertiesPs)
                            If arrayPropertiesPs.Length > 0 Then
                                IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
                            End If
                        ElseIf IDPSTable.IsNull = True Then
                            Dim boolFindPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, PS, arrayPropertiesPs)
                            If arrayPropertiesPs.Length > 0 Then
                                IDPSTable = FuncPropertySet.FuncCreatedSetProperties(PS, arrayPropertiesPs)
                            End If
                        End If
                        '===========================================================================================================================
                        Dim formatLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "NOTE")
                        Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                        Dim outerPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                        'выбираем замкнутую полилинию из чертежа или внешней ссылки по которой будет производится обрезка
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim selectOuterBoundary As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию по которой будет происходить обрезка: ")
                        If selectOuterBoundary.Status = PromptStatus.OK Then
                            Dim ids As List(Of ObjectId) = New List(Of ObjectId)(selectOuterBoundary.GetContainers())
                            ids.Reverse()
                            ids.Add(selectOuterBoundary.ObjectId)
                            For Each id As ObjectId In ids
                                Dim idStr As String = id.ObjectClass.Name
                                If idStr Like "AcDbPolyline" Then
                                    Dim putchSelectBoundares = id.Database.Filename
                                    Dim polyOuterBoundary As Polyline = New Polyline
                                    Using startPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Polyline)
                                        Dim dataOuterBoundary As StructureElement = Nothing
                                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(startPoly, dataOuterBoundary, StructureElement.tableXRecords.PROJECT_STRUCTURES, startPoly.Database)
                                        If IsNothing(dataOuterBoundary) = False Then
                                            Dim deskObjectOuterObject As String = dataOuterBoundary.Description
                                            Dim dictSelectObject As Dictionary(Of String, List(Of Geometry)) = SelectCadastralObject.getCadastralPolygon(startPoly.Database, deskObjectOuterObject)
                                            If dictSelectObject.Count > 0 Then
                                                Dim conditionatOuterBound As String = CadastralTools.JsonValidator.GetValueFromJson(dataOuterBoundary.KeyParameter, "ConditionalNumber")
                                                If IsNothing(conditionatOuterBound) = False Then
                                                    If conditionatOuterBound.Trim.Length > 0 Then
                                                        Dim listParcels As List(Of NetTopologySuite.Geometries.Geometry) = dictSelectObject.Item(deskObjectOuterObject)
                                                        Dim a = 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                        outerPolygon = acadTopologySuite.CreatePolygon(startPoly)
                                    End Using
                                End If
                            Next
                        Else
                            MsgBox("Прервано пользователем.")
                            tr.Abort()
                            locker.Dispose()
                            Exit Sub
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'выбираем полигоны которые необходимо загрузить
                        'выбираем один объект, остальные будут аналогами
                        Dim listPolygon As List(Of Polygon) = New List(Of Polygon)
                        Dim putchNameKPT As String = ""
                        Dim nameLayer As String = "0"
                        Dim selectInnerBoundary As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите кадастровый план: ")
                        If selectInnerBoundary.Status = PromptStatus.OK Then
                            Dim ids As List(Of ObjectId) = New List(Of ObjectId)(selectInnerBoundary.GetContainers())
                            ids.Reverse()
                            ids.Add(selectInnerBoundary.ObjectId)
                            For Each id As ObjectId In ids
                                Dim idStr As String = id.ObjectClass.Name
                                If idStr Like "AcDbPolyline" Then
                                    Dim nameFileKPT As String = id.Database.Filename
                                    Using selectObject As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                                        nameLayer = selectObject.Layer
                                        Dim strLayerArr As String() = nameLayer.Split("|")
                                        If strLayerArr.Length > 1 Then
                                            nameLayer = strLayerArr(1)
                                        End If
                                        If IO.File.Exists(nameFileKPT) = True Then
                                            putchNameKPT = nameFileKPT
                                            Exit For
                                        End If
                                    End Using
                                End If
                            Next
                        Else
                            MsgBox("Прервано пользователем.")
                            tr.Abort()
                            locker.Dispose()
                            Exit Sub
                        End If
                        '===================================================================================================================
                        'читаем все объекты
                        Dim sourceDb As Database = New Database(False, True)
                        sourceDb.ReadDwgFile(putchNameKPT, FileShare.Read, True, String.Empty)
                        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                        Using sourceTr
                            Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                            For Each idObject As ObjectId In sourceBtr
                                If idObject.IsNull = False Then
                                    Dim idStr As String = idObject.ObjectClass.Name
                                    If idStr Like "AcDbPolyline" Then
                                        Dim poly As Polyline = sourceTr.GetObject(idObject, OpenMode.ForRead)
                                        If poly.Closed = True Then
                                            If poly.Layer Like nameLayer Then
                                                Dim innerPolygon As NetTopologySuite.Geometries.Polygon = acadTopologySuite.CreatePolygon(poly)
                                                If IsNothing(innerPolygon) = False Then
                                                    Dim allPolygons As List(Of Polygon) = processor.GetIntersectionPolygonsWithCheck(outerPolygon, innerPolygon)
                                                    If allPolygons.Count > 0 Then
                                                        Dim dataCadastralObject As StructureElement = Nothing
                                                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(poly, dataCadastralObject, StructureElement.tableXRecords.PROJECT_STRUCTURES, sourceDb)
                                                        For i As Integer = 0 To allPolygons.Count - 1
                                                            Dim workPolygon As Polygon = allPolygons.Item(i)
                                                            If workPolygon.Area >= 1 Then
                                                                Dim newParcels As Polyline = classDrawObject.createPolylineByPolygon(workPolygon)
                                                                If IsNothing(dataCadastralObject) = False Then
                                                                    Dim dictDataObject As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(dataCadastralObject.KeyParameter)
                                                                    If IsNothing(dictDataObject) = False Then
                                                                        For k As Integer = 0 To dictionaryPropertiesObject.Count - 1
                                                                            Dim nameField As String = dictionaryPropertiesObject.ElementAt(k).Key
                                                                            If nameField Like "Cad_number" Then
                                                                                If dictDataObject.ContainsKey(nameField) = True Then
                                                                                    Dim value As String = dictDataObject.ContainsKey(nameField)
                                                                                    dictionaryPropertiesObject.Item("Cad_number") = value
                                                                                End If
                                                                            ElseIf nameField Like "Address" Then
                                                                                If dictDataObject.ContainsKey(nameField) = True Then
                                                                                    Dim value As String = dictDataObject.ContainsKey(nameField)
                                                                                    dictionaryPropertiesObject.Item("Address") = value
                                                                                End If
                                                                            ElseIf nameField Like "Category" Then
                                                                                If dictDataObject.ContainsKey(nameField) = True Then
                                                                                    Dim value As String = dictDataObject.ContainsKey(nameField)
                                                                                    If value Like "Земли населенных пунктов" Then
                                                                                        dictionaryPropertiesObject.Item("Category") = value
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Next
                                                                    End If
                                                                End If
                                                                'обновляем свойства
                                                                Dim conditionalParcel As String = numberParcel
                                                                If numberCounter > 0 Then
                                                                    conditionalParcel = numberParcel & "(" & numberCounter & ")"
                                                                    numberCounter += 1
                                                                Else
                                                                    numberParcel += 1
                                                                End If
                                                                'обновляем условный номер и локальный ID
                                                                If dictionaryPropertiesObject.ContainsKey("ConditionalNumber") = True Then
                                                                    dictionaryPropertiesObject.Item("ConditionalNumber") = conditionalParcel
                                                                End If
                                                                If dictionaryPropertiesObject.ContainsKey("LOCALID") = True Then
                                                                    dictionaryPropertiesObject.Item("LOCALID") = Guid.NewGuid.ToString
                                                                End If
                                                                Dim keyParam As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictionaryPropertiesObject)

                                                                Dim newDataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                                                newDataCadastralObject.Label = formCreateParcels.nameCategory
                                                                newDataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.TERRITORIAL_PLANNING
                                                                newDataCadastralObject.TypeGeometry = formCreateParcels.typeGeometry
                                                                newDataCadastralObject.NameTypeObject = nameObject
                                                                newDataCadastralObject.TableObject = nameTableObject
                                                                newDataCadastralObject.Description = deskObject
                                                                newDataCadastralObject.KeyParameter = keyParam
                                                                newDataCadastralObject.TablePS = formCreateParcels.PS
                                                                newDataCadastralObject.IdStructure = globalID
                                                                newDataCadastralObject.IdElement = Guid.NewGuid.ToString
                                                                newDataCadastralObject.IdObject = newParcels.Handle.ToString
                                                                Dim boolWriteDataOstriv As Boolean = FuncXRecords.setXRecords(newParcels, dataCadastralObject)
                                                                If IsNothing(IDPSTable) = False Then
                                                                    If IDPSTable.IsNull = False Then
                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(newParcels, IDPSTable)
                                                                        If boolPSObject = True Then
                                                                            Dim boolRec As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(newParcels, dataCadastralObject)
                                                                        End If
                                                                    End If
                                                                End If
                                                                styleObject.setStyle(newParcels)
                                                                '=========================================================================================================================
                                                                'рисуем выноску
                                                                '=========================================================================================================================
                                                                If arrayPropertiesLeader.Length > 0 Then
                                                                    'назначаем стиль выноске
                                                                    Dim innerPoint = classTopologySuite.GetInteriorPoint(workPolygon)
                                                                    Dim styleLeaderObject As CadastralLeaderStyle = New CadastralLeaderStyle()
                                                                    styleLeaderObject.TemplateXML = fileTemlateXML1
                                                                    styleLeaderObject.TemplateDWG = fileTemplateDWG
                                                                    styleLeaderObject.DBDrawing = db
                                                                    'PS
                                                                    Dim arrayPropertiesLabelPS As String(,) = {}
                                                                    Dim boolFindLabelPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(fileTemlateXML1, "PROJECT_LABEL", arrayPropertiesLabelPS)
                                                                    Dim IDLabelPSTable As ObjectId = Nothing
                                                                    If arrayPropertiesLabelPS.Length > 0 Then
                                                                        IDLabelPSTable = FuncPropertySet.FuncCreatedSetProperties("PROJECT_LABEL", arrayPropertiesLabelPS)
                                                                    End If
                                                                    '=======================================================================================================================
                                                                    Dim tempLayerLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Layer")
                                                                    If Mid(tempLayerLeader.Trim, 1, 1) Like "_" Then
                                                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader, newParcels.Layer)
                                                                    Else
                                                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader)
                                                                    End If
                                                                    '===============================================================
                                                                    'вторая точка
                                                                    Dim pointBound As Extents3d = newParcels.Bounds
                                                                    Dim maskLabel As String = ""
                                                                    Dim strNoteLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Note")
                                                                    If strNoteLeader.Trim.Length > 0 Then
                                                                        conditionalParcel = CadastralTools.DataFormatting.getTextLeader(strNoteLeader, dictionaryPropertiesObject, maskLabel)
                                                                        If conditionalParcel.Trim.Length > 0 Then
                                                                            Dim leaderObject As MLeader = classDrawObject.createLeader(New Point3d(innerPoint.X, innerPoint.Y, 0), conditionalParcel, pointBound.MaxPoint)
                                                                            If IsNothing(leaderObject) = False Then
                                                                                styleLeaderObject.setLeaderStyle(leaderObject)
                                                                                Dim createXdata As StructureElement = CadastralTools.CadastralLeader.getLeaderXData(dataCadastralObject, deskLeader, maskLabel, fileTemlateXML1)
                                                                                If IsNothing(createXdata) = False Then
                                                                                    Dim boolWriteDataLeader As Boolean = FuncXRecords.setXRecords(leaderObject, createXdata)
                                                                                    If IsNothing(IDLabelPSTable) = False Then
                                                                                        'формируем новый массив для Xrecords
                                                                                        Dim ArrayDataObj As String(,) = Nothing
                                                                                        ReDim ArrayDataObj(2, 4)
                                                                                        'условный номер
                                                                                        ArrayDataObj(0, 0) = "PROJECT_LABEL"
                                                                                        ArrayDataObj(1, 0) = "LABEL"
                                                                                        ArrayDataObj(2, 0) = deskLeader
                                                                                        'условный номер
                                                                                        ArrayDataObj(0, 1) = "PROJECT_LABEL"
                                                                                        ArrayDataObj(1, 1) = "PARENTID"
                                                                                        ArrayDataObj(2, 1) = dataCadastralObject.IdElement
                                                                                        'id контура
                                                                                        ArrayDataObj(0, 2) = "PROJECT_LABEL"
                                                                                        ArrayDataObj(1, 2) = "MASK"
                                                                                        ArrayDataObj(2, 2) = maskLabel
                                                                                        'локальный id
                                                                                        ArrayDataObj(0, 3) = "PROJECT_LABEL"
                                                                                        ArrayDataObj(1, 3) = "LOCALID"
                                                                                        ArrayDataObj(2, 3) = Guid.NewGuid.ToString
                                                                                        'примечание
                                                                                        ArrayDataObj(0, 4) = "PROJECT_LABEL"
                                                                                        ArrayDataObj(1, 4) = "NOTE"
                                                                                        ArrayDataObj(2, 4) = ""
                                                                                        If IsNothing(IDLabelPSTable) = False Then
                                                                                            If IDLabelPSTable.IsNull = False Then
                                                                                                Dim boolADDPSTBL As Boolean = FuncPropertySet.FuncAddDataSetProperties(leaderObject, IDLabelPSTable)
                                                                                                If boolADDPSTBL = True Then
                                                                                                    Dim boolADDPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(leaderObject, ArrayDataObj)
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                                numberParcel += 1
                                                            End If
                                                        Next i
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            tr.Commit()
                        End Using
                    End Using
                End Using
            End If
        End If
    End Sub
End Class
