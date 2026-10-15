Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports System.IO
Imports NetTopologySuite
Imports CadastralTools.StructureElement

Public Class MapPlanDataImport
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '=============================================================================================
    'импорт карты (плана) объекта землеустройства (xml-файл Росреестра interact-map-plan 2.0.1)
    'парсер — CadastralTools.ImportMapPlan.RosreestrImportMapPlan (общая библиотека для AutoCAD, Nanocad, Robur)
    'таблица данных объекта: LAND_CADASTRAL_BOUNDARY
    <CommandMethod("ImportRosreestrMapPlan")>
    Public Sub RosreestrImportMapPlan()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim categoryZn As String = "Земельный кадастр"
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
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Try
                Dim generalDirectory As String = FuncFiles.getSupportFolderPath("Civil3DToolsUtility")
                'передаем масштаб чертежа в форму
                Dim nameKvartal As String = ""
                Dim FormKPTImport As CadastralTools.FormImportKPT = New CadastralTools.FormImportKPT(generalDirectory, False, False, False, True, False, False, False)
                Dim tempPath As String = Path.Combine(Path.GetTempPath(), "rosreestr", "mapplan", Guid.NewGuid().ToString())
                FormKPTImport.initDirectory = tempPath
                'запускаем форму
                FormKPTImport.ShowDialog()
                If FormKPTImport.boolShow = False Then
                    Exit Sub
                End If
                Dim templateXML As String = FormKPTImport.PutchTemlate
                If IsNothing(templateXML) = True Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If
                If File.Exists(templateXML) = False Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If
                '================================================================================
                'выбираем файл XML карты (плана)
                Dim fileDialogXML As OpenFileDialog = New OpenFileDialog("Открыть файл XML карты (плана)", "", "xml", "xml", OpenFileDialogFlags.AllowAnyExtension)
                If fileDialogXML.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Exit Sub
                Dim putchXML As String = fileDialogXML.Filename
                If File.Exists(putchXML) = False Then Exit Sub

                '================================================================================
                'папка для промежуточных json-данных (как в ImportRosreestrKPT)
                Directory.CreateDirectory(tempPath)
                'парсер в общей библиотеке CadastralTools
                CadastralTools.ImportMapPlan.RosreestrImportMapPlan(putchXML, tempPath)
                Dim files As String() = Directory.GetFiles(tempPath)
                If files.Length = 0 Then
                    MsgBox("Не удалось корректно прочитать XML. Макро прервано.")
                    Exit Sub
                End If
                '================================================================================
                'рисуем границу в чертеже
                Dim PutchTemplateDwg As String = templateXML.Replace(".xml", ".dwg")
                Dim dictonaryStyle As Dictionary(Of String, CadastralPolylineStyle) = New Dictionary(Of String, CadastralPolylineStyle)
                Using tr As Transaction = Db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    '========================================================================================================================================
                    'копируем стили с файла-шаблона
                    If File.Exists(PutchTemplateDwg) = True Then
                        Dim ArrayLineProperties As String(,) = Nothing
                        Dim countArrayLineProperties As Integer = RosreestrDataImport.FuncListLineObjectDataTableDestDWG(PutchTemplateDwg, categoryZn, ArrayLineProperties)
                        If IsArray(ArrayLineProperties) = True Then
                            If ArrayLineProperties.GetUpperBound(1) > 0 Then
                                Dim boolCopyStyle As Boolean = AcadFunction.FuncCopyLineStylesDestDraw(PutchTemplateDwg, ArrayLineProperties, False)
                            End If
                        End If
                    End If
                    '========================================================================================================================================
                    'читаем объекты из json и выводим контуры границы
                    For Each file As String In files
                        Try
                            Dim listDictionaryData As List(Of Dictionary(Of String, Object)) = CadastralTools.JsonValidator.CreateObjectsFromJsonFile(file)
                            If listDictionaryData.Count = 0 Then Continue For
                            Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                            Dim dictLandProperties As Dictionary(Of String, String) = New Dictionary(Of String, String)
                            Dim PSTable As String = ""
                            Dim Label As String = ""
                            Dim idName As Integer = 0
                            Dim listPolygonCoordinates As List(Of String(,)) = New List(Of String(,))
                            For Each cadastralObject As Dictionary(Of String, Object) In listDictionaryData
                                If IsNothing(cadastralObject) = True Then Continue For
                                If cadastralObject.Count = 0 Then Continue For
                                If cadastralObject.ContainsKey("PS") = True Then
                                    PSTable = cadastralObject.Item("PS")
                                End If
                                If cadastralObject.ContainsKey("LABEL") = True Then
                                    Label = cadastralObject.Item("LABEL")
                                End If
                                If cadastralObject.ContainsKey("NameTypeObject") = True Then
                                    idName = Val(cadastralObject.Item("NameTypeObject"))
                                End If
                                'читаем свойства объекта
                                Dim boolReadDict As Boolean = CadastralTools.DataFormatting.ConvertToDictionaryOfString(cadastralObject, dictLandProperties)
                                '==========================================================================================================
                                'читаем координаты контуров
                                If cadastralObject.ContainsKey("ListCoord") Then
                                    Dim dataCoordParcel As Dictionary(Of String, Object) = cadastralObject("ListCoord")
                                    For Each kvp As KeyValuePair(Of String, Object) In dataCoordParcel
                                        Dim dataCoordCounter As List(Of Object) = TryCast(kvp.Value, List(Of Object))
                                        If dataCoordCounter Is Nothing OrElse dataCoordCounter.Count = 0 Then Continue For
                                        For Each coordItem As Object In dataCoordCounter
                                            Dim listCoordinates As List(Of Object) = TryCast(coordItem, List(Of Object))
                                            If listCoordinates Is Nothing OrElse listCoordinates.Count < 3 Then Continue For
                                            ' Извлекаем списки координат (индексы 1 и 2 — как в ImportRosreestrKPT)
                                            Dim elementCoordX As List(Of Object) = TryCast(listCoordinates(1), List(Of Object))
                                            Dim elementCoordY As List(Of Object) = TryCast(listCoordinates(2), List(Of Object))
                                            If elementCoordX Is Nothing OrElse elementCoordY Is Nothing Then Continue For
                                            If elementCoordX.Count = 0 OrElse elementCoordX.Count <> elementCoordY.Count Then Continue For
                                            ' Создаём массив координат [Точки x Y]
                                            Dim pointCount As Integer = elementCoordX.Count
                                            Dim arrayCoord(pointCount - 1, 1) As String  ' [индекс точки, 0=X, 1=Y]
                                            For k As Integer = 0 To pointCount - 1
                                                arrayCoord(k, 0) = elementCoordX(k).ToString()
                                                arrayCoord(k, 1) = elementCoordY(k).ToString()
                                            Next k
                                            listPolygonCoordinates.Add(arrayCoord)
                                        Next coordItem
                                    Next kvp
                                End If
                            Next
                            'преобразуем в единую строку параметров
                            Dim keyParam As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictLandProperties)
                            '==========================================================================================================
                            'ФОРМИРУЕМ СЕМАНТИКУ
                            dataCadastralObject.Label = Label
                            dataCadastralObject.NameTypeObject = CadastralTools.StructureElement.GetEnumByValue(GetType(nameObject), idName)
                            dataCadastralObject.Description = CadastralTools.StructureElement.GetEnumDescriptionByValue(GetType(nameObject), idName)
                            dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                            dataCadastralObject.KeyParameter = keyParam
                            dataCadastralObject.IdElement = Guid.NewGuid.ToString
                            dataCadastralObject.TablePS = PSTable
                            '======================================================================================================================
                            'таблица ВНЕШНИЙ ОСТРОВОК не требуется — контуры границы не бывают вложенными
                            '======================================================================================================================
                            'набор характеристик LAND_CADASTRAL_BOUNDARY
                            Dim IDPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(dataCadastralObject.TablePS)
                            Dim createPs As Boolean = False
                            If IsNothing(IDPSTable) = True Then
                                createPs = True
                            ElseIf IDPSTable.IsNull = True Then
                                createPs = True
                            End If
                            If createPs = True And templateXML.Trim.Length > 0 Then
                                Dim arrayPropertiesPS As String(,) = {}
                                Dim boolFindPropertiesPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, dataCadastralObject.TablePS, arrayPropertiesPS)
                                If arrayPropertiesPS.Length = 0 Then
                                    If System.IO.File.Exists(PutchTemplateDwg) = True Then
                                        Dim boolFindPS As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, dataCadastralObject.TablePS)
                                        IDPSTable = FuncPropertySet.FuncReturnIDDictSetProperties(dataCadastralObject.TablePS)
                                    End If
                                Else
                                    IDPSTable = FuncPropertySet.FuncCreatedSetProperties(dataCadastralObject.TablePS, arrayPropertiesPS)
                                End If
                            End If
                            '=========================================================================================================================
                            'назначаем стиль
                            Dim newStyle As CadastralPolylineStyle = New CadastralPolylineStyle
                            newStyle.TemplateXML = templateXML
                            newStyle.DBDrawing = Db
                            If dictonaryStyle.ContainsKey(dataCadastralObject.Description) = False Then
                                Dim arrayProperties As String(,) = {}
                                Dim boolFindPropertyObject As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZn, "ЕГРН-Границы территорий", dataCadastralObject.TypeGeometry, dataCadastralObject.Description, arrayProperties)
                                If IsArray(arrayProperties) = True Then
                                    newStyle.createLineStyle(arrayProperties)
                                End If
                                dictonaryStyle.Add(dataCadastralObject.Description, newStyle)
                            Else
                                newStyle = dictonaryStyle.Item(dataCadastralObject.Description)
                            End If
                            '=============================================================================================================================
                            'рисуем контуры границы
                            If listPolygonCoordinates.Count > 0 Then
                                For i As Integer = 0 To listPolygonCoordinates.Count - 1
                                    Dim arrayCoord As String(,) = listPolygonCoordinates.ElementAt(i)
                                    If arrayCoord.Length > 2 Then
                                        Dim acPolyline As Polyline = AcadFunction.createPolylineFromArrayCoordinates2(arrayCoord)
                                        If IsNothing(acPolyline) = False Then
                                            btr.AppendEntity(acPolyline)
                                            tr.AddNewlyCreatedDBObject(acPolyline, True)
                                            newStyle.setStyle(acPolyline)
                                            'граница выводится открытой полилинией (LineString)
                                            acPolyline.Closed = False
                                            'записываем xdata
                                            Dim boolWriteXRecords As Boolean = FuncXRecords.setXRecords(acPolyline, dataCadastralObject)
                                            'записываем PS
                                            If IsNothing(IDPSTable) = False Then
                                                If IDPSTable.IsNull = False Then
                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acPolyline, IDPSTable)
                                                    If boolPSObject = True Then
                                                        Dim boolRec As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(acPolyline, dataCadastralObject)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next i
                            End If
                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        End Try
                    Next
                    tr.Commit()
                End Using
            Catch ex As System.Exception
                MsgBox("Ошибка импорта карты (плана): " & ex.Message)
            End Try
        End Using
    End Sub
End Class
