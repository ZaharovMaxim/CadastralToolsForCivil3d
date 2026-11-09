Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports NetTopologySuite
Imports System.Linq
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports CadastralTools
Imports Microsoft.Office.Interop.Word
Imports Document = Autodesk.AutoCAD.ApplicationServices.Document
Imports Editor = Autodesk.AutoCAD.EditorInput.Editor
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class CadastralLeader
    '=============================================================================================
    'восстановить метки для участков
    <CommandMethod("RestoreLabelCadastralObject")>
    Public Sub RestoreLabelCadastralObject()
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
        Dim generalDidectory As String = FuncFiles.getSupportFolderPath("Civil3DToolsUtility")
        Dim listLayer As List(Of String) = StyleCadastralObject.getAllLayersDWG(db)
        Dim listTextStyle As List(Of String) = StyleCadastralObject.getAllTextStyleDwg(db)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'блокируем чертеж
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
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
                    Dim listTypeObject As List(Of String) = dictCadastralAllObject.Keys.ToList
                    'вызываем форму
                    Dim formRectoreLabel As CadastralTools.FormCreateLabelParcels = New CadastralTools.FormCreateLabelParcels(generalDidectory, listLayer, listTextStyle, listTypeObject)
                    formRectoreLabel.ShowDialog()
                    If formRectoreLabel.boolShow = False Then
                        tr.Abort()
                        locker.Dispose()
                        MsgBox("Прервано рользователем.")
                        Exit Sub
                    End If
                    Dim layerMLeader As String = formRectoreLabel.nameLayer
                    Dim nameStyleMLeader As String = formRectoreLabel.nameStyleLeder
                    Dim hText As Double = formRectoreLabel.Height
                    Dim mask As String = formRectoreLabel.maskTextLeader
                    Dim arrayMask As String() = CadastralTools.CadastralLeader.ParseNumeratorDenominator(mask)
                    If IsArray(arrayMask) = True Then
                        If arrayMask.Length > 0 Then
                            mask = arrayMask(0)
                        End If
                        If arrayMask.Length > 1 Then
                            mask = mask & vbLf & arrayMask(1)
                        End If
                    End If
                    Dim templateXML As String = formRectoreLabel.fileTemplateXML
                    Dim templateDWG As String = templateXML.Replace(".xml", ".dwg")
                    Dim nameTextStyle As String = formRectoreLabel.nameTextStyle
                    Dim idTextStyle As ObjectId = StyleCadastralObject.getIDTextStyleByName(db, nameTextStyle)
                    If IsNothing(idTextStyle) = True Then
                        idTextStyle = db.Textstyle
                    ElseIf idTextStyle.IsNull = True Then
                        idTextStyle = db.Textstyle
                    End If
                    Dim styleMleader As CadastralLeaderStyle = New CadastralLeaderStyle()
                    styleMleader.DBDrawing = db
                    styleMleader.TemplateDWG = templateDWG
                    styleMleader.TemplateXML = templateXML
                    Dim idStyleMleader As ObjectId = CadastralLeaderStyle.getStyleId(db, nameStyleMLeader)
                    If IsNothing(idStyleMleader) = True Then
                        Dim boolCopyStyles As Boolean = styleMleader.copyStyle(templateDWG, nameStyleMLeader, False)
                    ElseIf idStyleMleader.IsNull = True Then
                        Dim boolCopyStyles As Boolean = styleMleader.copyStyle(templateDWG, nameStyleMLeader, False)
                    End If
                    idStyleMleader = CadastralLeaderStyle.getStyleId(db, nameStyleMLeader)
                    If IsNothing(idStyleMleader) = True Then
                        MsgBox("Не удалось скопировать стиль из внешнего файла")
                        Exit Sub
                    ElseIf idStyleMleader.IsNull = True Then
                        MsgBox("Не удалось скопировать стиль из внешнего файла")
                        Exit Sub
                    End If
                    'копируем таблицу PS
                    Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_LABEL")
                    If idTablePs.IsNull = True Then
                        idTablePs = FuncPropertySet.FuncCreateTablePSByXML(templateXML, "PROJECT_LABEL")
                    End If
                    Dim listLeader As Dictionary(Of String, CadastralTools.StructureElement) = SelectCadastralObject.GetMLeaderDwg(db)
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    Dim acadNetTopologySuite As FuncNetTopologySuite = New FuncNetTopologySuite()
                    Dim cadastralNetTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                    'смотрим тип объектов
                    Dim deskObject As String = formRectoreLabel.typeObject
                    If dictCadastralAllObject.ContainsKey(deskObject) = True Then
                        'получаем нужный тип объектов (если это полигоны, то еще с учетом дырок, если они там есть
                        Dim dictCadastralObject As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Geometry)) = CadastralTools.SelectCadastralObject.getCadastralObjectByDesk(dictCadastralAllObject, deskObject)
                        If dictCadastralObject.Count > 0 Then
                            Dim cadastralStyle As CadastralLeaderStyle = New CadastralLeaderStyle
                            cadastralStyle.TemplateDWG = templateDWG
                            cadastralStyle.TemplateXML = templateXML
                            cadastralStyle.DBDrawing = db
                            cadastralStyle.NameLayer = layerMLeader
                            cadastralStyle.IDLeaderStyle = idStyleMleader
                            cadastralStyle.Label = formRectoreLabel.labelMLeader
                            cadastralStyle.TextStyleLeader = formRectoreLabel.nameTextStyle
                            CreateCadastralLeader(dictCadastralObject, cadastralStyle, mask, listLeader, idTablePs)
                        End If
                    End If
                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    Public Sub CreateCadastralLeader(ByVal dictCadastralObject As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Geometry)), ByRef leaderCadastralStyle As CadastralLeaderStyle, ByVal mask As String, ByRef listLeader As Dictionary(Of String, CadastralTools.StructureElement), ByVal idPSTable As ObjectId)
        If IsNothing(leaderCadastralStyle) = True Then Exit Sub
        Dim db As Database = leaderCadastralStyle.DBDrawing
        If IsNothing(db) = True Then Exit Sub
        If IsNothing(dictCadastralObject) = False Then
            If dictCadastralObject.Count > 0 Then
                'обозначение участка
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim acadNetTopologySuite As FuncNetTopologySuite = New FuncNetTopologySuite()
                    Dim cadastralNetTopologySuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                    For i As Integer = 0 To dictCadastralObject.Count - 1
                        Dim listCadastralObject As List(Of NetTopologySuite.Geometries.Geometry) = dictCadastralObject.ElementAt(i).Value
                        'последовательно достаем объекты
                        For Each cadastralObject As NetTopologySuite.Geometries.Geometry In listCadastralObject
                            If IsNothing(cadastralObject) = True Then Continue For
                            Dim cadastralData As CadastralTools.StructureElement = cadastralObject.UserData
                            'габариты (прямоугольник)
                            Dim bounds = cadastralNetTopologySuite.GetBounds(cadastralObject)
                            If IsNothing(cadastralData) = False Then
                                Dim keyParam As String = cadastralData.KeyParameter
                                If IsNothing(keyParam) = True Then Continue For
                                If keyParam.Trim.Length = 0 Then Continue For
                                Dim dictCadastralData As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(keyParam) 'FuncPropertySet.FuncReadDataIoDictionary(acEnt, namePs)
                                If IsNothing(dictCadastralData) = True Then Continue For
                                'составляем метку
                                If dictCadastralData.Count > 0 And mask.Trim.Length > 0 Then
                                    Dim textLeader As String = mask
                                    For j As Integer = 0 To dictCadastralData.Count - 1
                                        Dim nameField As String = dictCadastralData.ElementAt(j).Key
                                        If IsNothing(nameField) = True Then Continue For
                                        If nameField.Trim.Length = 0 Then Continue For
                                        Dim value As String = dictCadastralData.ElementAt(j).Value
                                        'поле Area и поле Lenght обраьатывает особым образом
                                        If nameField Like "Area" Or nameField Like "area" Then
                                            If TypeOf cadastralObject Is Geometries.Polygon Then
                                                Dim cadastralPolugon As NetTopologySuite.Geometries.Polygon = cadastralObject
                                                'id основного контура
                                                Dim idPline As ObjectId = AcadFunction.FuncReturnObjectID(cadastralData.IdObject)
                                                'забираем островки
                                                Dim arrayLineRing As Geometries.LinearRing() = cadastralPolugon.Holes
                                                Dim listRingId As New List(Of String)
                                                If IsArray(arrayLineRing) = True Then
                                                    For k As Integer = 0 To arrayLineRing.Length - 1
                                                        Dim lineR As Geometries.LinearRing = arrayLineRing(k)
                                                        Dim userData As StructureElement = lineR.UserData
                                                        If IsNothing(userData) = False Then
                                                            Dim hgObject As String = userData.IdObject
                                                            Dim idPlineRing As ObjectId = AcadFunction.FuncReturnObjectID(hgObject)
                                                            listRingId.Add(idPlineRing.ToString)
                                                        End If
                                                    Next k
                                                End If
                                                value = CadastralTools.CadastralLeader.getAreaCadastralObject(idPline.ToString, listRingId)
                                            End If
                                        End If
                                        Dim strNameField As String = "<" & nameField & ">"
                                        If textLeader.IndexOf(strNameField) > -1 Then
                                            textLeader = textLeader.Replace(strNameField, value)
                                        End If
                                    Next j
                                    'оформление метки
                                    Dim localId As String = cadastralData.IdElement
                                    Dim textStyle As String = leaderCadastralStyle.TextStyleLeader
                                    Dim idTextStyle As ObjectId = CadastralLeaderStyle.getIDTextStyleByName(db, textStyle)
                                    '==============================================================================================================
                                    'ищем существующую мультивыноску
                                    Dim dataMleader As CadastralTools.StructureElement = CadastralTools.CadastralLeader.getMleaderCadastralObject(localId, listLeader)
                                    Dim boolCurrentMLeader As Boolean = False
                                    If IsNothing(dataMleader) = False Then
                                        'мультивыноска существует
                                        Dim hgStr As String = dataMleader.IdObject
                                        If IsNothing(hgStr) = False Then
                                            If hgStr.Trim.Length > 0 Then
                                                Dim idMLeader As ObjectId = AcadFunction.FuncReturnObjectID(hgStr)
                                                If IsNothing(idMLeader) = False Then
                                                    If idMLeader.IsNull = False Then
                                                        Dim MLeader1 As MLeader = tr.GetObject(idMLeader, OpenMode.ForWrite)
                                                        If MLeader1.IsErased = False Then
                                                            If MLeader1.IsWriteEnabled = True Then
                                                                Dim acMText1 As MText = MLeader1.MText
                                                                acMText1.SetDatabaseDefaults()
                                                                acMText1.Contents = textLeader
                                                                MLeader1.MText = acMText1
                                                                If leaderCadastralStyle.HeightTextLeader > 0 Then
                                                                    acMText1.Height = leaderCadastralStyle.HeightTextLeader
                                                                End If
                                                                leaderCadastralStyle.setLeaderStyle(MLeader1)
                                                                'пишем служебную информацию
                                                                Dim dataNewMLeader As CadastralTools.StructureElement = CadastralTools.CadastralLeader.getLeaderXData(cadastralData, leaderCadastralStyle.Label, mask, leaderCadastralStyle.TemplateXML)
                                                                Dim boolWriteData As Boolean = FuncXRecords.setXRecords(MLeader1, dataNewMLeader)
                                                                If IsNothing(idPSTable) = False Then
                                                                    If idPSTable.IsNull = False Then
                                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idPSTable)
                                                                        Dim boolWriteDAta1 As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(MLeader1, dataNewMLeader)
                                                                    End If
                                                                End If
                                                                boolCurrentMLeader = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    If boolCurrentMLeader = False Then
                                        Dim MLeader1 As MLeader = New MLeader
                                        MLeader1.SetDatabaseDefaults()
                                        'точка вставки
                                        Dim CoordText As Point3d = New Point3d
                                        Dim MaxPoint As Point3d = New Point3d(bounds.Value.MaxX, bounds.Value.MaxY, 0)
                                        If cadastralData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon Then
                                            Dim pt As NetTopologySuite.Geometries.Point = cadastralNetTopologySuite.GetInteriorPoint(cadastralObject)
                                            CoordText = New Point3d(pt.X, pt.Y, 0)
                                            Dim leaderIndex As Integer = MLeader1.AddLeader()
                                            Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                                            MLeader1.AddFirstVertex(leaderLineIndex, CoordText)
                                            MLeader1.AddLastVertex(leaderLineIndex, MaxPoint)
                                            MLeader1.ContentType = ContentType.MTextContent
                                            Dim acMText1 As MText = New MText()
                                            acMText1.SetDatabaseDefaults()
                                            acMText1.Contents = textLeader
                                            MLeader1.MText = acMText1
                                            If leaderCadastralStyle.HeightTextLeader > 0 Then
                                                acMText1.Height = leaderCadastralStyle.HeightTextLeader
                                            End If
                                            leaderCadastralStyle.setLeaderStyle(MLeader1)
                                            acBlkTblRec.AppendEntity(MLeader1)
                                            tr.AddNewlyCreatedDBObject(MLeader1, True)
                                            Dim dataNewMLeader As CadastralTools.StructureElement = CadastralTools.CadastralLeader.getLeaderXData(cadastralData, leaderCadastralStyle.Label, mask, leaderCadastralStyle.TemplateXML)
                                            Dim boolWriteData As Boolean = FuncXRecords.setXRecords(MLeader1, dataNewMLeader)
                                            If IsNothing(idPSTable) = False Then
                                                If idPSTable.IsNull = False Then
                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idPSTable)
                                                    Dim boolWriteDAta1 As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(MLeader1, dataNewMLeader)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next i
                    tr.Commit()
                End Using
            End If
        End If
    End Sub
End Class
