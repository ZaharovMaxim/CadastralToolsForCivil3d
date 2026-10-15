Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports CadastralTools

Public Class UpdateCadastralObject
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование объектов по аналогу
    <CommandMethod("RenumberCadastralObject")>
    Public Sub RenumberCadastralObject()
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
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объект
            Dim SelectNewObject As PromptEntityResult = ed.GetEntity(vbLf & "Выберите один объект для образца: ")
            If SelectNewObject.Status = PromptStatus.OK Then
                Dim entityObject As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectNewObject.ObjectId, OpenMode.ForWrite)
                Dim dataCadastralObject As CadastralTools.StructureElement = Nothing
                Dim boolReadData As Boolean = FuncXRecords.getXRecords(entityObject, dataCadastralObject, CadastralTools.StructureElement.tableXRecords.CADASTRAL_OBJECT)
                If IsNothing(dataCadastralObject) = False Then
                    If IsNothing(dataCadastralObject.Description) = False Then
                        Dim dictSelect As Dictionary(Of String, NetTopologySuite.Geometries.Geometry) = SelectCadastralObject.getCadastralObjectByDesk(acCurDb, dataCadastralObject.Description, dataCadastralObject.TypeGeometry)
                        If dictSelect.Count > 1 Then
                            Dim listCadastralObject As List(Of NetTopologySuite.Geometries.Geometry) = dictSelect.Values.ToList
                            If listCadastralObject.Count > 1 Then
                                listCadastralObject = CadastralTools.FuncNetTopologySuite.SortByNearestNeighbor(listCadastralObject)
                                Dim nextNumber As Integer = 1
                                Dim dictNumbers As New Dictionary(Of Integer, Integer)
                                For i As Integer = 0 To listCadastralObject.Count - 1
                                    Dim polygon As NetTopologySuite.Geometries.Geometry = listCadastralObject.Item(i)
                                    Dim dataObject As CadastralTools.StructureElement = polygon.UserData
                                    Dim conditionalNumber As String = CadastralTools.JsonValidator.GetValueFromJson(dataObject.KeyParameter, "ConditionalNumber")
                                    Dim numberParcel As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalNumber)
                                    Dim numberCounter As Integer = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalNumber)
                                    Dim newConditionalNumber As String = numberParcel
                                    If dictNumbers.ContainsKey(numberParcel) = True Then
                                        numberCounter = dictNumbers.Item(numberParcel) + 1
                                        dictNumbers.Item(numberParcel) = numberCounter
                                        newConditionalNumber = numberParcel & "(" & numberCounter & ")"
                                    Else
                                        dictNumbers.Add(numberParcel, 1)
                                    End If
                                    Dim newKeyParameter As String = CadastralTools.JsonValidator.SetJsonFieldValue(dataObject.KeyParameter, "ConditionalNumber", newConditionalNumber)
                                    dataObject.KeyParameter = newKeyParameter
                                    'получаем доступ к полилинии
                                    Dim handlePolyline As String = dataObject.IdObject
                                    Dim idPolyline As ObjectId = AcadFunction.FuncReturnObjectID(handlePolyline)
                                    Dim polyline As Entity = idPolyline.GetObject(OpenMode.ForWrite)
                                    Dim boolXRecords As Boolean = FuncXRecords.setXRecords(polyline, dataObject)
                                    Dim boolUpdatePS As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(polyline, dataObject)
                                    'обновляем мультивыноску
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Sub CreateLeader()

    End Sub

End Class
