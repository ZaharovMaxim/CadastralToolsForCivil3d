Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Xml
Imports Autodesk.AutoCAD.Runtime

Public Class FuncPropertySet
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'SET PROPERTIES
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все данные объекта (выходные данные - массив - 0-имя словаря, 1-имя поля, 2-значение, 3-описание поля)
    Public Shared Function FuncReadAllDataSetProperties(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByRef ArrayData As String(,), Optional ArrayTableXML As String(,) = Nothing) As Boolean
        FuncReadAllDataSetProperties = False
        Dim CountArrayData As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim NameTBL As String = pset.PropertySetDefinitionName
                                If IsArray(ArrayTableXML) = True Then
                                    For i As Integer = 0 To ArrayTableXML.GetUpperBound(1)
                                        If NameTBL Like ArrayTableXML(0, i) Then
                                            If psetDataColl.Count > 0 Then
                                                For j As Integer = 0 To psetDataColl.Count - 1
                                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = psetDataColl.Item(j)
                                                    Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                    Dim data As String = pTempData.GetData() 'значение
                                                    ReDim Preserve ArrayData(3, CountArrayData)
                                                    ArrayData(0, CountArrayData) = NameTBL
                                                    ArrayData(1, CountArrayData) = dataField
                                                    ArrayData(2, CountArrayData) = data
                                                    ArrayData(3, CountArrayData) = pTempData.Description
                                                    CountArrayData += 1
                                                Next j
                                            End If
                                        End If
                                    Next i
                                Else
                                    If psetDataColl.Count > 0 Then
                                        For j As Integer = 0 To psetDataColl.Count - 1
                                            Try
                                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = psetDataColl.Item(j)
                                                Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                Dim data As String = pTempData.GetData() 'значение
                                                ReDim Preserve ArrayData(3, CountArrayData)
                                                ArrayData(0, CountArrayData) = NameTBL
                                                ArrayData(1, CountArrayData) = dataField
                                                ArrayData(2, CountArrayData) = data
                                                ArrayData(3, CountArrayData) = pTempData.Description
                                                CountArrayData += 1
                                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            End Try
                                        Next j
                                    End If
                                End If
                            Catch ex As System.Exception
                            Catch ex1 As System.InvalidCastException
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                End Try
            End Using
        End Using
        FuncReadAllDataSetProperties = True
    End Function
    'функция читает все данные объекта из определенной таблицы (выходные данные - массив - имя словаря, имя поля, значение,дескриптор)
    Public Shared Function FuncReadDataSetPropertiesTbl(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByRef ArrayData As String(,), ByVal NameTBL As String) As Boolean
        FuncReadDataSetPropertiesTbl = False
        Dim CountArrayData As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                If setIds.Count > 0 Then
                    For Each psId As ObjectId In setIds
                        Try
                            Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                            Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                            Dim pTempData As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetData = Nothing
                            Dim NameTempTBL As String = "NoN"
                            NameTempTBL = pset.PropertySetDefinitionName
                            If NameTBL Like NameTempTBL Then
                                If psetDataColl.Count > 0 Then
                                    For j As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(j)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            ReDim Preserve ArrayData(3, CountArrayData)
                                            ArrayData(0, CountArrayData) = pset.PropertySetDefinitionName
                                            ArrayData(1, CountArrayData) = dataField
                                            ArrayData(2, CountArrayData) = data
                                            ArrayData(3, CountArrayData) = pTempData.Description
                                            CountArrayData += 1
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next j
                                End If
                            End If
                        Catch ex1 As System.InvalidCastException
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            End Using
        End Using
        FuncReadDataSetPropertiesTbl = True
    End Function
    'функция читает все данные объекта из определенной таблицы (выходные данные словарь)
    Public Shared Function FuncReadDataIoDictionary(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal NameTBL As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                If setIds.Count > 0 Then
                    For Each psId As ObjectId In setIds
                        Try
                            Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                            Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                            Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                            Dim NameTempTBL As String = "NoN"
                            NameTempTBL = pset.PropertySetDefinitionName
                            If NameTBL Like NameTempTBL Then
                                If psetDataColl.Count > 0 Then
                                    For j As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(j)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            If result.ContainsKey(dataField) = False Then
                                                result.Add(dataField, data)
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next j
                                End If
                            End If
                        Catch ex1 As System.InvalidCastException
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            End Using
        End Using
        Return result
    End Function
    'функция читает определенное поле объекта (выходные данные - массив - имя словаря, имя поля, значение,дескриптор)
    Public Shared Function FuncReadDataSetProperties(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal NameFieldStr As String, ByRef ArrayData As String(,)) As Boolean
        FuncReadDataSetProperties = False
        Dim CountArrayData As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                If setIds.Count > 0 Then
                    For Each psId As ObjectId In setIds
                        Try
                            Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                            Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                            Dim pTempData As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetData = Nothing
                            If psetDataColl.Count > 0 Then
                                For i As Integer = 0 To psetDataColl.Count - 1
                                    Try
                                        pTempData = psetDataColl.Item(i)
                                        Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                        Dim data As String = pTempData.GetData() 'значение
                                        If dataField Like NameFieldStr Then
                                            ReDim Preserve ArrayData(3, CountArrayData)
                                            ArrayData(0, CountArrayData) = pset.PropertySetDefinitionName
                                            ArrayData(1, CountArrayData) = dataField
                                            ArrayData(2, CountArrayData) = data
                                            ArrayData(3, CountArrayData) = pTempData.Description
                                            CountArrayData += 1
                                        End If
                                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                    Catch ex As System.Exception
                                    End Try
                                Next i
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            End Using
        End Using
        FuncReadDataSetProperties = True
    End Function
    'функция читает определенное поле объекта (выходные данные значение поля)
    Public Shared Function FuncReadValueSetProperties(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal NameFieldStr As String, Optional NameTLB As String = "") As String
        FuncReadValueSetProperties = ""
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                If IsNothing(dbObj) = False Then
                    If dbObj.Id.IsNull = False Then
                        Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Try
                                    Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                                    If NameTLB.Trim.Length > 0 Then
                                        If pset.PropertySetDefinitionName Like NameTLB Then
                                            Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                            Dim pTempData As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetData = Nothing
                                            If psetDataColl.Count > 0 Then
                                                For i As Integer = 0 To psetDataColl.Count - 1
                                                    Try
                                                        pTempData = psetDataColl.Item(i)
                                                        Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                        Dim data As String = pTempData.GetData() 'значение
                                                        If dataField Like NameFieldStr Then
                                                            If IsNothing(data) = True Then
                                                                data = ""
                                                            End If
                                                            FuncReadValueSetProperties = data
                                                            Exit Function
                                                        End If
                                                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                    Catch ex As System.Exception
                                                    End Try
                                                Next i
                                            End If
                                        End If
                                    Else
                                        Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                        Dim pTempData As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetData = Nothing
                                        If psetDataColl.Count > 0 Then
                                            For i As Integer = 0 To psetDataColl.Count - 1
                                                Try
                                                    pTempData = psetDataColl.Item(i)
                                                    Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                    Dim data As String = pTempData.GetData() 'значение
                                                    If dataField Like NameFieldStr Then
                                                        If data.Trim.Length > 0 Then
                                                            If IsNothing(data) = True Then
                                                                data = ""
                                                            End If
                                                            FuncReadValueSetProperties = data
                                                            Exit Function
                                                        End If
                                                    End If
                                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                Catch ex As System.Exception
                                                End Try
                                            Next i
                                        End If
                                    End If
                                Catch ex As System.Exception
                                End Try
                            Next
                        End If
                    End If
                End If
            End Using
        End Using
    End Function
    'функция перезаписывает все данные объекта (входные данные - массив - имя словаря, имя поля, значение)
    Public Shared Function FuncWriteAllDataSetProperties(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByRef ArrayData As String(,), Optional indexField As Integer = 1, Optional indexValue As Integer = 2) As Boolean
        FuncWriteAllDataSetProperties = False
        If IsArray(ArrayData) = False Then Exit Function
        If ArrayData.GetUpperBound(0) < 2 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    If dbObj.IsWriteEnabled = False Then
                        dbObj.UpgradeOpen()
                    End If
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                If psetDataColl.Count > 0 Then
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                                Dim NameField As String = ArrayData(indexField, j)
                                                If dataField Like NameField Then
                                                    If pset.PropertySetDefinitionName Like ArrayData(0, j) Then
                                                        Dim tempData As String = ArrayData(indexValue, j)
                                                        If pTempData.DataType = Autodesk.AEC.PropertyData.DataType.Real Then
                                                            Dim RealData As Double = 0
                                                            If IsNothing(tempData) = True Then
                                                                RealData = 0
                                                            ElseIf tempData.Trim.Length = 0 Then
                                                                RealData = 0
                                                            Else
                                                                If IsNumeric(tempData) = True Then
                                                                    RealData = Val(tempData)
                                                                Else
                                                                    RealData = 0
                                                                End If
                                                            End If
                                                            pTempData.SetData(RealData)
                                                        Else
                                                            If IsNothing(tempData) = True Then
                                                                tempData = ""
                                                            End If
                                                            pTempData.SetData(tempData)
                                                        End If
                                                    End If
                                                End If
                                            Next j
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
        FuncWriteAllDataSetProperties = True
    End Function
    'функция перезаписывает все данные объекта (входные данные класс StructureElements)
    Public Shared Function FuncWriteDataSetByStructureElement(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal dataElement As CadastralTools.StructureElement) As Boolean
        FuncWriteDataSetByStructureElement = False
        If IsNothing(dataElement) = True Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim nameTablePs As String = dataElement.TablePS
        Dim arrayData As String(,) = {}

        ReDim Preserve arrayData(2, 2)
        arrayData(0, 0) = nameTablePs
        arrayData(1, 0) = "LABEL"
        arrayData(2, 0) = dataElement.Label

        arrayData(0, 1) = nameTablePs
        arrayData(1, 1) = "LOCALID"
        arrayData(2, 1) = dataElement.IdElement

        arrayData(0, 2) = nameTablePs
        arrayData(1, 2) = "GlobalID"
        arrayData(2, 2) = dataElement.IdStructure

        Dim keyParam As String = dataElement.KeyParameter
        Dim dictParam As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(keyParam)
        If dictParam.Count > 0 Then
            Dim countArrayData As Integer = arrayData.GetUpperBound(1) + 1
            For i As Integer = 0 To dictParam.Count - 1
                Dim value As Object = dictParam.ElementAt(i).Value
                If TypeOf value Is String Or TypeOf value Is Boolean Or TypeOf value Is Integer Or TypeOf value Is Double Then
                    ReDim Preserve arrayData(2, countArrayData)
                    arrayData(0, countArrayData) = nameTablePs
                    arrayData(1, countArrayData) = dictParam.ElementAt(i).Key
                    arrayData(2, countArrayData) = value
                    countArrayData += 1
                End If
            Next i
        End If

        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    If dbObj.IsWriteEnabled = False Then
                        dbObj.UpgradeOpen()
                    End If
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                If psetDataColl.Count > 0 Then
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            For j As Integer = 0 To arrayData.GetUpperBound(1)
                                                Dim NameField As String = arrayData(1, j)
                                                If dataField Like NameField Then
                                                    If pset.PropertySetDefinitionName Like arrayData(0, j) Then
                                                        Dim tempData As String = arrayData(2, j)
                                                        If pTempData.DataType = Autodesk.AEC.PropertyData.DataType.Real Then
                                                            Dim RealData As Double = 0
                                                            If IsNothing(tempData) = True Then
                                                                RealData = 0
                                                            ElseIf tempData.Trim.Length = 0 Then
                                                                RealData = 0
                                                            Else
                                                                If IsNumeric(tempData) = True Then
                                                                    RealData = Val(tempData)
                                                                Else
                                                                    RealData = 0
                                                                End If
                                                            End If
                                                            pTempData.SetData(RealData)
                                                        Else
                                                            If IsNothing(tempData) = True Then
                                                                tempData = ""
                                                            End If
                                                            pTempData.SetData(tempData)
                                                        End If
                                                    End If
                                                End If
                                            Next j
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
        FuncWriteDataSetByStructureElement = True
    End Function
    'функция перезаписывает все данные объекта (входные данные - массив - имя поля, значение)
    Public Shared Function FuncWriteAllDataSetProperties2(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByRef ArrayData As String(,), Optional indexField As Integer = 1, Optional indexValue As Integer = 2) As Boolean
        FuncWriteAllDataSetProperties2 = False
        If IsArray(ArrayData) = False Then Exit Function
        If ArrayData.GetUpperBound(0) < 2 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    If dbObj.IsWriteEnabled = False Then
                        dbObj.UpgradeOpen()
                    End If
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                If psetDataColl.Count > 0 Then
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                                Dim NameField As String = ArrayData(indexField, j)
                                                If dataField Like NameField Then
                                                    Dim tempData As String = ArrayData(indexValue, j)
                                                    If pTempData.DataType = Autodesk.AEC.PropertyData.DataType.Real Then
                                                        Dim RealData As Double = 0
                                                        If IsNothing(tempData) = True Then
                                                            RealData = 0
                                                        ElseIf tempData.Trim.Length = 0 Then
                                                            RealData = 0
                                                        Else
                                                            If IsNumeric(tempData) = True Then
                                                                RealData = Val(tempData)
                                                            Else
                                                                RealData = 0
                                                            End If
                                                        End If
                                                        pTempData.SetData(RealData)
                                                    Else
                                                        If IsNothing(tempData) = True Then
                                                            tempData = ""
                                                        End If
                                                        pTempData.SetData(tempData)
                                                    End If
                                                End If
                                            Next j
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
        FuncWriteAllDataSetProperties2 = True
    End Function
    'функция создает новый словарь Set Properties с пользовательскими данными (входные данные - имя словаря, массив - имя поля, описание)
    Public Shared Function FuncCreatedSetProperties(ByVal propSetDefName As String, ByRef ArrayField As String(,), Optional deskSetDef As String = "") As ObjectId
        FuncCreatedSetProperties = Nothing
        If IsArray(ArrayField) = False Then Exit Function
        If ArrayField.GetUpperBound(0) < 1 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim propSetDef As New Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDefinition
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    propSetDef.SetToStandard(db)
                    propSetDef.SubSetDatabaseDefaults(db)
                    propSetDef.Description = deskSetDef
                    Dim appliedTo = New Specialized.StringCollection()
                    appliedTo.Add(propSetDefName)
                    propSetDef.AppliesToAll = True
                    'добавляем поля
                    For i As Integer = 0 To ArrayField.GetUpperBound(1)
                        Dim NameFieldStr As String = ArrayField(0, i)
                        Dim DeskFieldStr As String = ArrayField(1, i)
                        Dim propDefManual = New Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinition()
                        propDefManual.SetToStandard(db)
                        propDefManual.SubSetDatabaseDefaults(db)
                        propDefManual.Name = NameFieldStr
                        propDefManual.Description = DeskFieldStr
                        propDefManual.DataType = Autodesk.Aec.PropertyData.DataType.Text
                        propDefManual.DefaultData = "-"
                        propSetDef.Definitions.Add(propDefManual)
                        propSetDef.SetDisplayOrder(propDefManual, i + 1)
                    Next i
                    'добавляем словарь в чертеж
                    Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                    If dictPropSetDef.Has(propSetDefName, tr) Then
                        FuncCreatedSetProperties = dictPropSetDef.DictionaryId
                    Else
                        dictPropSetDef.AddNewRecord(propSetDefName, propSetDef)
                        tr.AddNewlyCreatedDBObject(propSetDef, True)
                    End If
                Catch ex As Exception
                    ed.WriteMessage("error in CreatePropSetDef: " & ex.ToString & vbCrLf)
                End Try
                tr.Commit()
            End Using
        End Using
        Return propSetDef.Id
    End Function
    'функция создает новый словарь Set Properties с пользовательскими данными (входные данные - имя словаря, массив - имя поля, значение, описание)
    Public Shared Function FuncCreatedSetProperties2(ByVal propSetDefName As String, ByRef ArrayField As String(,), Optional deskSetDef As String = "") As ObjectId
        FuncCreatedSetProperties2 = Nothing
        If IsArray(ArrayField) = False Then Exit Function
        If ArrayField.GetUpperBound(0) < 1 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim propSetDef As New Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDefinition
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    propSetDef.SetToStandard(db)
                    propSetDef.SubSetDatabaseDefaults(db)
                    propSetDef.Description = deskSetDef
                    Dim appliedTo = New Specialized.StringCollection()
                    appliedTo.Add(propSetDefName)
                    propSetDef.AppliesToAll = True
                    'добавляем поля
                    For i As Integer = 0 To ArrayField.GetUpperBound(1)
                        Dim NameFieldStr As String = ArrayField(0, i)
                        Dim DeskFieldStr As String = ArrayField(2, i)
                        Dim propDefManual = New Autodesk.AEC.PropertyData.DatabaseServices.PropertyDefinition()
                        propDefManual.SetToStandard(db)
                        propDefManual.SubSetDatabaseDefaults(db)
                        propDefManual.Name = NameFieldStr
                        propDefManual.Description = DeskFieldStr
                        propDefManual.DataType = Autodesk.AEC.PropertyData.DataType.Text
                        propDefManual.DefaultData = "-"
                        propSetDef.Definitions.Add(propDefManual)
                        propSetDef.SetDisplayOrder(propDefManual, i + 1)
                    Next i
                    'добавляем словарь в чертеж
                    Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                    If dictPropSetDef.Has(propSetDefName, tr) Then
                        FuncCreatedSetProperties2 = dictPropSetDef.DictionaryId
                    Else
                        dictPropSetDef.AddNewRecord(propSetDefName, propSetDef)
                        tr.AddNewlyCreatedDBObject(propSetDef, True)
                    End If
                Catch ex As Exception
                    ed.WriteMessage("error in CreatePropSetDef: " & ex.ToString & vbCrLf)
                End Try
                tr.Commit()
            End Using
        End Using
        Return propSetDef.Id
    End Function
    'функция добавляет примитиву словарь Set Properties с пользовательскими данными (входные данные - примитив , id словаря)
    Public Shared Function FuncAddDataSetProperties(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal propSetDef As ObjectId) As Boolean
        FuncAddDataSetProperties = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Using tr
                Try
                    If IsDBNull(dbObj) = False Then
                        If dbObj.IsWriteEnabled = False Then
                            dbObj.UpgradeOpen()
                        End If
                        Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.AddPropertySet(dbObj, propSetDef)
                        FuncAddDataSetProperties = True
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
    End Function
    'функция возвращает id словаря по его имени Set Properties 
    Public Shared Function FuncReturnIDDictSetProperties(ByVal NameDict As String) As ObjectId
        FuncReturnIDDictSetProperties = Nothing
        If IsNothing(NameDict) = True Then
            Exit Function
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Using tr
                Try
                    If dictPropSetDef.Has(NameDict, tr) Then
                        FuncReturnIDDictSetProperties = dictPropSetDef.GetAt(NameDict)
                        Exit Function
                    End If
                Catch ex As System.Exception
                    FuncReturnIDDictSetProperties = Nothing
                End Try
            End Using
        End Using
    End Function
    'функция создает словарь Properties SET на основе xml
    Public Shared Function FuncCreateTablePSByXML(ByVal puchTemplateXML As String, ByVal NameTablePS As String) As ObjectId
        FuncCreateTablePSByXML = Nothing
        If IO.File.Exists(puchTemplateXML) = False Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Dim propSetDef As New Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDefinition
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim dictPropSetDefTemp = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                Try
                    If dictPropSetDefTemp.Has(NameTablePS, tr) Then
                        FuncCreateTablePSByXML = dictPropSetDefTemp.GetAt(NameTablePS)
                        Exit Function
                    End If
                Catch ex As System.Exception
                End Try
                Try
                    Dim reader As XmlTextReader = New XmlTextReader(puchTemplateXML)
                    Dim findUserTablePS As Boolean = False
                    Dim deskTablePS As String = ""
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "PropertyTable" Then
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            If reader.Name Like "TableName" Then
                                                If reader.Value Like NameTablePS Then
                                                    findUserTablePS = True
                                                End If
                                            ElseIf reader.Name Like "DeskTable" Then
                                                deskTablePS = reader.Value
                                            End If
                                        End While
                                        If findUserTablePS = True Then
                                            propSetDef.SetToStandard(db)
                                            propSetDef.SubSetDatabaseDefaults(db)
                                            propSetDef.Description = deskTablePS
                                            Dim appliedTo = New Specialized.StringCollection()
                                            appliedTo.Add(NameTablePS)
                                            propSetDef.AppliesToAll = True
                                            'propSetDef.SetAppliesToFilter(appliedTo, False)
                                        End If
                                    End If
                                ElseIf NameBlock Like "Field" And findUserTablePS = True Then
                                    If findUserTablePS = True Then
                                        Dim NameField As String = ""
                                        Dim DeskField As String = ""
                                        Dim DefaultVal As String = ""
                                        Dim DataType As Integer = 3
                                        Dim drawOrd As Integer = 0
                                        While (reader.MoveToNextAttribute())
                                            If reader.Name Like "Name" Then
                                                NameField = reader.Value
                                            ElseIf reader.Name Like "Description" Then
                                                DeskField = reader.Value
                                            ElseIf reader.Name Like "DefaultData" Then
                                                DefaultVal = reader.Value
                                            ElseIf reader.Name Like "TypeField" Then
                                                If IsNumeric(reader.Value) = True Then
                                                    DataType = Val(reader.Value)
                                                End If
                                            ElseIf reader.Name Like "DisplayOrder" Then
                                                If IsNumeric(reader.Value) = True Then
                                                    drawOrd = Val(reader.Value)
                                                End If
                                            End If
                                        End While
                                        If Not (NameField Like "") Then
                                            Dim propDefManual = New Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinition()
                                            propDefManual.SetToStandard(db)
                                            propDefManual.SubSetDatabaseDefaults(db)
                                            propDefManual.Name = NameField
                                            propDefManual.Description = DeskField
                                            If DataType = 1 Then
                                                propDefManual.DataType = Autodesk.Aec.PropertyData.DataType.Real
                                            ElseIf DataType = 0 Then
                                                propDefManual.DataType = Autodesk.Aec.PropertyData.DataType.Integer
                                            Else
                                                propDefManual.DataType = Autodesk.Aec.PropertyData.DataType.Text
                                            End If
                                            propDefManual.DefaultData = DefaultVal
                                            propSetDef.Definitions.Add(propDefManual)
                                            propSetDef.SetDisplayOrder(propDefManual, drawOrd)
                                        End If
                                    End If
                                End If
                            Case XmlNodeType.EndElement
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "PropertyTable" Then
                                    If findUserTablePS = True Then
                                        'добавляем словарь в чертеж
                                        Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                                        If dictPropSetDef.Has(NameTablePS, tr) Then
                                            FuncCreateTablePSByXML = dictPropSetDef.DictionaryId
                                            Exit While
                                        End If
                                        dictPropSetDef.AddNewRecord(NameTablePS, propSetDef)
                                        tr.AddNewlyCreatedDBObject(propSetDef, True)
                                        Exit While
                                    End If
                                End If
                        End Select
                    End While
                    tr.Commit()
                Catch ex As System.Exception
                End Try
            End Using
        End Using
        Return propSetDef.Id
    End Function
    'функция возвращает все словари подключенные к объекту (выходные данные - массив - имя словаря)
    Public Shared Function FuncReadAllDataSetTables(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByRef ArrayData As String(), Optional ByVal ArrayTBLXM As String(,) = Nothing) As Boolean
        FuncReadAllDataSetTables = False
        Dim CountArrayData As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim NameTbl As String = pset.PropertySetDefinitionName
                                If IsArray(ArrayTBLXM) = True Then
                                    For i As Integer = 0 To ArrayTBLXM.GetUpperBound(1)
                                        If ArrayTBLXM(0, i) Like NameTbl Then
                                            ReDim Preserve ArrayData(CountArrayData)
                                            ArrayData(CountArrayData) = pset.PropertySetDefinitionName
                                            CountArrayData += 1
                                            Exit For
                                        End If
                                    Next i
                                Else
                                    ReDim Preserve ArrayData(CountArrayData)
                                    ArrayData(CountArrayData) = pset.PropertySetDefinitionName
                                    CountArrayData += 1
                                End If
                            Catch ex1 As System.InvalidCastException
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As Exception
                End Try
            End Using
        End Using
        FuncReadAllDataSetTables = True
    End Function
    'функция проверяет подключен ли словарь к объекту
    Public Shared Function FuncFindPSetTable(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal NameTBL As String) As Boolean
        FuncFindPSetTable = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
            If setIds.Count > 0 Then
                For Each psId As ObjectId In setIds
                    Try
                        Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                        Dim NameTempTbl As String = pset.PropertySetDefinitionName
                        If NameTempTbl Like NameTBL Then
                            FuncFindPSetTable = True
                            Exit Function
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
            End If
        End Using
    End Function
    'функция удаляет словарь из объекта
    Public Shared Function FuncDeleteDataSetTables(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal TableName As String, Optional ArrayTBLXM As String(,) = Nothing) As Boolean
        FuncDeleteDataSetTables = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Dim boolFindXml As Boolean = False
        If IsArray(ArrayTBLXM) = False Then
            Exit Function
        Else
            For i As Integer = 0 To ArrayTBLXM.GetUpperBound(1)
                Dim TempName As String = ArrayTBLXM(0, i)
                If TempName Like TableName Then
                    boolFindXml = True
                    Exit For
                End If
            Next i
        End If
        If boolFindXml = True Then
            Using loc
                Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
                Using tr
                    Try
                        Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Try
                                    Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, True, False)
                                    Dim NamepSet As String = pset.PropertySetDefinitionName
                                    If NamepSet Like TableName Then
                                        Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.RemovePropertySet(dbObj, psId)
                                    End If
                                Catch ex As System.Exception
                                End Try
                            Next
                        End If
                    Catch ex As System.Exception
                    End Try
                    tr.Commit()
                End Using
            End Using
            FuncDeleteDataSetTables = True
        End If
    End Function
    'функция удаляет массив словарей из объекта
    Public Shared Function FuncDeleteDataSetTables1(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal TableName As String()) As Boolean
        FuncDeleteDataSetTables1 = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        If IsArray(TableName) = False Then
            Exit Function
        End If
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim NamepSet As String = pset.PropertySetDefinitionName
                                For i As Integer = 0 To TableName.Length - 1
                                    If NamepSet Like TableName(i) Then
                                        Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.RemovePropertySet(dbObj, psId)
                                    End If
                                Next i
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As Exception
                End Try
                tr.Commit()
                loc.Dispose()
            End Using
        End Using
        FuncDeleteDataSetTables1 = True
    End Function
    'функция удаляет словарь из объекта
    Public Shared Function FuncDeleteDataSetTables1(ByVal dbObj As Autodesk.AutoCAD.DatabaseServices.DBObject, ByVal propSetDef As ObjectId) As Boolean
        FuncDeleteDataSetTables1 = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        If IsNothing(propSetDef) = True Then
            Exit Function
        End If
        If propSetDef.IsNull Then
            Exit Function
        End If
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Using tr
                Try
                    Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.RemovePropertySet(dbObj, propSetDef)
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
        FuncDeleteDataSetTables1 = True
    End Function
    'функция возвращает массив с именами таблиц в чертеже
    Public Shared Function FuncReturnArrayNamesDictGWG(ByRef ArrayNameDict As String()) As Integer
        FuncReturnArrayNamesDictGWG = Nothing
        Dim CountArrayNameDict As Integer = 0
        If IsArray(ArrayNameDict) = True Then
            CountArrayNameDict = ArrayNameDict.Length
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Using tr
                Try
                    Dim strNameDictCollection As Specialized.StringCollection = dictPropSetDef.NamesInUse
                    If strNameDictCollection.Count > 0 Then
                        For i As Integer = 0 To strNameDictCollection.Count - 1
                            ReDim Preserve ArrayNameDict(CountArrayNameDict)
                            ArrayNameDict(CountArrayNameDict) = strNameDictCollection.Item(i)
                            CountArrayNameDict += 1
                        Next
                    End If
                    FuncReturnArrayNamesDictGWG = CountArrayNameDict
                Catch ex As System.Exception
                    FuncReturnArrayNamesDictGWG = 0
                End Try
            End Using
        End Using
    End Function
    'функция перезаписывает все данные объекта (входные данные - массив - имя словаря, имя поля, значение)
    Public Shared Function FuncWriteAllDataSetProperties1(ByVal dbObj As Autodesk.Civil.DatabaseServices.Entity, ByRef ArrayData As String(,)) As Boolean
        FuncWriteAllDataSetProperties1 = False
        If IsArray(ArrayData) = False Then Exit Function
        If ArrayData.GetUpperBound(0) < 2 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    If dbObj.IsWriteEnabled = False Then
                        dbObj.UpgradeOpen()
                    End If
                    Dim setIds As ObjectIdCollection = Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(dbObj)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, False, False)
                                Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim pTempData As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetData = Nothing
                                If psetDataColl.Count > 0 Then
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            For j As Integer = 0 To ArrayData.GetUpperBound(1)
                                                Dim NAMEF1 As String = ArrayData(1, j)
                                                If dataField Like NAMEF1 Then
                                                    If pset.PropertySetDefinitionName Like ArrayData(0, j) Then
                                                        Dim tempData As String = ArrayData(2, j)
                                                        If pTempData.DataType = Autodesk.Aec.PropertyData.DataType.Real Then
                                                            Dim RealData As Double = 0
                                                            If IsNothing(tempData) = True Then
                                                                RealData = 0
                                                            ElseIf tempData.Trim.Length = 0 Then
                                                                RealData = 0
                                                            Else
                                                                If IsNumeric(tempData) = True Then
                                                                    RealData = Val(tempData)
                                                                Else
                                                                    RealData = 0
                                                                End If
                                                            End If
                                                            pTempData.SetData(RealData)
                                                        Else
                                                            If IsNothing(tempData) = True Then
                                                                tempData = ""
                                                            End If
                                                            pTempData.SetData(tempData)
                                                        End If
                                                    End If
                                                End If
                                            Next j
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
        FuncWriteAllDataSetProperties1 = True
    End Function
    'функция добавляет примитиву словарь Set Properties с пользовательскими данными (входные данные - примитив , id словаря)
    Public Shared Function FuncAddDataSetProperties1(ByVal dbObj As Autodesk.Civil.DatabaseServices.Entity, ByVal propSetDef As ObjectId) As Boolean
        FuncAddDataSetProperties1 = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            Using tr
                Try
                    If IsDBNull(dbObj) = False Then
                        If dbObj.IsWriteEnabled = False Then
                            dbObj.UpgradeOpen()
                        End If
                        Autodesk.Aec.PropertyData.DatabaseServices.PropertyDataServices.AddPropertySet(dbObj, propSetDef)
                        FuncAddDataSetProperties1 = True
                    End If
                Catch ex As System.Exception
                End Try
                tr.Commit()
            End Using
        End Using
    End Function
    'функция возвращает значение поля таблицы
    Public Shared Function FuncReadValueFieldSetTables(ByVal NameTablePS As String, ByVal NameField As String) As String
        FuncReadValueFieldSetTables = ""
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                    Try
                        If dictPropSetDef.Has(NameTablePS, tr) Then
                            Dim idTablePS As ObjectId = dictPropSetDef.GetAt(NameTablePS)
                            Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDefinition = tr.GetObject(idTablePS, OpenMode.ForRead, False, False)
                            Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinitionCollection = pset.Definitions
                            For Each field As Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinition In psetDataColl
                                Dim TempNameField As String = field.Name
                                Dim data As String = field.DefaultData
                                If TempNameField Like NameField Then
                                    FuncReadValueFieldSetTables = data
                                    Exit For
                                End If
                            Next
                        End If
                    Catch ex As System.Exception
                    End Try
                Catch ex As Exception
                End Try
            End Using
        End Using
    End Function
    'функция возвращает имя таблицы, имена полей значение  таблицы PS
    Public Shared Function FuncReadAllValueFieldSetTables(ByVal NameTablePS As String, ByRef ArrayFields As String(,)) As Boolean
        FuncReadAllValueFieldSetTables = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim loc As Autodesk.AutoCAD.ApplicationServices.DocumentLock = acDoc.LockDocument
        Using loc
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Try
                    Dim dictPropSetDef = New Autodesk.Aec.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                    Try
                        If dictPropSetDef.Has(NameTablePS, tr) Then
                            Dim idTablePS As ObjectId = dictPropSetDef.GetAt(NameTablePS)
                            Dim pset As Autodesk.Aec.PropertyData.DatabaseServices.PropertySetDefinition = tr.GetObject(idTablePS, OpenMode.ForRead, False, False)
                            Dim psetDataColl As Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinitionCollection = pset.Definitions
                            Dim countArrayFields As Integer = 0
                            For Each field As Autodesk.Aec.PropertyData.DatabaseServices.PropertyDefinition In psetDataColl
                                Dim TempNameField As String = field.Name
                                Dim data As String = field.DefaultData
                                Dim TypeField As Autodesk.Aec.PropertyData.DataType = field.DataType
                                ReDim Preserve ArrayFields(6, countArrayFields)
                                ArrayFields(0, countArrayFields) = NameTablePS
                                ArrayFields(1, countArrayFields) = TempNameField
                                ArrayFields(2, countArrayFields) = data
                                ArrayFields(3, countArrayFields) = field.Description
                                ArrayFields(4, countArrayFields) = TypeField
                                ArrayFields(5, countArrayFields) = field.DefaultData
                                ArrayFields(6, countArrayFields) = field.DisplayOrder
                                countArrayFields += 1
                            Next
                            If IsArray(ArrayFields) = True Then
                                If ArrayFields.GetUpperBound(1) > 0 Then
                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayFields, 6)
                                End If
                                FuncReadAllValueFieldSetTables = True
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Catch ex As Exception
                End Try
            End Using
        End Using
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все объекты с указанной таблицы
    Public Shared Function FuncReturnDwgObjectByNameTablePS(ByRef userObjectCollection As DBObjectCollection, ByVal NameTablePS As String, Optional boolShowLayer As Boolean = True) As Boolean
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If boolShowLayer = False Then
                        Dim idLayer As ObjectId = ent.LayerId
                        Dim acLayer As LayerTableRecord = idLayer.GetObject(OpenMode.ForRead)
                        If acLayer.IsHidden = True Then
                            Continue For
                        ElseIf acLayer.IsFrozen = True Then
                            Continue For
                        End If
                    End If
                    Dim boolTable As Boolean = FuncFindPSetTable(ent, NameTablePS)
                    If boolTable = True Then
                        userObjectCollection.Add(ent)
                    End If
                Next
            Catch ex As System.Exception
            End Try
        End Using
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все объекты с указанных таблиц в словарь
    Public Shared Function FuncReturnDwgObjectByNameTablesPS(ByRef userDictionaryObject As Dictionary(Of String, DBObjectCollection), ByVal NameTablePS As String(), Optional boolShowLayer As Boolean = True) As Boolean
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If boolShowLayer = False Then
                        Dim idLayer As ObjectId = ent.LayerId
                        Dim acLayer As LayerTableRecord = idLayer.GetObject(OpenMode.ForRead)
                        If acLayer.IsHidden = True Then
                            Continue For
                        ElseIf acLayer.IsFrozen = True Then
                            Continue For
                        End If
                    End If
                    For i As Integer = 0 To NameTablePS.Length - 1
                        Dim boolTable As Boolean = FuncFindPSetTable(ent, NameTablePS(i))
                        If boolTable = True Then
                            If userDictionaryObject.Count > 0 Then
                                Dim boolFindObj As Boolean = userDictionaryObject.ContainsKey(NameTablePS(i))
                                If boolFindObj = True Then
                                    Dim arrayObj As DBObjectCollection = userDictionaryObject.Item(NameTablePS(i))
                                    arrayObj.Add(ent)
                                    Exit For
                                End If
                            End If
                            Dim entcol As DBObjectCollection = New DBObjectCollection
                            entcol.Add(ent)
                            userDictionaryObject.Add(NameTablePS(i), entcol)
                        End If
                    Next i
                Next
            Catch ex As System.Exception
            End Try
        End Using
        Return True
    End Function


    <CommandMethod("LoadTablePropertySetByXML")>
    Public Sub LoadTablePropertySetByXML()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
        Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Введите имя таблицы: ")
        pStrOpts.AllowSpaces = True
        Dim pStrRes As PromptResult = acDoc.Editor.GetString(pStrOpts)
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. Путь к изображениям объектов
        Dim PutchFile As String = ""
        Dim NameFile As String = "ClassObject.xml"
        If IO.Directory.Exists(parentPatch2 & "\FileResources\MapClassification\") Then
            PutchFile = parentPatch2 & "\FileResources\MapClassification\"
        Else
            MsgBox("Не удалось найти каталог. Макро прервано!!!")
            Exit Sub
        End If
        If IO.File.Exists(PutchFile & NameFile) Then
            Dim IDLoadTBL As ObjectId = FuncCreateTablePSByXML(PutchFile, pStrRes.StringResult.Trim)
            If IsNothing(IDLoadTBL) = True Then
                MsgBox("Не удалось загрузить таблицу PropertySet " & pStrRes.StringResult)
            ElseIf IDLoadTBL.IsNull = True Then
                MsgBox("Не удалось загрузить таблицу PropertySet " & pStrRes.StringResult)
            Else
                MsgBox("Таблица PropertySet успешно загружена." & pStrRes.StringResult)
            End If
        Else
            MsgBox("Не удалось найти файл ClassObject.xml. Макро прервано!!!")
            Exit Sub
        End If
    End Sub
End Class
