Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports System.IO
Imports System.Xml
Imports System.Xaml
Imports NetTopologySuite
Imports CadastralTools.StructureElement

Public Class RosreestrDataImport
    'функция копирует все строки с кодами от росреестра в один массив (0-desk, 1-name, 2-layer, 3-color, 4-width, 5-M, 6-GWidth, 7-PS, 8-Note)? Для линий
    Public Shared Function FuncListLineObjectDataTableDestDWG(ByVal PutshDraw As String, ByVal CategoryNameTable As String, ByRef ArrayTableTypeObject(,) As String, Optional NameTable As String = "") As Integer
        FuncListLineObjectDataTableDestDWG = 0
        Dim CountArrayTableTypeObject As Integer = 0
        If IsArray(ArrayTableTypeObject) = True Then
            CountArrayTableTypeObject = ArrayTableTypeObject.GetUpperBound(1)
        End If
        If File.Exists(PutshDraw) = False Then Return CountArrayTableTypeObject
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return 0
        End Try
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                For Each id As ObjectId In sourceBtr
                    Try
                        If IsNothing(id) = True Then Continue For
                        If IsDBNull(id) = True Then Continue For
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                        If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                            Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                            If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                                Dim row1 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(0)
                                Dim CategoryName As String = tbl.Cells(0, 0).Value
                                Dim NameTbl As String = tbl.Cells(1, 0).Value
                                If NameTable.Trim.Length > 0 Then
                                    If Not (NameTable.Trim Like NameTbl.Trim) Then Continue For
                                End If
                                If CategoryName.Trim Like CategoryNameTable Then
                                    'считываем заголовок
                                    Dim PosDesk As Integer = -1
                                    Dim PosName As Integer = -1
                                    Dim PosLayer As Integer = -1
                                    Dim PosType As Integer = -1
                                    Dim PosColor As Integer = -1
                                    Dim PosWLine As Integer = -1
                                    Dim PosMLine As Integer = -1
                                    Dim PosGLine As Integer = -1
                                    Dim PosPS As Integer = -1
                                    Dim PosNote As Integer = -1
                                    For i As Integer = 0 To tbl.Columns.Count - 1
                                        Dim FieldName As String = tbl.Cells(2, i).Value
                                        If IsNothing(FieldName) Then Continue For
                                        FieldName = FieldName.Trim
                                        If FieldName.Length = 0 Then Continue For
                                        If FieldName Like "Description" Then
                                            PosDesk = i
                                        ElseIf FieldName Like "Name" Then
                                            PosName = i
                                        ElseIf FieldName Like "Layer" Then
                                            PosLayer = i
                                        ElseIf FieldName Like "Type" Then
                                            PosType = i
                                        ElseIf FieldName Like "Color" Then
                                            PosColor = i
                                        ElseIf FieldName Like "LineWidth" Then
                                            PosWLine = i
                                        ElseIf FieldName Like "LineScale" Then
                                            PosMLine = i
                                        ElseIf FieldName Like "LineGWidth" Then
                                            PosGLine = i
                                        ElseIf FieldName Like "PS" Then
                                            PosPS = i
                                        ElseIf FieldName Like "Note" Then
                                            PosNote = i
                                        End If
                                    Next i
                                    '===================================================================================================
                                    'считываем все строки в массив включая заголовок
                                    For i As Integer = 3 To tbl.Rows.Count - 1
                                        Dim ValStrTemp As String = tbl.Cells(i, PosType).Value
                                        If IsNothing(ValStrTemp) = False Then
                                            If ValStrTemp.Trim.Length > 0 Then
                                                If Not (ValStrTemp Like "Полилиния") Then
                                                    Continue For
                                                End If
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            Continue For
                                        End If
                                        If PosNote > -1 Then
                                            ValStrTemp = tbl.Cells(i, PosNote).Value
                                            If IsNothing(ValStrTemp) = False Then
                                                If ValStrTemp.Trim.Length > 0 Then
                                                    If ValStrTemp.IndexOf("Code") > -1 Then
                                                        '(0-desk, 1-name, 2-layer, 3-color, 4-width, 5-M, 6-GWidth, 7-PS, 8-Note)
                                                        ReDim Preserve ArrayTableTypeObject(8, CountArrayTableTypeObject)
                                                        ArrayTableTypeObject(8, CountArrayTableTypeObject) = ValStrTemp
                                                        If PosDesk > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosDesk).Value
                                                            ArrayTableTypeObject(0, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosName > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosName).Value
                                                            ArrayTableTypeObject(1, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosLayer > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosLayer).Value
                                                            ArrayTableTypeObject(2, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosColor > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosColor).Value
                                                            ArrayTableTypeObject(3, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosWLine > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosWLine).Value
                                                            ArrayTableTypeObject(4, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosMLine > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosMLine).Value
                                                            ArrayTableTypeObject(5, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosGLine > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosGLine).Value
                                                            ArrayTableTypeObject(6, CountArrayTableTypeObject) = ValStrTemp
                                                        End If

                                                        If PosPS > -1 Then
                                                            ValStrTemp = tbl.Cells(i, PosPS).Value
                                                            ArrayTableTypeObject(7, CountArrayTableTypeObject) = ValStrTemp
                                                        End If
                                                        CountArrayTableTypeObject += 1
                                                    Else
                                                        Continue For
                                                    End If
                                                Else
                                                    Continue For
                                                End If
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            Exit For
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                Next
            End Using
        End Using
        FuncListLineObjectDataTableDestDWG = CountArrayTableTypeObject
    End Function
    '=========================================================================================================================================================
    'функция копирует линейные объекты из файла-шаблона по заданному массиву
    Public Shared Function FuncCopyLineStylesDestDraw(ByVal PutshDraw As String, ByVal ArrayNameObject As String(,), ByVal boolReplase As Boolean) As Boolean
        FuncCopyLineStylesDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return 0
        End Try
        If IsArray(ArrayNameObject) = True Then
            If ArrayNameObject.GetUpperBound(1) > 0 Then
                Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
                Using locker
                    Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                    Using sourceTr
                        Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
                        Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                        Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                        For i As Integer = 0 To ArrayNameObject.GetUpperBound(1)
                            Try
                                'копируем тип линии
                                Dim ValStr As String = ArrayNameObject(1, i).Trim
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If acLineTypTbl.Has(ValStr) = True Then
                                            Dim IDTypeLine As ObjectId = acLineTypTbl.Item(ValStr)
                                            Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                            ObjColID.Add(IDTypeLine)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim db1 As Database = New Database(False, False)
                                            Dim destDb1 As Database = acDoc.Database
                                            Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                                            Dim iMap1 As IdMapping = New IdMapping()
                                            If boolReplase = True Then
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                            Else
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                            End If
                                        End If
                                    End If
                                End If
                                'копируем слой
                                ValStr = ArrayNameObject(2, i).Trim
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If acLyrTbl.Has(ValStr) = True Then
                                            Dim IDLayer As ObjectId = acLyrTbl.Item(ValStr)
                                            Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                            ObjColID.Add(IDLayer)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb1 As Database = acDoc.Database
                                            Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                            Dim iMap1 As IdMapping = New IdMapping()
                                            If boolReplase = False Then
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                            Else
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                            End If
                                        End If
                                    End If
                                End If
                                'копируем таблицу PS
                                ValStr = ArrayNameObject(7, i).Trim
                                If IsNothing(ValStr) = False Then
                                    If ValStr.Trim.Length > 0 Then
                                        If dictPropSetDef.Has(ValStr, sourceTr) = True Then
                                            Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(ValStr)
                                            Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                                            CollDictPropSet.Add(IDPSTable)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb As Database = acDoc.Database
                                            Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                                            Using destTr
                                                Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                                                If dictPropSetDefTemp.Has(ValStr, destTr) = False Then
                                                    Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                                                    Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                                                    helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                                                End If
                                                destTr.Commit()
                                            End Using
                                        End If
                                    End If
                                End If
                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            End Try
                        Next
                        sourceTr.Commit()
                    End Using
                End Using
            End If
        End If
        FuncCopyLineStylesDestDraw = True
    End Function
    '===================================================================================================================================================================
    'функция копирует все строки с кодами от росреестра в один массив (0-desk, 1-name, 2-layer, 3-color, 4-размер атрибута, 5-M, 6-WidthLine, 7-PS, 8-Note) Для блоков
    Public Shared Function FuncListBlkObjectDataTableDestDWG(ByVal PutshDraw As String, ByVal CategoryNameTable As String, ByRef ArrayTableTypeObject(,) As String, Optional NameTable As String = "") As Integer
        FuncListBlkObjectDataTableDestDWG = 0
        Dim CountArrayTableTypeObject As Integer = 0
        If IsArray(ArrayTableTypeObject) = True Then
            CountArrayTableTypeObject = ArrayTableTypeObject.GetUpperBound(1)
        End If
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return 0
        End Try
        Try
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                For Each id As ObjectId In sourceBtr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                        Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                        If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                            Dim row1 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(0)
                            Dim CategoryName As String = tbl.Cells(0, 0).Value
                            Dim NameTbl As String = tbl.Cells(1, 0).Value
                            If NameTable.Trim.Length > 0 Then
                                If Not (NameTable.Trim Like NameTbl.Trim) Then Continue For
                            End If
                            If CategoryName.Trim Like CategoryNameTable Then
                                'считываем заголовок
                                Dim PosDesk As Integer = -1
                                Dim PosName As Integer = -1
                                Dim PosLayer As Integer = -1
                                Dim PosType As Integer = -1
                                Dim PosColor As Integer = -1
                                Dim PosWLine As Integer = -1
                                Dim PosMBlk As Integer = -1
                                Dim PosSizeAttr As Integer = -1
                                Dim PosPS As Integer = -1
                                Dim PosNote As Integer = -1
                                For i As Integer = 0 To tbl.Columns.Count - 1
                                    Dim FieldName As String = tbl.Cells(2, i).Value
                                    If IsNothing(FieldName) Then Continue For
                                    FieldName = FieldName.Trim
                                    If FieldName.Length = 0 Then Continue For
                                    If FieldName Like "Description" Then
                                        PosDesk = i
                                    ElseIf FieldName Like "Name" Then
                                        PosName = i
                                    ElseIf FieldName Like "Layer" Then
                                        PosLayer = i
                                    ElseIf FieldName Like "Type" Then
                                        PosType = i
                                    ElseIf FieldName Like "Color" Then
                                        PosColor = i
                                    ElseIf FieldName Like "LineWidth" Then
                                        PosWLine = i
                                    ElseIf FieldName Like "Scale" Then
                                        PosMBlk = i
                                    ElseIf FieldName Like "AttributeHeight" Then
                                        PosSizeAttr = i
                                    ElseIf FieldName Like "PS" Then
                                        PosPS = i
                                    ElseIf FieldName Like "Note" Then
                                        PosNote = i
                                    End If
                                Next i
                                '===================================================================================================
                                'считываем все строки в массив включая заголовок
                                For i As Integer = 3 To tbl.Rows.Count - 1
                                    Dim ValStrTemp As String = tbl.Cells(i, PosType).Value
                                    If IsNothing(ValStrTemp) = False Then
                                        If ValStrTemp.Trim.Length > 0 Then
                                            If Not (ValStrTemp Like "Блок") Then
                                                Continue For
                                            End If
                                        Else
                                            Continue For
                                        End If
                                    Else
                                        Continue For
                                    End If
                                    If PosNote > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosNote).Value
                                        If IsNothing(ValStrTemp) = False Then
                                            If ValStrTemp.Trim.Length > 0 Then
                                                If ValStrTemp.IndexOf("Code") > -1 Then
                                                    ReDim Preserve ArrayTableTypeObject(8, CountArrayTableTypeObject)
                                                    ArrayTableTypeObject(8, CountArrayTableTypeObject) = ValStrTemp
                                                End If
                                            End If
                                        End If
                                    End If
                                    '(0-desk, 1-name, 2-layer, 3-color, 4-размер атрибута, 5-M, 6-WidthLine, 7-PS, 8-Note)
                                    If PosDesk > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosDesk).Value
                                        ArrayTableTypeObject(0, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosName > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosName).Value
                                        ArrayTableTypeObject(1, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosLayer > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosLayer).Value
                                        ArrayTableTypeObject(2, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosColor > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosColor).Value
                                        ArrayTableTypeObject(3, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosSizeAttr > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosSizeAttr).Value
                                        ArrayTableTypeObject(4, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosMBlk > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosMBlk).Value
                                        ArrayTableTypeObject(5, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosWLine > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosWLine).Value
                                        ArrayTableTypeObject(6, CountArrayTableTypeObject) = ValStrTemp
                                    End If

                                    If PosPS > -1 Then
                                        ValStrTemp = tbl.Cells(i, PosPS).Value
                                        ArrayTableTypeObject(7, CountArrayTableTypeObject) = ValStrTemp
                                    End If
                                    CountArrayTableTypeObject += 1
                                Next i
                            End If
                        End If
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        FuncListBlkObjectDataTableDestDWG = CountArrayTableTypeObject
        sourceDb.CloseInput(True)
    End Function
    '=========================================================================================================================================================
    'функция копирует блоки из файла-шаблона по заданному массиву
    Public Shared Function FuncCopyBlkStylesDestDraw(ByVal PutshDraw As String, ByVal ArrayNameObject As String(,), ByVal boolReplase As Boolean) As Boolean
        FuncCopyBlkStylesDestDraw = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return 0
        End Try
        If IsArray(ArrayNameObject) = True Then
            If ArrayNameObject.GetUpperBound(1) > 0 Then
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim acBlkTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = sourceTr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                    Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                    For i As Integer = 0 To ArrayNameObject.GetUpperBound(1)
                        Try
                            'копируем определение блока
                            Dim ValStr As String = ArrayNameObject(1, i).Trim
                            If IsNothing(ValStr) = False Then
                                If ValStr.Trim.Length > 0 Then
                                    If acBlkTbl.Has(ValStr) = True Then
                                        Dim IDBlk As ObjectId = acBlkTbl.Item(ValStr)
                                        Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                        ObjColID.Add(IDBlk)
                                        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                        Dim db1 As Database = New Database(False, False)
                                        Dim destDb1 As Database = acDoc.Database
                                        Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                                        Dim iMap1 As IdMapping = New IdMapping()
                                        If boolReplase = True Then
                                            destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                        Else
                                            destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                        End If
                                    End If
                                End If
                            End If
                            'копируем слой
                            ValStr = ArrayNameObject(2, i).Trim
                            If IsNothing(ValStr) = False Then
                                If ValStr.Trim.Length > 0 Then
                                    If acLyrTbl.Has(ValStr) = True Then
                                        Dim IDLayer As ObjectId = acLyrTbl.Item(ValStr)
                                        Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                        ObjColID.Add(IDLayer)
                                        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                        Dim destDb1 As Database = acDoc.Database
                                        Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                        Dim iMap1 As IdMapping = New IdMapping()
                                        If boolReplase = False Then
                                            destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                        Else
                                            destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                        End If
                                    End If
                                End If
                            End If
                            'копируем таблицу PS
                            ValStr = ArrayNameObject(7, i).Trim
                            If IsNothing(ValStr) = False Then
                                If ValStr.Trim.Length > 0 Then
                                    If dictPropSetDef.Has(ValStr, sourceTr) = True Then
                                        Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(ValStr)
                                        Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                                        CollDictPropSet.Add(IDPSTable)
                                        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                        Dim destDb As Database = acDoc.Database
                                        Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                                        Using destTr
                                            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                                            If dictPropSetDefTemp.Has(ValStr, destTr) = False Then
                                                Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                                                Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                                                helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                                            End If
                                            destTr.Commit()
                                        End Using
                                    End If
                                End If
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                    sourceTr.Commit()
                End Using
            End If
        End If
        sourceDb.CloseInput(True)
        FuncCopyBlkStylesDestDraw = True
    End Function

    Public Shared Function FuncListObjectDataTableDestDWG2(ByVal PutshDraw As String, ByVal CategoryNameTable As String, ByVal NameTable As String, ByRef ArrayTableTypeObject(,) As String) As Integer
        FuncListObjectDataTableDestDWG2 = 0
        Dim CountArrayTableTypeObject As Integer = 0
        If IsArray(ArrayTableTypeObject) = True Then
            CountArrayTableTypeObject = ArrayTableTypeObject.GetUpperBound(1)
        End If
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Return False
        End Try
        Try
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                For Each id As ObjectId In sourceBtr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.Table Then
                        Dim tbl As Autodesk.AutoCAD.DatabaseServices.Table = ent
                        If tbl.Rows.Count > 3 And tbl.Columns.Count > 2 Then
                            Dim row1 As Autodesk.AutoCAD.DatabaseServices.Row = tbl.Rows.Item(0)
                            Dim CategoryName As String = tbl.Cells(0, 0).Value
                            Dim NameTbl As String = tbl.Cells(1, 0).Value
                            If CategoryName.Trim Like CategoryNameTable Then
                                If NameTable.Trim Like NameTbl.Trim Then
                                    '===================================================================================================
                                    'считываем все строки в массив включая заголовок
                                    Dim CountColumns As Integer = tbl.Columns.Count - 1
                                    For i As Integer = 2 To tbl.Rows.Count - 1
                                        Dim ValStrTemp As String = tbl.Cells(i, 1).Value
                                        If IsNothing(ValStrTemp) = True Then
                                            Continue For
                                        ElseIf ValStrTemp.Trim.Length = 0 Then
                                            Continue For
                                        End If
                                        ReDim Preserve ArrayTableTypeObject(CountColumns, CountArrayTableTypeObject)
                                        For j As Integer = 0 To tbl.Columns.Count - 1
                                            Dim ValStr As String = tbl.Cells(i, j).Value
                                            If IsNothing(ValStr) = False Then
                                                ArrayTableTypeObject(j, CountArrayTableTypeObject) = ValStr
                                            Else
                                                ArrayTableTypeObject(j, CountArrayTableTypeObject) = ""
                                            End If
                                        Next j
                                        CountArrayTableTypeObject += 1
                                    Next i
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        FuncListObjectDataTableDestDWG2 = CountArrayTableTypeObject
        sourceDb.CloseInput(True)
    End Function
    '=========================================================================================================================================================
    'функция копирует объекты из файла-шаблона
    Public Shared Function FuncCopyStylesDestDraw2(ByVal PutshDraw As String, ByVal ArrayNameObject As String(,), ByVal boolReplase As Boolean) As Boolean
        FuncCopyStylesDestDraw2 = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(PutshDraw, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            If IsArray(ArrayNameObject) = True Then
                If ArrayNameObject.GetUpperBound(1) > 0 Then
                    '==========================================================================================
                    'ищем столбец с типом объекта
                    Dim PosTypeObject As Integer = -1
                    For i As Integer = 1 To ArrayNameObject.GetUpperBound(0)
                        Dim NameColumn As String = ArrayNameObject(i, 0)
                        If NameColumn.Trim Like "Type" Then
                            PosTypeObject = i
                            Exit For
                        End If
                    Next
                    If PosTypeObject > -1 Then
                        For i As Integer = 1 To ArrayNameObject.GetUpperBound(1)
                            Dim TypeObject As String = ArrayNameObject(PosTypeObject, i)
                            For j As Integer = 1 To ArrayNameObject.GetUpperBound(0)
                                Try
                                    Dim ValStr As String = ArrayNameObject(j, i).Trim
                                    Dim NameColumn As String = ArrayNameObject(j, 0).Trim
                                    If NameColumn.Trim Like "Name" Then
                                        If TypeObject Like "Полилиния" Then
                                            Dim acLineTypTbl As LinetypeTable = sourceTr.GetObject(sourceDb.LinetypeTableId, OpenMode.ForRead)
                                            If acLineTypTbl.Has(ValStr) = True Then
                                                Dim IDTypeLine As ObjectId = acLineTypTbl.Item(ValStr)
                                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                                ObjColID.Add(IDTypeLine)
                                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                                Dim db1 As Database = New Database(False, False)
                                                Dim destDb1 As Database = acDoc.Database
                                                Dim destDictId1 As ObjectId = destDb1.LinetypeTableId
                                                Dim iMap1 As IdMapping = New IdMapping()
                                                If boolReplase = True Then
                                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                                Else
                                                    destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                                End If
                                            End If
                                        ElseIf TypeObject Like "Блок" Then
                                            Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                                            If acBlockTbl.Has(ValStr) = True Then
                                                Dim IDBlk As ObjectId = acBlockTbl.Item(ValStr)
                                                Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                                ObjColID.Add(IDBlk)
                                                Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                                Dim destDb1 As Database = acDoc.Database
                                                Dim destDictId1 As ObjectId = destDb1.BlockTableId
                                                Dim iMap1 As IdMapping = New IdMapping()
                                                If boolReplase = False Then
                                                    destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Ignore, False)
                                                Else
                                                    destDb1.WblockCloneObjects(ObjColID, destDb1.BlockTableId, iMap1, DuplicateRecordCloning.Replace, False)
                                                End If
                                            End If
                                        End If
                                    ElseIf NameColumn.Trim Like "Layer" Then
                                        Dim acLyrTbl As LayerTable = sourceTr.GetObject(sourceDb.LayerTableId, OpenMode.ForRead)
                                        If acLyrTbl.Has(ValStr) = True Then
                                            Dim IDLayer As ObjectId = acLyrTbl.Item(ValStr)
                                            Dim ObjColID As ObjectIdCollection = New ObjectIdCollection
                                            ObjColID.Add(IDLayer)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb1 As Database = acDoc.Database
                                            Dim destDictId1 As ObjectId = destDb1.LayerTableId
                                            Dim iMap1 As IdMapping = New IdMapping()
                                            If boolReplase = False Then
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                                            Else
                                                destDb1.WblockCloneObjects(ObjColID, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                                            End If
                                        End If
                                    ElseIf NameColumn.Trim Like "PS" Then
                                        Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(sourceDb)
                                        If dictPropSetDef.Has(ValStr, sourceTr) = True Then
                                            Dim IDPSTable As ObjectId = dictPropSetDef.GetAt(ValStr)
                                            Dim CollDictPropSet As ObjectIdCollection = New ObjectIdCollection
                                            CollDictPropSet.Add(IDPSTable)
                                            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                            Dim destDb As Database = acDoc.Database
                                            Dim destTr As Transaction = destDb.TransactionManager.StartTransaction()
                                            Using destTr
                                                Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(destDb)
                                                If dictPropSetDefTemp.Has(ValStr, destTr) = False Then
                                                    Dim helpme As Autodesk.AEC.ApplicationServices.Utility.CloningHelper = New Autodesk.AEC.ApplicationServices.Utility.CloningHelper
                                                    Dim RXClassPS As RXClass = RXClass.GetClass(GetType(Autodesk.AEC.PropertyData.DatabaseServices.PropertySet))
                                                    helpme.Clone(sourceDb, destDb, CollDictPropSet, RXClassPS, True)
                                                End If
                                                destTr.Commit()
                                            End Using
                                        End If
                                    End If
                                Catch ex As System.Exception
                                End Try
                            Next
                        Next
                    End If
                End If
            End If
            sourceTr.Commit()
        End Using
        sourceDb.CloseInput(True)
        FuncCopyStylesDestDraw2 = True
    End Function
    '=====================================================================================================================================================
    'функция возвращает значение элемента в массиве характеристик объекта
    Public Shared Function FuncReturnValueArray(ByVal ArrayNameObject As String(,), ByVal nameFieldColumn As String, ByVal NameObject As String) As String
        FuncReturnValueArray = ""
        If IsArray(ArrayNameObject) = True Then
            If ArrayNameObject.GetUpperBound(1) > 0 Then
                '==========================================================================================
                'ищем столбец с типом объекта
                Dim PosDeskObject As Integer = -1
                For i As Integer = 1 To ArrayNameObject.GetUpperBound(0)
                    Dim NameColumn As String = ArrayNameObject(i, 0)
                    If NameColumn.Trim Like "Description" Then
                        PosDeskObject = i
                        Exit For
                    End If
                Next
                '==========================================================================================
                'ищем столбец по которому идет поиск
                Dim PosnameFieldColumn As Integer = -1
                For i As Integer = 1 To ArrayNameObject.GetUpperBound(0)
                    Dim NameColumn As String = ArrayNameObject(i, 0)
                    If NameColumn.Trim Like nameFieldColumn Then
                        PosnameFieldColumn = i
                        Exit For
                    End If
                Next
                '=========================================================================================
                'ищем значение
                If PosDeskObject > -1 And PosnameFieldColumn > -1 Then
                    For i As Integer = 1 To ArrayNameObject.GetUpperBound(1)
                        Dim TypeObject As String = ArrayNameObject(PosDeskObject, i)
                        If TypeObject.Trim Like NameObject Then
                            FuncReturnValueArray = ArrayNameObject(PosnameFieldColumn, i)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кадастровый план территории (новая версия)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    <CommandMethod("ImportRosreestrKPT")>
    Public Sub RosreestrImportKPT2()
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
                Dim FormKPTImport As CadastralTools.FormImportKPT = New CadastralTools.FormImportKPT(generalDirectory)
                Dim tempPath As String = Path.GetTempPath()
                tempPath = tempPath & "rosreestr"
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

                Dim dictCadastralNumber As List(Of String) = New List(Of String) 'список кадастровый номеров для проверки дубликатов
                If FormKPTImport.boolImportXML = True Then
                    If Directory.Exists(FormKPTImport.initDirectory) = True Then
                        Try
                            tempPath = Path.Combine(FormKPTImport.initDirectory, Guid.NewGuid().ToString())
                            Directory.CreateDirectory(tempPath)
                        Catch ex As Exception
                            MsgBox("Не удалось создать папку для выходных данных.")
                            Exit Sub
                        End Try
                    Else
                        Exit Sub
                    End If
                    If FormKPTImport.boolDubleObject = True Then
                        dictCadastralNumber = FuncCadastr.getNumberCadastralObject()
                    End If
                    Try
                        If Directory.Exists(tempPath) = False Then
                            Directory.CreateDirectory(tempPath)
                        End If
                        If Directory.Exists(tempPath) = True Then
                            CadastralTools.ImportKPT.RosreestrImportKPT(tempPath, nameKvartal)
                        End If
                    Catch ex As System.Exception
                    End Try
                Else
                    tempPath = FormKPTImport.initDirectory
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim files As String() = Directory.GetFiles(tempPath)
                If files.Length = 0 Then
                    MsgBox("Не удалось корректно прочитать XML. Макро прервано.")
                    Exit Sub
                End If
                Dim cadastralNetSuit As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                'библиотека стилей отображения
                Dim dictonaryStyle As Dictionary(Of String, CadastralPolylineStyle) = New Dictionary(Of String, CadastralPolylineStyle)
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    '========================================================================================================================================
                    'копируем стили с файла-шаблона
                    Dim PutchTemplateDwg As String = templateXML.Replace(".xml", ".dwg")
                    Dim ArrayLineProperties As String(,) = Nothing
                    Dim countArrayLineProperties As Integer = FuncListLineObjectDataTableDestDWG(PutchTemplateDwg, "Земельный кадастр", ArrayLineProperties)
                    If System.IO.File.Exists(PutchTemplateDwg) = False Then
                        MsgBox("Шаблон не найден (файл формата dwg. Макро прервано!!!")
                        Exit Sub
                    End If
                    If IsArray(ArrayLineProperties) = True Then
                        If ArrayLineProperties.GetUpperBound(1) > 0 Then
                            Dim boolCopyStyle As Boolean = FuncCopyLineStylesDestDraw(PutchTemplateDwg, ArrayLineProperties, False)
                        End If
                    End If
                    'загружаем в чертеж тавлицу Внутренний островок
                    Dim NamePSSetTableOstrov As String = "CADASTRAL_OSTROV"
                    Dim boolFindPSOstrov As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NamePSSetTableOstrov)
                    Dim IDPSTableOstrov As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_OSTROV")
                    '===============================================================================================================================================================
                    'выбираем внешнюю границу для отсечки лишних объектов
                    Dim classTopo As FuncNetTopologySuite = New FuncNetTopologySuite
                    Dim boundPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                    If FormKPTImport.boolSelectBound = True Then
                        Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите внешнюю границу:")
                        If ptResult.Status = PromptStatus.OK Then
                            Dim ent As Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                            If TypeOf ent Is Polyline Then
                                Dim boundPline As Polyline = TryCast(ent, Polyline)
                                If boundPline.Area > 0 Then
                                    boundPolygon = classTopo.CreatePolygon(boundPline)
                                End If
                            End If
                        End If
                    End If
                    'читаем земельные участки
                    For Each file As String In files
                        Try
                            Dim listDictionaryData As List(Of Dictionary(Of String, Object)) = CadastralTools.JsonValidator.CreateObjectsFromJsonFile(file)
                            If listDictionaryData.Count > 0 Then
                                Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                Dim dictLandProperties As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                dataCadastralObject.Label = "Земельный кадастр"
                                Dim idStructure As String = ""
                                Dim tableZn As String = ""
                                Dim PSTable As String = ""
                                'Dim typeZn As String = "Полилиния"
                                Dim Label As String = ""
                                Dim idName As Integer = 0
                                Dim countOstrov As Integer = 1
                                Dim listPolygonCoordinates As List(Of String(,)) = New List(Of String(,))
                                For Each cadastralObject As Dictionary(Of String, Object) In listDictionaryData
                                    If IsNothing(cadastralObject) = True Then Continue For
                                    If cadastralObject.Count > 0 Then
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

                                        If cadastralObject.ContainsKey("ListCoord") Then
                                            Dim dataCoordParcel As Dictionary(Of String, Object) = cadastralObject("ListCoord")

                                            For Each kvp As KeyValuePair(Of String, Object) In dataCoordParcel
                                                Dim dataCoordCounter As List(Of Object) = TryCast(kvp.Value, List(Of Object))
                                                If dataCoordCounter Is Nothing OrElse dataCoordCounter.Count = 0 Then Continue For

                                                For Each coordItem As Object In dataCoordCounter
                                                    Dim listCoordinates As List(Of Object) = TryCast(coordItem, List(Of Object))
                                                    If listCoordinates Is Nothing OrElse listCoordinates.Count < 3 Then Continue For

                                                    ' Извлекаем списки координат (индексы 1 и 2)
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

                                                    If IsNothing(boundPolygon) = False Then
                                                        If boundPolygon.Area > 1 Then
                                                            If Not (PSTable Like "LAND_CADASTRAL_DIVISION") Then
                                                                If arrayCoord.GetUpperBound(1) = 0 Then
                                                                    'это точка
                                                                    Dim point As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(arrayCoord(0, 0), arrayCoord(1, 0))
                                                                    Dim boolIntersect As Boolean = cadastralNetSuit.IsPointInsidePolygon(point, boundPolygon)
                                                                    If boolIntersect = False Then
                                                                        Continue For
                                                                    End If
                                                                Else
                                                                    Dim lineStr As NetTopologySuite.Geometries.LineString = cadastralNetSuit.CreateLineString(arrayCoord)
                                                                    Dim boolIntersect As Boolean = cadastralNetSuit.DoesLineIntersectPolygon(lineStr, boundPolygon)
                                                                    If boolIntersect = False Then
                                                                        Continue For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    listPolygonCoordinates.Add(arrayCoord)
                                                Next coordItem
                                            Next kvp
                                        End If
                                    End If
                                Next
                                'преобразуйм в единую строку
                                Dim keyParam As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictLandProperties)
                                '==========================================================================================================
                                'включена опция проверки дубликатов
                                Dim cadNumber As String = ""
                                Dim cadObjectNumber As Object = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Cad_number")
                                If TypeOf cadObjectNumber Is String Then
                                    cadNumber = cadObjectNumber
                                Else
                                    cadObjectNumber = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Reg_numb_border")
                                    If TypeOf cadObjectNumber Is String Then
                                        cadNumber = cadObjectNumber
                                    End If
                                End If
                                If dictCadastralNumber.Count > 0 Then
                                    If dictCadastralNumber.Contains(cadNumber) = True Then
                                        Continue For
                                    End If
                                End If
                                'ФОРМИРУЕМ СЕМАНТИКУ
                                dataCadastralObject.NameTypeObject = CadastralTools.StructureElement.GetEnumByValue(GetType(nameObject), idName)
                                dataCadastralObject.Description = CadastralTools.StructureElement.GetEnumDescriptionByValue(GetType(nameObject), idName)
                                dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                                dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                dataCadastralObject.KeyParameter = keyParam
                                dataCadastralObject.IdElement = Guid.NewGuid.ToString
                                dataCadastralObject.TablePS = PSTable
                                Select Case PSTable
                                    Case "LAND_CADASTRAL_DIVISION"
                                        tableZn = "ЕГРН-Кадастровое деление"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                    Case "LAND_CADASTRAL_PARCEL"
                                        tableZn = "ЕГРН-Кадастровые участки"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                    Case "LAND_CADASTRAL_PART"
                                        tableZn = "ЕГРН-Кадастровые участки"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                    Case "LAND_CADASTRAL_BUILDING"
                                        tableZn = "ЕГРН-Здания (строения)"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                    Case "LAND_CADASTRAL_CONSTRUCTION"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                                        tableZn = "ЕГРН-Сооружения"
                                    Case "LAND_CADASTRAL_UNFINISHED_CONSTRUCTION"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                                        tableZn = "ЕГРН-Объекты незавершенного строительства"
                                    Case "LAND_CADASTRAL_ZONE"
                                        tableZn = "ЕГРН-Территориальные зоны"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                    Case "LAND_CADASTRAL_BOUNDARY"
                                        tableZn = "ЕГРН-Границы территорий"
                                        If dataCadastralObject.NameTypeObject = nameObject.WaterBoundary Then
                                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                                        Else
                                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                        End If
                                    Case "LAND_CADASTRAL_REDLINE"
                                        tableZn = "ЕГРН-Границы территорий"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                                    Case "LAND_CADASTRAL_COASTLINE"
                                        tableZn = "ЕГРН-Границы территорий"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polyline
                                End Select
                                '======================================================================================================================
                                Dim IDPSTable As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(dataCadastralObject.TablePS)
                                Dim createPs As Boolean = False
                                If IsNothing(IDPSTable) = True Then
                                    createPs = True
                                ElseIf IDPSTable.IsNull = True Then
                                    createPs = True
                                End If
                                If createPs = True Then
                                    Dim arrayPropertiesPS As String(,) = {}
                                    Dim boolFindPropertiesPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, dataCadastralObject.TablePS, arrayPropertiesPS)
                                    If arrayPropertiesPS.Length = 0 Then
                                        Dim boolFindPS As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, dataCadastralObject.TablePS)
                                        IDPSTable = FuncPropertySet.FuncReturnIDDictSetProperties(dataCadastralObject.TablePS)
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
                                    Dim boolFindPropertyObject As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZn, tableZn, dataCadastralObject.TypeGeometry, dataCadastralObject.Description, arrayProperties)
                                    If IsArray(arrayProperties) = True Then
                                        newStyle.createLineStyle(arrayProperties)
                                    End If
                                    dictonaryStyle.Add(dataCadastralObject.Description, newStyle)
                                Else
                                    newStyle = dictonaryStyle.Item(dataCadastralObject.Description)
                                End If
                                '=============================================================================================================================
                                'рисуем контур
                                If listPolygonCoordinates.Count > 0 Then
                                    Dim listPolygon As Dictionary(Of Object, NetTopologySuite.Geometries.Polygon) = New Dictionary(Of Object, NetTopologySuite.Geometries.Polygon)
                                    For i As Integer = 0 To listPolygonCoordinates.Count - 1
                                        Dim arrayCoord As String(,) = listPolygonCoordinates.ElementAt(i)
                                        If arrayCoord.Length > 2 Then
                                            Dim acPolyline As Polyline = AcadFunction.createPolylineFromArrayCoordinates2(arrayCoord)
                                            If IsNothing(acPolyline) = False Then
                                                btr.AppendEntity(acPolyline)
                                                tr.AddNewlyCreatedDBObject(acPolyline, True)
                                                newStyle.setStyle(acPolyline)
                                                If dataCadastralObject.TypeGeometry = typeGeometryObject.polygon Then
                                                    acPolyline.Closed = True
                                                    Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = classTopo.CreatePolygon(acPolyline)
                                                    listPolygon.Add(acPolyline, sourcePolygon)
                                                    If listPolygon.Count > 1 Then
                                                        Dim boolFindOstrov As Boolean = False
                                                        For j As Integer = 0 To listPolygon.Count - 2
                                                            Dim destPolygon As NetTopologySuite.Geometries.Polygon = listPolygon.ElementAt(j).Value
                                                            If destPolygon.Area > sourcePolygon.Area Then
                                                                If cadastralNetSuit.IsPolygonInside(destPolygon, sourcePolygon) = True Then
                                                                    boolFindOstrov = True
                                                                    Exit For
                                                                ElseIf cadastralNetSuit.IsPolygonInside(sourcePolygon, destPolygon) = True Then
                                                                    boolFindOstrov = True
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next j
                                                        If boolFindOstrov = True Then
                                                            Dim objOstrov As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                                            objOstrov.Label = dataCadastralObject.Label
                                                            objOstrov.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                                                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                                            objOstrov.NameTypeObject = dataCadastralObject.NameTypeObject
                                                            objOstrov.Description = "Внутренний островок"
                                                            objOstrov.IdStructure = dataCadastralObject.IdElement
                                                            objOstrov.IdElement = Guid.NewGuid.ToString
                                                            objOstrov.TablePS = "CADASTRAL_OSTROV"
                                                            Dim dictOstrov As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                            dictOstrov.Add("ConditionalNumber", countOstrov)
                                                            dictOstrov.Add("ParentNumber", cadNumber)
                                                            Dim keyParamOstrov As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictOstrov)
                                                            objOstrov.KeyParameter = keyParamOstrov
                                                            'записываем xdata
                                                            Dim boolWriteXRecordsOstrov As Boolean = FuncXRecords.setXRecords(acPolyline, objOstrov)
                                                            'записываем PS
                                                            Dim boolPSObjectOstrov As Boolean = FuncPropertySet.FuncAddDataSetProperties(acPolyline, IDPSTableOstrov)
                                                            If boolPSObjectOstrov = True Then
                                                                Dim boolRecOstrov As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(acPolyline, objOstrov)
                                                            End If
                                                            countOstrov += 1
                                                            Continue For
                                                        End If
                                                    End If
                                                End If
                                                'записываем xdata
                                                Dim boolWriteXRecords As Boolean = FuncXRecords.setXRecords(acPolyline, dataCadastralObject)
                                                'записываем PS

                                                Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acPolyline, IDPSTable)
                                                If boolPSObject = True Then
                                                    Dim boolRec As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(acPolyline, dataCadastralObject)
                                                End If
                                            End If
                                        Else
                                            Dim acCircle As Circle = New Circle()
                                            acCircle.SetDatabaseDefaults()
                                            acCircle.Center = New Point3d(arrayCoord(0, 0), arrayCoord(0, 1), 0)
                                            acCircle.Radius = 0.5
                                            btr.AppendEntity(acCircle)
                                            tr.AddNewlyCreatedDBObject(acCircle, True)
                                            newStyle.setStyle(acCircle)
                                            dataCadastralObject.TypeGeometry = typeGeometryObject.block
                                            'записываем xdata
                                            Dim boolWriteXRecordsOstrov As Boolean = FuncXRecords.setXRecords(acCircle, dataCadastralObject)
                                            'записываем PS
                                            Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acCircle, IDPSTable)
                                            If boolPSObject = True Then
                                                Dim boolRec As Boolean = FuncPropertySet.FuncWriteDataSetByStructureElement(acCircle, dataCadastralObject)
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        End Try
                    Next
                    tr.Commit()
                End Using
                If FormKPTImport.boolImportXML = True Then
                    If nameKvartal.Trim.Length > 0 Then
                        Dim boolKVName As Boolean = CadastralTools.SupportFiles.RenameDirectory(tempPath, nameKvartal)
                    End If
                End If
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
            End Try
        End Using
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '=============================================================================================
    'отрисовка и нумерация полигона на основе кадастрового плана территории (xml-файл) (Росреестр)
    <CommandMethod("RosreestrImportKPT2")>
    Public Sub RosreestrImportKPT()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim LineErrNumber As Integer = 0
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
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'читаем масштаб аннотаций (текущий)
                Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
                'передаем масштаб чертежа в форму
                Dim FormKPTImport As FormRosreestrKPTImport = New FormRosreestrKPTImport
                FormKPTImport.NumericUpDown1.Value = MAnno
                'запускаем форму
                FormKPTImport.ShowDialog()
                If FormKPTImport.boolShow = False Then
                    Exit Sub
                End If

                'начальный путь к файлу-шаблону
                Dim PutchTemplateXml As String = FormKPTImport.ComboBox1.Tag & FormKPTImport.ComboBox1.Text & ".xml"
                Dim PutchTemplateDwg As String = FormKPTImport.ComboBox1.Tag & FormKPTImport.ComboBox1.Text & ".dwg"

                'Dim boolFindTeplates As Boolean = FuncFiles.FuncFindTemplateFiles(nameFileXML, nameFileDwg, PutchTemplateXml, PutchTemplateDwg)
                '=============================================================================================================
                'проверяем и загружаем все стили из шаблона
                '=============================================================================================================
                '3. Путь к файлам шаблонам с условными знами
                'If boolFindTeplates = False Then
                '    MsgBox("Шаблон не найден. Добавьте файлы Стандарт_ADSK.xml и Стандарт_ADSK.dwg в вспомогательные файлы поддержки. Макро прервано!!!")
                '    Exit Sub
                'End If
                Dim CountArrayLineProperties As Integer = 0
                Dim ArrayLineProperties As String(,) = Nothing
                CountArrayLineProperties = FuncListLineObjectDataTableDestDWG(PutchTemplateDwg, "Земельный кадастр", ArrayLineProperties)

                'If FormKPTImport.CheckBox6.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Кадастровое деление", "Полилиния", ArrayLineProperties)
                'End If
                'If FormKPTImport.CheckBox1.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Кадастровые участки", "Полилиния", ArrayLineProperties)
                'End If
                'If FormKPTImport.CheckBox2.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Здания (строения)", "Полилиния", ArrayLineProperties)
                'End If
                'If FormKPTImport.CheckBox3.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Сооружения", "Полилиния", ArrayLineProperties)
                'End If
                'If FormKPTImport.CheckBox4.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Границы территорий", "Полилиния", ArrayLineProperties)
                'End If
                'If FormKPTImport.CheckBox5.Checked = True Then
                '    CountArrayLineProperties = TOPOSymbols.FuncReadDataTableToXml(PutchTemplateXml, "Земельный кадастр", "ЕГРН-Территориальные зоны", "Полилиния", ArrayLineProperties)
                'End If
                '=============================================================================================================================================================
                'загружаем все объекты в существующую базу
                If System.IO.File.Exists(PutchTemplateDwg) = False Then
                    MsgBox("Шаблон не найден (файл формата dwg. Макро прервано!!!")
                    Exit Sub
                End If
                If IsArray(ArrayLineProperties) = True Then
                    If ArrayLineProperties.GetUpperBound(1) > 0 Then
                        Dim boolCopyStyle As Boolean = FuncCopyLineStylesDestDraw(PutchTemplateDwg, ArrayLineProperties, False)
                    End If
                End If
                '===============================================================================================================================================================
                'загружаем в чертеж тавлицу Внутренний островок
                Dim NamePSSetTableOstrov As String = "CADASTRAL_OSTROV"
                Dim IDPSSetTableOstrov As ObjectId = Nothing
                '=======================================================================================
                'задаем коэффициент масштаба
                Dim M As Double = FormKPTImport.NumericUpDown1.Value / 1000
                '================================================================================
                'задаем высоту текстов
                Dim HtextKV As Double = FormKPTImport.NumericUpDown6.Value
                Dim HtextParcel As Double = FormKPTImport.NumericUpDown2.Value
                Dim HtextConstruction As Double = FormKPTImport.NumericUpDown4.Value
                Dim HtextBuild As Double = FormKPTImport.NumericUpDown3.Value
                Dim HtextZone As Double = FormKPTImport.NumericUpDown5.Value
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '============================================================================================
                'начальные данные для считывания файла xml
                '============================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ObjectCollections As String(,) = Nothing
                Dim CountObjectCollection As Integer = 0
                Dim boolKvartal As Boolean = False 'обрабатывается квартал
                Dim boolspatial_data As Boolean = False
                Dim boolBuild As Boolean = False 'обрабатывается здание
                Dim boolParcel As Boolean = False 'обрабатывается участок
                Dim boolPart As Boolean = False 'обрабатываеся часть участка, здания, сооружения
                Dim boolCounter As Boolean = False 'обрабатывается контур участка
                Dim boolconstruction As Boolean = False 'обрабатывается сооружение
                Dim boolunderconstruction As Boolean = False 'обрабатывается объект незавершенного строительства
                Dim boolboundary As Boolean = False 'обрабатывается граница
                Dim boolcityboundary As Boolean = False 'обрабатывается граница города
                Dim boolzone As Boolean = False 'обрабатывается зона
                Dim boolcoastline As Boolean = False 'береговая полоса
                Dim boolNastr As Boolean = False 'флаг Говорит о том что нужно ли проверять настройки (слой, цвет и т.п.)
                '============================================================================================
                Dim boolNumberSubParcel As Boolean = False
                Dim boolSpatialElements As Boolean = False
                Dim boolCode As Boolean = False
                Dim boolType As Boolean = False
                Dim boolval As Boolean = False
                Dim boolcomdata As Boolean = False
                Dim boolunit As Boolean = False
                'начальное значение номера точки
                Dim Numb As Integer = 1 'номер
                Dim NXYZ As String(,) = Nothing 'массив для хранения координат объектов
                Dim boolX As Boolean = False
                Dim X As String = ""
                Dim boolY As Boolean = False
                Dim Y As String = ""
                Dim boolOrgNumb As Boolean = False
                Dim OrgNumb As String = "" 'номер
                Dim boolnum_geopoint As Boolean = False
                Dim num_geopoint As String = ""
                Dim boolGeopoint_opred As Boolean = False
                Dim Geopoint_opred As String = ""
                Dim boolDelta As Boolean = False
                Dim Delta As String = "0" 'погрешность
                Dim boolgeopoint_zacrep As Boolean = False
                Dim geopoint_zacrep As String = "-" 'закрепление
                Dim boolR As Boolean = False
                Dim r As String = "" 'радиус окружности
                Dim boolGeopointValue As Boolean = False
                Dim GeopointValue As String = ""
                Dim ArrayCounterParcel As String(,) = Nothing
                Dim CountArrayCounterParcel As Integer = 0
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тип объекта
                '0-кадастровый квартал
                Dim boolarea_quarter As Boolean = False
                Dim Area_quarter As String = ""
                Dim CadastralBlock As String = "-"
                Dim boolKvartalNumber As Boolean = False
                '1. Общие данные для участков, ОКС и сооружений
                Dim boolCadNumber As Boolean = False 'кадастровый номер объекта
                Dim Cad_Number As String = "" 'полный кадастровый номер объекта
                Dim Cad_Number_Note As String = "" 'для единого землепользования
                Dim Cad_Number_Short As String = "" 'краткий кадастровый номер
                Dim CodTypeObject As String = "" 'код объекта (Земельный участок, сооружение, и т.п.)
                Dim ValTypeObject As String = "" 'значение объекта
                Dim boolDateReg As Boolean = False
                Dim dateReg As String = ""
                'общие для участков и зданий (участки-разрешенное использование, знания - тип здания (жилое, нежилое)
                Dim boolpermitted_use As Boolean = False 'разрешенное использование
                '1. Земельные участки
                Dim StateSubParcel As String = "" 'тип подучастка 01-ранее учтенный 06-учтенный ....
                Dim NumParcel As String = "" 'краткий кадастровый номер
                Dim boolNumberContour As Boolean = False
                Dim NumberContour As String = "" 'номер контура
                Dim NumberSubParcel As String = "" 'номер части
                Dim areaPart As String = "" 'площадь части участка
                Dim boolcommon_land_cad_number As Boolean = False
                Dim common_land_cad_number As String = ""
                'тип участка
                Dim boolSubTypeParcel As Boolean = False
                Dim SubTypeParcel As String = "" 'тип подучастка землепользование, многоконтурный участок и т.п
                'вид участка
                Dim StateParcel As String = "06" '06-учтенный
                Dim boolStateParcel As Boolean = False
                'категория земель
                Dim boolCategory As Boolean = False
                Dim Category As String = ""
                Dim CodCategory As String = ""
                'разрешенное использование 
                Dim boolRazrIsp As Boolean = False
                Dim byRazrIspByDoc As Boolean = False
                Dim RazrIsp As String = ""
                Dim RazrIspByDoc As String = ""
                'адрес
                '===============================================================
                Dim boolAdresParcel As Boolean = False
                Dim AddresParcel As String = ""
                Dim boolokato As Boolean = False
                Dim okato As String = ""
                Dim boolkladr As Boolean = False
                Dim kladr As String = ""
                Dim boolpostcode As Boolean = False
                Dim postcode As String = ""
                Dim boolregion As Boolean = False
                Dim region As String = ""
                Dim boolDistrict As Boolean = False
                Dim booltypeDistrict As Boolean = False
                Dim typeDistrict As String = ""
                Dim boolnameDistrict As Boolean = False
                Dim nameDistrict As String = ""
                Dim boolLocality As Boolean = False
                Dim booltypeLocality As Boolean = False
                Dim typeLocality As String = ""
                Dim boolnameLocality As Boolean = False
                Dim nameLocalyty As String = ""
                Dim boolStreet As Boolean = False
                Dim booltypeStreet As Boolean = False
                Dim type_street As String = ""
                Dim boolnameStreet As Boolean = False
                Dim name_street As String = ""
                Dim boolLevel1 As Boolean = False
                Dim booltypeLevel1 As Boolean = False
                Dim typeLevel1 As String = ""
                Dim boolnameLevel1 As Boolean = False
                Dim nameLevel1 As String = ""
                Dim boolLevel2 As Boolean = False
                Dim booltypeLevel2 As Boolean = False
                Dim typeLevel2 As String = ""
                Dim boolnameLevel2 As Boolean = False
                Dim nameLevel2 As String = ""
                Dim boolLevel3 As Boolean = False
                Dim booltypeLevel3 As Boolean = False
                Dim typeLevel3 As String = ""
                Dim boolnameLevel3 As Boolean = False
                Dim nameLevel3 As String = ""
                Dim readable_address As String = ""
                Dim boolreadable_address As Boolean = False
                '====================================================================
                'здания
                Dim boolpurpose As Boolean = False 'тип здания жилое нежилое
                Dim purpose_build As String = ""
                Dim boolposition_description As Boolean = False 'описание местоположения
                Dim pos_description_build As String = ""
                Dim codeBuild As String = "" 'назначение код здания
                Dim permitted_use As String = ""
                Dim boolname As Boolean = False
                Dim Assignatuin As String = ""
                Dim MaterialWall As String = ""
                Dim CountFloat As String = ""
                '============================================================================
                'сооружения
                Dim purpose_construction As String = ""
                Dim boolparams As Boolean = False 'параметры сооружения
                Dim boolbase_parameter As Boolean = False
                Dim booldepth As Boolean = False 'глубина залегания
                Dim baseParameters As String = ""
                Dim depth As String = ""
                Dim boolextension As Boolean = False 'длина
                Dim extension As String = ""
                Dim boolvolume As Boolean = False
                Dim volume As String = ""
                Dim boolbuilt_up_area As Boolean = False
                Dim boolareaConstr As Boolean = False
                '============================================================================
                'границы и зоны
                Dim booltype_boundary As Boolean = False
                Dim type_boundary As String = ""
                Dim boolregistration_date As Boolean = False
                Dim registr_data As String = ""
                Dim boolRegNumbZone As Boolean = False 'номер территориальной зоны
                Dim RegNumbZone As String = "" 'номер зоны
                Dim boolType_zone As Boolean = False
                Dim cod_boundary As String = ""
                Dim type_zona As String = ""
                Dim code_zona As String = ""
                Dim boolNumberZona As Boolean = False
                Dim boolIndexZona As Boolean = False
                Dim NumberZona As String = ""
                Dim IndexZona As String = ""
                '====================================================================================
                'береговые линии
                Dim boolwater_object_type As Boolean = False
                Dim boolwater_object_name As Boolean = False
                Dim water_object_name As String = ""
                Dim codwater_object_type As String = ""
                Dim typewater_object_type As String = ""
                '====================================================================================
                'собственник
                Dim boolRightType As Boolean = False
                Dim RightType As String = ""
                Dim boolOwner As Boolean = False
                Dim boolOwnerName As Boolean = False
                Dim ownerName As String = ""
                Dim boolresident As Boolean = False
                Dim encumbranceName As String = ""
                Dim boolrecord_info As Boolean = False
                Dim boolRightName As Boolean = False
                Dim boolRight_holder As Boolean = False
                '===============================================================================
                'реквизиты выписки
                Dim booldate_formation As Boolean = False
                Dim dataVP As String = "-"
                Dim boolregistration_number As Boolean = False
                Dim NumberVP As String = ""
                Dim boolgroup_top_requisites As Boolean = False
                '========================================================
                'площадь участка общая
                Dim boolGlobalArea As Boolean = False
                Dim GlobalArea As String = ""
                Dim deltaArea As String = ""
                Dim booldeltaArea As Boolean = False
                'кадастровая стоимость
                Dim boolCostParcel As Boolean = False
                Dim CostParcel As String = ""
                'Счетчик вершин контура (для массива координат)
                Dim countcoord As Integer = -1
                Dim TypeLineSubParcel As String = ""
                Dim LocalID As Integer = 0
                Dim GlobalID As Integer = 0
                Dim rng As Random = New Random()
                '=======================================================================================================
                'переменные для оформления контуров
                Dim LayerObject As String = "0"
                Dim LineTypeObj As String = "Continuous"
                Dim IndexColorObj As Integer = 256
                Dim WidthrObj As Integer = 20
                Dim ScLine As Double = 1
                Dim GlobalContLine As Double = 0
                Dim NamePSSetTable As String = "OBJECT"
                Dim IDPSTAble As ObjectId = Nothing
                Dim OldCodObject As String = ""
                '========================================================================================================
                'массив с кадастровыми номерами всех кадастровых объектов
                Dim ArrayOldCadastrNumbers As String() = Nothing
                Dim CountArrayOldCadastrNumbers As Integer = 0
                Dim ArrayOldCadastrRegNumbers As String() = Nothing
                Dim CountArrayOldRegNumbers As Integer = 0
                Dim ArrayOldCadastrKVNumbers As String() = Nothing
                Dim CountArrayOldKVNumbers As Integer = 0
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тело программы
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбор файла
                Dim NameFiles As String() = {""}
                Dim fileDialog1 As OpenFileDialog = New OpenFileDialog("Открыть файл xml", "", "xml", "xml", OpenFileDialogFlags.AllowMultiple)
                Dim FileRez As String = fileDialog1.ShowDialog()
                NameFiles = fileDialog1.GetFilenames
                If IsNothing(NameFiles) = True Then Exit Sub
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                    Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
                    Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
                    Dim StText As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                    'загружаем стиль текста
                    Dim idTTextStyle As ObjectId = AcadFunction.FuncReturnIDTextStylebyName("CADASTRAL_OBJECT")
                    If IsNothing(idTTextStyle) = True Or idTTextStyle.IsNull = True Then
                        Dim boolCopyTextStyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, "CADASTRAL_OBJECT")
                        If boolCopyTextStyle = True Then
                            idTTextStyle = AcadFunction.FuncReturnIDTextStylebyName("CADASTRAL_OBJECT")
                        Else
                            idTTextStyle = acCurDb.Textstyle
                        End If
                    End If
                    'загружаем таблицу PS с островками
                    If dictPropSetDefTemp.Has("CADASTRAL_OSTROV", acTrans) = False Then
                        Dim AddTBLCadOstrod As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, "CADASTRAL_OSTROV")
                        IDPSSetTableOstrov = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_OSTROV")
                        If IsNothing(IDPSSetTableOstrov) = True Then
                            MsgBox("Не удалось из шаблона скопировать набор данных CADASTR_OSTROV.")
                        End If
                        If IDPSSetTableOstrov.IsNull = True Then
                            MsgBox("Не удалось из шаблона скопировать набор данных CADASTR_OSTROV.")
                        End If
                    Else
                        IDPSSetTableOstrov = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_OSTROV")
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'читаем файл xml
                    For i As Integer = 0 To NameFiles.GetUpperBound(0)
                        Try
                            Dim FileName As String = NameFiles(i)
                            Dim reader As XmlTextReader = New XmlTextReader(FileName)
                            Dim countFile As Integer = 0
                            While reader.Read()
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                LineErrNumber += 1
                                If LineErrNumber = 45253 Then
                                    MsgBox("")
                                End If
                                Select Case reader.NodeType
                                    Case XmlNodeType.Element
                                        'блок общих данных
                                        '=====================================================================
                                        If NameBlock Like "common_data" Then
                                            boolcomdata = True
                                        ElseIf NameBlock Like "type" Then
                                            boolType = True
                                        ElseIf NameBlock Like "code" Then
                                            boolCode = True
                                        ElseIf NameBlock Like "value" Then
                                            boolval = True
                                        ElseIf NameBlock Like "unit" Then
                                            boolunit = True
                                        ElseIf NameBlock Like "name" Then
                                            boolname = True
                                        ElseIf NameBlock Like "registration_date" Then
                                            boolDateReg = True
                                        ElseIf NameBlock Like "State" Then
                                            boolStateParcel = True
                                            '============================================================================
                                            'кадастровый квартал
                                        ElseIf NameBlock Like "area_quarter" Then
                                            boolarea_quarter = True 'площадь квартала
                                        ElseIf NameBlock Like "cadastral_number" Then
                                            boolKvartalNumber = True 'кадастровый номер квартала (КПТ)
                                        ElseIf NameBlock Like "quarter_cad_number" Then
                                            boolKvartalNumber = True 'кадастровый номер квартала (выписка)
                                            '============================================================================
                                            'часть участка и контур участка
                                        ElseIf NameBlock Like "object_parts" Then
                                            boolNastr = False 'сбрасываем настройки
                                        ElseIf NameBlock Like "object_part" Then
                                            boolPart = True 'часть кадастрового участка
                                            NumberSubParcel = ""
                                        ElseIf NameBlock Like "contours" Then
                                            'CountCounter = 0
                                        ElseIf NameBlock Like "contour" Then
                                            'контур объекта
                                            boolCounter = True
                                            NumberContour = ""
                                            '=========================================================
                                            'адрес
                                        ElseIf NameBlock Like "address" Then
                                            boolAdresParcel = True
                                        ElseIf NameBlock Like "okato" Then
                                            boolokato = True
                                        ElseIf NameBlock Like "kladr" Then
                                            boolkladr = True
                                        ElseIf NameBlock Like "postal_code" Then
                                            boolpostcode = True
                                        ElseIf NameBlock Like "region" Then
                                            boolregion = True
                                        ElseIf NameBlock Like "type_district" Then
                                            booltypeDistrict = True
                                        ElseIf NameBlock Like "name_district" Then
                                            boolnameDistrict = True
                                        ElseIf NameBlock Like "type_locality" Then
                                            booltypeLocality = True
                                        ElseIf NameBlock Like "name_locality" Then
                                            boolnameLocality = True
                                        ElseIf NameBlock Like "type_street" Then
                                            booltypeStreet = True
                                        ElseIf NameBlock Like "name_street" Then
                                            boolnameStreet = True
                                        ElseIf NameBlock Like "type_level1" Then
                                            booltypeLevel1 = True
                                        ElseIf NameBlock Like "name_level1" Then
                                            boolnameLevel1 = True
                                        ElseIf NameBlock Like "type_level2" Then
                                            booltypeLevel2 = True
                                        ElseIf NameBlock Like "name_level2" Then
                                            boolnameLevel2 = True
                                        ElseIf NameBlock Like "type_level3" Then
                                            booltypeLevel3 = True
                                        ElseIf NameBlock Like "name_level3" Then
                                            boolnameLevel3 = True
                                        ElseIf NameBlock Like "readable_address" Then
                                            boolreadable_address = True 'неструктурированный адрес
                                        ElseIf NameBlock Like "position_description" Then
                                            boolposition_description = True 'описание местоположения
                                            '===============================================================================
                                            'кадастровые номера и обозначения зон
                                        ElseIf NameBlock Like "cad_number" Then
                                            boolCadNumber = True 'кадастровый номер участка здания стоения ОНС
                                        ElseIf NameBlock Like "part_number" Then
                                            boolNumberSubParcel = True 'номер части участка
                                        ElseIf NameBlock Like "number_pp" Then
                                            boolNumberContour = True 'номер контура
                                        ElseIf NameBlock Like "reg_numb_border" Then
                                            boolRegNumbZone = True 'обозначение границы и ЗОУИТ
                                        ElseIf NameBlock Like "common_land_cad_number" Then
                                            boolcommon_land_cad_number = True
                                            '=================================================================================
                                            'площади
                                        ElseIf NameBlock Like "area" And boolbase_parameter = False Then
                                            boolGlobalArea = True 'площадь объекта
                                        ElseIf NameBlock Like "inaccuracy" Then
                                            booldeltaArea = True 'погрешность определения площади
                                            '====================================================================================
                                            'кадастровая стоимость
                                        ElseIf NameBlock Like "cost" Then
                                            boolCostParcel = True
                                            '====================================================================================
                                            'тип участка
                                        ElseIf NameBlock Like "subtype" Then
                                            boolSubTypeParcel = True 'землепользование, обособленный участок и т.п.
                                        ElseIf NameBlock Like "category" Then
                                            boolCategory = True 'категория земель
                                        ElseIf NameBlock Like "permitted_use" Then
                                            boolpermitted_use = True
                                        ElseIf NameBlock Like "by_document" Then
                                            byRazrIspByDoc = True 'разрешенное использование по документу
                                            '=======================================
                                            'здания, сооружения (характеристики)
                                        ElseIf NameBlock Like "purpose" Then
                                            boolpurpose = True 'жилое нежилое многоквартирный дом
                                        ElseIf NameBlock Like "permitted_use" Then
                                            boolpermitted_use = True 'разрешенное использование, для здания - жилое, нежилое
                                            '==============================================
                                            'характеристики строения
                                        ElseIf NameBlock Like "params" Then
                                            boolparams = True 'блок параметров
                                        ElseIf NameBlock Like "base_parameter" Then
                                            boolbase_parameter = True 'блок с базовыми параметрами
                                        ElseIf NameBlock Like "occurence_depth" Then
                                            booldepth = True 'глубина
                                        ElseIf NameBlock Like "depth" Then
                                            booldepth = True 'тоже глубина
                                        ElseIf NameBlock Like "extension" Then
                                            boolextension = True 'протяженность
                                        ElseIf NameBlock Like "volume" Then
                                            boolvolume = True 'объем
                                        ElseIf NameBlock Like "area" And boolbase_parameter = True Then
                                            boolareaConstr = True 'площадь
                                        ElseIf NameBlock Like "built_up_area" And boolbase_parameter = True Then
                                            boolbuilt_up_area = True 'площадь
                                            '=======================================================
                                            'территориальные зоны границы
                                        ElseIf NameBlock Like "record_info" Then
                                            boolrecord_info = True
                                        ElseIf NameBlock Like "type_boundary" Then
                                            booltype_boundary = True 'тип границы
                                        ElseIf NameBlock Like "type_zone" Then
                                            boolType_zone = True 'описание зоны
                                        ElseIf NameBlock Like "registration_date" Then
                                            boolregistration_date = True 'дата регистрации зоны
                                        ElseIf NameBlock Like "reg_numb_border" Then
                                            boolRegNumbZone = True 'номер зоны
                                        ElseIf NameBlock Like "number" And boolzone = True Then
                                            boolNumberZona = True 'номер зоны
                                        ElseIf NameBlock Like "index" And boolzone = True Then
                                            boolIndexZona = True 'номер зоны
                                            '=======================================================
                                            'береговые полосы
                                        ElseIf NameBlock Like "water_object_type" Then
                                            boolwater_object_type = True
                                        ElseIf NameBlock Like "water_object_name" Then
                                            boolwater_object_name = True
                                            '=====================================================================
                                            'сведения о правах
                                        ElseIf NameBlock Like "right_type" Then
                                            boolRightType = True 'тип собственности
                                        ElseIf NameBlock Like "right_holder" Then
                                            boolRight_holder = True 'сведения о правах
                                        ElseIf NameBlock Like "name" And boolRight_holder = True Then
                                            boolOwnerName = True 'тип собственности
                                        ElseIf NameBlock Like "resident" Then
                                            boolresident = True 'арендатор
                                            '=====================================================================
                                            'реквизиты выписки
                                        ElseIf NameBlock Like "group_top_requisites" Then
                                            boolgroup_top_requisites = True 'блок с реквизитами
                                        ElseIf NameBlock Like "date_formation" Then
                                            booldate_formation = True
                                        ElseIf NameBlock Like "registration_number" Then
                                            boolregistration_number = True
                                            '=============================================================
                                            'координаты
                                        ElseIf NameBlock Like "ordinate" Then
                                            countcoord += 1
                                            ReDim Preserve NXYZ(8, countcoord) 'тип,номер, x,y,delta,desk,r
                                            NXYZ(0, countcoord) = "Точка"
                                        ElseIf NameBlock Like "x" Then
                                            boolX = True
                                        ElseIf NameBlock Like "y" Then
                                            boolY = True
                                        ElseIf NameBlock Like "geopoint_opred" Then
                                            boolGeopoint_opred = True
                                        ElseIf NameBlock Like "ord_nmb" Then
                                            boolOrgNumb = True
                                        ElseIf NameBlock Like "num_geopoint" Then
                                            boolnum_geopoint = True
                                        ElseIf NameBlock Like "delta_geopoint" Then
                                            boolDelta = True
                                        ElseIf NameBlock Like "r" Then
                                            boolR = True
                                            '==========================================================================
                                            'блоки с объектами
                                        ElseIf NameBlock Like "land_record" Then 'земельный участок
                                            AddresParcel = ""
                                            Cad_Number = ""
                                            StateParcel = ""
                                            Category = ""
                                            CodCategory = ""
                                            RazrIsp = ""
                                            RazrIspByDoc = ""
                                            CostParcel = ""
                                            GlobalArea = ""
                                            Cad_Number_Note = ""
                                            deltaArea = ""
                                            RightType = ""
                                            encumbranceName = ""
                                            okato = ""
                                            Cad_Number_Short = ""
                                            boolParcel = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "build_record" Then 'здание
                                            Cad_Number = ""
                                            AddresParcel = ""
                                            okato = ""
                                            RightType = ""
                                            encumbranceName = ""
                                            CostParcel = ""
                                            GlobalArea = ""
                                            Cad_Number_Note = ""
                                            purpose_build = ""
                                            pos_description_build = ""
                                            Assignatuin = ""
                                            MaterialWall = ""
                                            CountFloat = ""
                                            StateParcel = ""
                                            boolBuild = True
                                            purpose_build = ""
                                            Cad_Number_Short = ""
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "construction_record" Then 'конструкция
                                            Cad_Number = ""
                                            AddresParcel = ""
                                            CostParcel = ""
                                            GlobalArea = ""
                                            Cad_Number_Note = ""
                                            baseParameters = ""
                                            purpose_construction = ""
                                            Cad_Number_Short = ""
                                            boolconstruction = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "object_under_construction_record" Then 'объект незавершенного строительсва
                                            Cad_Number = ""
                                            AddresParcel = ""
                                            okato = ""
                                            RightType = ""
                                            encumbranceName = ""
                                            CostParcel = ""
                                            GlobalArea = ""
                                            Cad_Number_Note = ""
                                            purpose_build = ""
                                            pos_description_build = ""
                                            Assignatuin = ""
                                            MaterialWall = ""
                                            CountFloat = ""
                                            StateParcel = ""
                                            Cad_Number_Short = ""
                                            boolunderconstruction = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "inhabited_locality_boundary_record" Then 'граница населенного пункта
                                            RegNumbZone = ""
                                            type_boundary = ""
                                            cod_boundary = ""
                                            registr_data = ""
                                            boolcityboundary = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "municipal_boundary_record" Then 'граница административная
                                            RegNumbZone = ""
                                            type_boundary = ""
                                            cod_boundary = ""
                                            registr_data = ""
                                            boolboundary = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "spatial_data" Then 'квартал
                                            boolspatial_data = True
                                            boolKvartal = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "zones_and_territories_record" Or NameBlock Like "zones_and_territories" Then 'охранная зона
                                            type_zona = ""
                                            RegNumbZone = ""
                                            code_zona = ""
                                            registr_data = ""
                                            boolzone = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "coastline_record" Then 'береговая линия
                                            type_zona = ""
                                            RegNumbZone = ""
                                            code_zona = ""
                                            registr_data = ""
                                            ValTypeObject = ""
                                            water_object_name = ""
                                            codwater_object_type = ""
                                            typewater_object_type = ""
                                            type_boundary = ""
                                            boolcoastline = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "public_easement_record" Then 'охранная зона
                                            type_zona = ""
                                            RegNumbZone = ""
                                            code_zona = ""
                                            registr_data = ""
                                            boolzone = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "surveying_project_record" Then 'охранная зона
                                            type_zona = ""
                                            RegNumbZone = ""
                                            code_zona = ""
                                            registr_data = ""
                                            boolzone = True
                                            LocalID = rng.Next
                                            GlobalID = 0
                                            Erase ObjectCollections
                                            CountObjectCollection = 0
                                        ElseIf NameBlock Like "land_records" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "build_records" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "construction_records" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "object_under_construction_records" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "municipal_boundaries" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "inhabited_locality_boundaries" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            cod_boundary = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "zones_and_territories_boundaries" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            ValTypeObject = ""
                                            code_zona = ""
                                            cod_boundary = ""
                                            IndexZona = ""
                                            NumberZona = ""
                                            boolNastr = False
                                        ElseIf NameBlock Like "coastline_boundaries" Then
                                            LayerObject = "0"
                                            LineTypeObj = "Continuous"
                                            IndexColorObj = 256
                                            WidthrObj = 20
                                            ScLine = 1
                                            GlobalContLine = 0
                                            NamePSSetTable = "OBJECT"
                                            boolNastr = False
                                        End If
                                    '==========================================================================================
                                    'значения
                                    '==========================================================================================
                                    Case XmlNodeType.Text
                                        'кадастровый квартал
                                        If boolKvartalNumber = True Then 'номер кадастрового квартала
                                            CadastralBlock = reader.Value
                                            boolKvartalNumber = False
                                        ElseIf boolarea_quarter = True And boolGlobalArea = True Then 'площадь квартала
                                            Area_quarter = reader.Value
                                        ElseIf boolDateReg = True Then
                                            dateReg = reader.Value
                                            boolDateReg = False
                                        ElseIf boolarea_quarter = True And boolunit = True Then 'еденицы измерения площади квартала
                                            Dim unit As Integer = CInt(reader.Value)
                                            Area_quarter = Area_quarter & " " & FuncCadastr.FuncFindCadastrUnitByCode(unit)
                                        ElseIf boolcomdata = True And boolType = True Then 'тип объекта (земельный участок, здание и т.п.)
                                            CodTypeObject = reader.Value
                                            boolType = False
                                        ElseIf boolcomdata = True And boolval = True Then
                                            ValTypeObject = reader.Value
                                            boolval = False
                                        ElseIf boolCadNumber = True And boolcomdata = True Then 'кадастровый номер участка, сооружения, здания
                                            Cad_Number = reader.Value
                                            boolCadNumber = False
                                            'отделяем краткий номер
                                            Dim posN As Integer = Cad_Number.LastIndexOf(":")
                                            If posN > 6 Then
                                                Cad_Number_Short = Mid(Cad_Number, posN + 1)
                                            End If
                                        ElseIf boolcommon_land_cad_number = True And boolCadNumber = True Then 'кадастровый номер участка, сооружения, здания
                                            Cad_Number_Note = reader.Value
                                            boolCadNumber = False
                                        ElseIf boolNumberSubParcel = True Then 'номер части учаcтка
                                            NumberSubParcel = reader.Value
                                            boolNumberSubParcel = False
                                        ElseIf boolNumberContour = True Then 'номер контура
                                            NumberContour = reader.Value
                                            boolNumberContour = False
                                        ElseIf boolRegNumbZone = True Then 'номер территориальной зоны
                                            RegNumbZone = reader.Value
                                            boolRegNumbZone = False
                                        ElseIf boolSubTypeParcel = True And boolval = True Then 'тип участка 01-землепользование, многоконтурный и т.п.
                                            SubTypeParcel = reader.Value
                                            If SubTypeParcel Like "Многоконтурный участок" Or SubTypeParcel Like "Единое землепользование" Then
                                                GlobalID = rng.Next
                                            Else
                                                GlobalID = 0
                                            End If
                                            boolSubTypeParcel = False
                                        ElseIf boolStateParcel = True And boolCode = True Then 'статус участка
                                            StateParcel = reader.Value
                                        ElseIf boolCategory = True And boolCode = True Then 'категория земель
                                            CodCategory = reader.Value
                                        ElseIf boolCategory = True And boolval = True Then
                                            Category = reader.Value
                                        ElseIf boolpermitted_use = True And byRazrIspByDoc = True Then 'разрешенное использование
                                            RazrIspByDoc = reader.Value
                                            byRazrIspByDoc = False
                                        ElseIf boolCostParcel = True And boolval = True Then 'кадастровая стоимость
                                            CostParcel = reader.Value
                                            boolCostParcel = False
                                        ElseIf boolPart = True And boolGlobalArea = True And boolval = True Then
                                            areaPart = reader.Value
                                            boolval = False
                                            boolGlobalArea = False
                                        ElseIf boolGlobalArea = True And boolval = True Then 'площадь участка
                                            GlobalArea = reader.Value
                                            boolval = False
                                        ElseIf boolparams = True And boolGlobalArea = True And boolParcel = False Then 'площадь здания
                                            GlobalArea = reader.Value
                                            boolGlobalArea = False
                                        ElseIf boolGlobalArea = True And booldeltaArea = True Then 'погрешность площади
                                            deltaArea = reader.Value
                                            booldeltaArea = False
                                            '========================================================================
                                            'адрес
                                        ElseIf boolAdresParcel = True And boolokato = True Then
                                            okato = reader.Value
                                            boolokato = False
                                        ElseIf boolAdresParcel = True And boolkladr = True Then
                                            kladr = reader.Value
                                            boolkladr = False
                                        ElseIf boolAdresParcel = True And boolpostcode = True Then
                                            postcode = reader.Value
                                            boolpostcode = False
                                        ElseIf boolAdresParcel = True And boolregion = True And boolval = True Then
                                            region = reader.Value
                                            boolregion = False
                                            boolval = False
                                        ElseIf boolAdresParcel = True And booltypeDistrict = True Then
                                            typeDistrict = reader.Value
                                            booltypeDistrict = False
                                        ElseIf boolAdresParcel = True And boolnameDistrict = True Then
                                            nameDistrict = reader.Value
                                            boolnameDistrict = False
                                        ElseIf boolAdresParcel = True And booltypeLocality = True Then
                                            typeLocality = reader.Value
                                            booltypeLocality = False
                                        ElseIf boolAdresParcel = True And boolnameLocality = True Then
                                            nameLocalyty = reader.Value
                                            boolnameLocality = False
                                        ElseIf boolAdresParcel = True And booltypeStreet = True Then
                                            type_street = reader.Value
                                            booltypeStreet = False
                                        ElseIf boolAdresParcel = True And boolnameStreet = True Then
                                            name_street = reader.Value
                                            boolnameStreet = False
                                        ElseIf boolAdresParcel = True And booltypeLevel1 = True Then
                                            typeLevel1 = reader.Value
                                            booltypeLevel1 = False
                                        ElseIf boolAdresParcel = True And boolnameLevel1 = True Then
                                            nameLevel1 = reader.Value
                                            boolnameLevel1 = False
                                        ElseIf boolAdresParcel = True And booltypeLevel2 = True Then
                                            typeLevel2 = reader.Value
                                            booltypeLevel2 = False
                                        ElseIf boolAdresParcel = True And boolnameLevel2 = True Then
                                            nameLevel2 = reader.Value
                                            boolnameLevel2 = False
                                        ElseIf boolAdresParcel = True And booltypeLevel3 = True Then
                                            typeLevel3 = reader.Value
                                            booltypeLevel3 = False
                                        ElseIf boolAdresParcel = True And boolnameLevel3 = True Then
                                            nameLevel3 = reader.Value
                                            boolnameLevel3 = False
                                        ElseIf boolreadable_address = True Then
                                            readable_address = reader.Value
                                            boolreadable_address = False
                                            '============================================================================
                                        ElseIf boolRightType = True And boolval = True Then
                                            RightType = reader.Value
                                            boolRightType = False
                                            boolval = False
                                        ElseIf boolRight_holder = True And boolname = True Then
                                            ownerName = reader.Value
                                            boolOwner = False
                                            boolOwnerName = False
                                            boolRight_holder = False
                                        ElseIf boolresident = True And boolname = True Then
                                            encumbranceName = reader.Value
                                            boolresident = False
                                            boolname = False
                                            '============================================================
                                            'здания
                                        ElseIf boolpurpose = True And boolval = True And boolparams = True Then
                                            purpose_build = reader.Value 'назначение
                                            boolval = False
                                        ElseIf boolpurpose = True And boolCode = True And boolparams = True Then
                                            codeBuild = reader.Value 'код здания
                                            boolCode = False
                                        ElseIf boolposition_description = True Then
                                            pos_description_build = reader.Value 'описание местоположения
                                            boolposition_description = False
                                        ElseIf boolpermitted_use = True And boolname = True Then
                                            permitted_use = reader.Value
                                            '===================================================================
                                            'сооружения
                                        ElseIf boolparams = True And boolpurpose = True Then
                                            purpose_construction = reader.Value
                                            boolpurpose = False
                                        ElseIf boolbase_parameter = True And boolextension = True Then
                                            baseParameters = baseParameters & "Длина " & reader.Value & " "
                                            boolextension = False
                                        ElseIf boolbase_parameter = True And booldepth = True Then
                                            baseParameters = baseParameters & "Глубина " & reader.Value & " "
                                            booldepth = False
                                        ElseIf boolbase_parameter = True And booldepth = True Then
                                            baseParameters = baseParameters & "Глубина " & reader.Value & " "
                                            booldepth = False
                                        ElseIf boolbase_parameter = True And boolvolume = True Then
                                            baseParameters = baseParameters & "Объем " & reader.Value & " "
                                            boolvolume = False
                                        ElseIf boolbase_parameter = True And boolareaConstr = True Then
                                            baseParameters = baseParameters & "Общая площадь " & reader.Value & " "
                                            boolareaConstr = False
                                        ElseIf boolbase_parameter = True And boolbuilt_up_area = True Then
                                            baseParameters = baseParameters & "Общая площадь застройки " & reader.Value & " "
                                            boolbuilt_up_area = False
                                            '=======================================================================
                                            'границы
                                        ElseIf booltype_boundary = True And boolval = True Then
                                            type_boundary = reader.Value
                                            boolval = False
                                        ElseIf booltype_boundary = True And boolCode = True Then
                                            cod_boundary = reader.Value
                                            boolCode = False
                                            'береговые линии
                                        ElseIf boolwater_object_type = True And boolCode = True Then
                                            codwater_object_type = reader.Value
                                            boolCode = False
                                        ElseIf boolwater_object_type = True And boolType = True Then
                                            typewater_object_type = reader.Value
                                            boolType = False
                                        ElseIf boolwater_object_name = True Then
                                            water_object_name = reader.Value
                                            boolwater_object_name = False
                                            '==============================================================================
                                            'территориальные зоны
                                        ElseIf boolregistration_date = True And boolrecord_info = True Then
                                            registr_data = reader.Value
                                            boolregistration_date = False
                                        ElseIf boolType_zone = True And boolval = True Then 'тип ЗОУИТ
                                            type_zona = reader.Value
                                            boolval = False
                                        ElseIf boolType_zone = True And boolCode = True Then 'код ЗОУИТ
                                            code_zona = reader.Value
                                            boolCode = False
                                        ElseIf boolRegNumbZone = True Then
                                            RegNumbZone = reader.Value
                                            boolRegNumbZone = False
                                        ElseIf boolzone = True And boolIndexZona = True Then 'индекс зоны
                                            IndexZona = reader.Value
                                            boolIndexZona = False
                                        ElseIf boolzone = True And boolNumberZona = True Then 'номер зоны
                                            NumberZona = reader.Value
                                            boolNumberZona = False
                                            'координаты точек
                                            '======================================================================================
                                        ElseIf boolOrgNumb = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(1, countcoord) = reader.Value 'номер точки
                                                OrgNumb = reader.Value
                                            End If
                                            boolOrgNumb = False
                                        ElseIf boolX = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(2, countcoord) = reader.Value 'X
                                                X = reader.Value
                                            End If
                                            boolX = False
                                        ElseIf boolY = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(3, countcoord) = reader.Value 'Y
                                                Y = reader.Value
                                            End If
                                            boolY = False
                                        ElseIf boolDelta = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(4, countcoord) = reader.Value 'Delta
                                                Delta = reader.Value
                                            End If
                                            boolDelta = False
                                        ElseIf boolnum_geopoint = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(5, countcoord) = reader.Value 'старый номер
                                            End If
                                            boolnum_geopoint = False
                                        ElseIf boolgeopoint_zacrep = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(6, countcoord) = reader.Value 'тип закрепления
                                            End If
                                            boolnum_geopoint = False
                                        ElseIf boolGeopoint_opred = True And boolval = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(7, countcoord) = reader.Value
                                            End If
                                            boolGeopointValue = False
                                        ElseIf boolR = True Then
                                            If IsArray(NXYZ) = True Then
                                                NXYZ(8, countcoord) = reader.Value
                                                NXYZ(0, countcoord) = "Окружность"
                                            End If
                                            boolR = False
                                            '===================================================================================
                                            'реквизиты выписки
                                        ElseIf booldate_formation = True Then
                                            dataVP = reader.Value
                                            booldate_formation = False
                                        ElseIf boolregistration_number = True Then
                                            NumberVP = reader.Value
                                            boolregistration_number = False
                                        End If
                                    '====================================================================================================
                                    'блоки закрытия
                                    Case XmlNodeType.EndElement
                                        '=======================================================================
                                        'блок общих данных
                                        If NameBlock Like "common_data" Then
                                            boolcomdata = False
                                        ElseIf NameBlock Like "code" Then
                                            boolCode = False
                                        ElseIf NameBlock Like "value" Then
                                            boolval = False
                                        ElseIf NameBlock Like "type" Then
                                            boolType = False
                                        ElseIf NameBlock Like "unit" Then
                                            boolunit = False
                                        ElseIf NameBlock Like "name" Then
                                            boolname = False
                                            '=========================================================
                                        ElseIf NameBlock Like "area_quarter" Then
                                            boolarea_quarter = False
                                        ElseIf NameBlock Like "group_top_requisites" Then
                                            boolgroup_top_requisites = False
                                        ElseIf NameBlock Like "cad_number" Then
                                            boolCadNumber = False
                                        ElseIf NameBlock Like "common_land_cad_number" Then
                                            boolcommon_land_cad_number = False
                                        ElseIf NameBlock Like "subtype" Then
                                            boolSubTypeParcel = False
                                        ElseIf NameBlock Like "State" Then
                                            boolStateParcel = False
                                        ElseIf NameBlock Like "category" Then
                                            boolCategory = False
                                        ElseIf NameBlock Like "permitted_use" Then
                                            boolRazrIsp = False
                                        ElseIf NameBlock Like "cost" Then
                                            boolCostParcel = False
                                        ElseIf NameBlock Like "area" Then
                                            boolGlobalArea = False
                                            '=========================================================
                                            'адрес
                                        ElseIf NameBlock Like "address" Then
                                            boolAdresParcel = False
                                            If okato.Trim.Length > 0 Then
                                                AddresParcel = "ОКАТО " & okato & " "
                                                okato = ""
                                            End If
                                            If kladr.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & "КЛАДР " & kladr & " "
                                                kladr = ""
                                            End If
                                            If postcode.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & "Индекс " & postcode & " "
                                                postcode = ""
                                            End If
                                            If region.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & region & " "
                                                region = ""
                                            End If
                                            If typeDistrict.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & typeDistrict & " "
                                                typeDistrict = ""
                                            End If
                                            If nameDistrict.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & nameDistrict & " "
                                                nameDistrict = ""
                                            End If
                                            If typeLocality.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & typeLocality & " "
                                                typeLocality = ""
                                            End If
                                            If nameLocalyty.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & nameLocalyty & " "
                                                nameLocalyty = ""
                                            End If
                                            If type_street.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & type_street & " "
                                                type_street = ""
                                            End If
                                            If name_street.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & name_street & " "
                                                name_street = ""
                                            End If
                                            If typeLevel1.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & typeLevel1 & " "
                                                typeLevel1 = ""
                                            End If
                                            If nameLevel1.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & nameLevel1 & " "
                                                nameLevel1 = ""
                                            End If
                                            If typeLevel2.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & typeLevel2 & " "
                                                typeLevel2 = ""
                                            End If
                                            If nameLevel2.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & nameLevel2 & " "
                                                nameLevel2 = ""
                                            End If
                                            If typeLevel3.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & typeLevel3 & " "
                                                typeLevel3 = ""
                                            End If
                                            If nameLevel3.Trim.Length > 0 Then
                                                AddresParcel = AddresParcel & nameLevel3 & " "
                                                nameLevel3 = ""
                                            End If
                                        ElseIf NameBlock Like "okato" Then
                                            boolokato = False
                                        ElseIf NameBlock Like "kladr" Then
                                            boolkladr = False
                                        ElseIf NameBlock Like "region" Then
                                            boolregion = False
                                        ElseIf NameBlock Like "type_district" Then
                                            booltypeDistrict = False
                                        ElseIf NameBlock Like "name_district" Then
                                            boolnameDistrict = False
                                        ElseIf NameBlock Like "type_locality" Then
                                            booltypeLocality = False
                                        ElseIf NameBlock Like "name_locality" Then
                                            boolnameLocality = False
                                        ElseIf NameBlock Like "type_street" Then
                                            booltypeStreet = False
                                        ElseIf NameBlock Like "name_street" Then
                                            boolnameStreet = False
                                        ElseIf NameBlock Like "type_level1" Then
                                            booltypeLevel1 = False
                                        ElseIf NameBlock Like "name_level1" Then
                                            boolnameLevel1 = False
                                        ElseIf NameBlock Like "type_level2" Then
                                            booltypeLevel2 = False
                                        ElseIf NameBlock Like "name_level2" Then
                                            boolnameLevel2 = False
                                        ElseIf NameBlock Like "type_level3" Then
                                            booltypeLevel3 = False
                                        ElseIf NameBlock Like "name_level3" Then
                                            boolnameLevel3 = False
                                            '============================================================================
                                        ElseIf NameBlock Like "purpose" Then
                                            boolpurpose = False
                                        ElseIf NameBlock Like "params" Then
                                            boolparams = False
                                        ElseIf NameBlock Like "permitted_use" Then
                                            boolpermitted_use = False
                                        ElseIf NameBlock Like "base_parameter" Then
                                            boolbase_parameter = False
                                        ElseIf NameBlock Like "type_boundary" Then
                                            booltype_boundary = False
                                        ElseIf NameBlock Like "type_zone" Then
                                            boolType_zone = False
                                        ElseIf NameBlock Like "record_info" Then
                                            boolrecord_info = False
                                        ElseIf NameBlock Like "registration_date" Then
                                            boolregistration_date = False 'дата регистрации зоны
                                        ElseIf NameBlock Like "object_part" Then
                                            boolPart = False
                                            NumberSubParcel = ""
                                        ElseIf NameBlock Like "number_pp" Then 'номер контура
                                            boolNumberContour = False
                                        ElseIf NameBlock Like "area" Then
                                            boolGlobalArea = False
                                        ElseIf NameBlock Like "spatials_elements" Then
                                            boolSpatialElements = False
                                        ElseIf NameBlock Like "geopoint_opred" Then
                                            boolGeopoint_opred = False
                                        ElseIf NameBlock Like "x" Then
                                            boolX = False
                                        ElseIf NameBlock Like "y" Then
                                            boolY = False
                                        ElseIf NameBlock Like "geopoint_opred" Then
                                            boolGeopoint_opred = False
                                        ElseIf NameBlock Like "ord_nmb" Then
                                            boolOrgNumb = False
                                        ElseIf NameBlock Like "num_geopoint" Then
                                            boolnum_geopoint = False
                                        ElseIf NameBlock Like "delta_geopoint" Then
                                            boolDelta = False
                                        ElseIf NameBlock Like "r" Then
                                            boolR = False
                                            '===================================================================================================
                                            'конец блока с координатами контура
                                        ElseIf NameBlock Like "spatial_element" Then
                                            '=====================================================================================
                                            'делаем оформление
                                            If boolParcel = True Then
                                                'не требуется выводить участки
                                                If FormKPTImport.CheckBox1.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            ElseIf boolBuild = True Or boolunderconstruction = True Then
                                                If FormKPTImport.CheckBox2.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            ElseIf boolconstruction = True Then
                                                If FormKPTImport.CheckBox3.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            ElseIf boolboundary = True Or boolcityboundary Or boolcoastline Then
                                                If FormKPTImport.CheckBox4.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            ElseIf boolzone = True Then
                                                If FormKPTImport.CheckBox5.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            ElseIf boolKvartal = True Then
                                                If FormKPTImport.CheckBox6.Checked = False Then
                                                    Erase NXYZ
                                                    countcoord = -1
                                                    Continue While
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'блок лля проверки дубликатов объектов
                                            If boolParcel = True Or boolBuild = True Or boolconstruction = True Or boolunderconstruction = True Then
                                                If IsArray(ArrayOldCadastrNumbers) = True Then
                                                    For k As Integer = 0 To ArrayOldCadastrNumbers.Length - 1
                                                        Dim OldNumber As String = ArrayOldCadastrNumbers(k)
                                                        If Cad_Number Like OldNumber Then
                                                            Erase NXYZ
                                                            countcoord = -1
                                                            Continue While
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            If boolcityboundary = True Or boolboundary = True Or boolzone = True Or boolcoastline = True Then
                                                If IsArray(ArrayOldCadastrRegNumbers) = True Then
                                                    For k As Integer = 0 To ArrayOldCadastrRegNumbers.Length - 1
                                                        Dim OldNumber As String = ArrayOldCadastrRegNumbers(k)
                                                        If RegNumbZone Like OldNumber Then
                                                            Erase NXYZ
                                                            countcoord = -1
                                                            Continue While
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            If boolKvartal = True Then
                                                If IsArray(ArrayOldCadastrKVNumbers) = True Then
                                                    For k As Integer = 0 To ArrayOldCadastrKVNumbers.Length - 1
                                                        Dim OldNumber As String = ArrayOldCadastrKVNumbers(k)
                                                        If CadastralBlock Like OldNumber Then
                                                            Erase NXYZ
                                                            countcoord = -1
                                                            Continue While
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            Dim codObject = CodTypeObject
                                            'если это участок
                                            If boolParcel = True Then
                                                'если это часть участка
                                                If boolPart = True Then
                                                    codObject = "002001001000_"
                                                End If
                                                'если это здание 
                                            ElseIf boolBuild = True Then
                                                If boolPart = True Then
                                                    codObject = "002001002000_"
                                                Else
                                                    If codeBuild.Trim.Length > 0 Then
                                                        codObject = codeBuild
                                                    End If
                                                End If
                                                'если это конструкция 
                                            ElseIf boolconstruction = True Then
                                                If boolPart = True Then
                                                    codObject = "002001004000_"
                                                Else
                                                    If purpose_construction.Trim.Length > 0 Then
                                                        If purpose_construction.IndexOf("газо") > -1 Or purpose_construction.IndexOf("Газо") > -1 Then
                                                            codObject = "002001004000_5"
                                                        ElseIf purpose_construction.IndexOf("водо") > -1 Or purpose_construction.IndexOf("Водо") > -1 Then
                                                            codObject = "002001004000_4"
                                                        ElseIf purpose_construction.IndexOf("связ") > -1 Or purpose_construction.IndexOf("Саяз") > -1 Then
                                                            codObject = "002001004000_2"
                                                        ElseIf purpose_construction.IndexOf("канализ") > -1 Or purpose_construction.IndexOf("Канализ") > -1 Then
                                                            codObject = "002001004000_1"
                                                        ElseIf purpose_construction.IndexOf("тепл") > -1 Or purpose_construction.IndexOf("Тепл") > -1 Then
                                                            codObject = "002001004000_6"
                                                        ElseIf purpose_construction.IndexOf("электр") > -1 Or purpose_construction.IndexOf("Электр") > -1 Then
                                                            codObject = "002001004000_3"
                                                        ElseIf purpose_construction.IndexOf("доро") > -1 Or purpose_construction.IndexOf("Доро") > -1 Then
                                                            codObject = "002001004000_7"
                                                        End If
                                                    End If
                                                End If
                                                'усли это квартал
                                            ElseIf boolspatial_data = True Then
                                                codObject = "000000000001"
                                            ElseIf boolzone = True Then
                                                If code_zona.Trim.Length > 0 Then
                                                    codObject = code_zona.Trim
                                                Else
                                                    If cod_boundary.Trim.Length > 0 Then
                                                        codObject = cod_boundary
                                                    End If
                                                End If
                                            ElseIf cod_boundary.Trim.Length > 0 Then
                                                codObject = cod_boundary
                                            End If
                                            'если предыдущий объект по коду был другим, то меняем установки
                                            If Not (codObject Like OldCodObject) Then
                                                'ищем настройки в массиве
                                                Dim indN As Integer = -1
                                                If IsArray(ArrayLineProperties) = True Then
                                                    For j As Integer = 0 To ArrayLineProperties.GetUpperBound(1)
                                                        Dim ArrayStr As String() = ArrayLineProperties(8, j).Split(";")
                                                        If IsArray(ArrayStr) = True Then
                                                            For k As Integer = 0 To ArrayStr.Length - 1
                                                                If ArrayStr(k).IndexOf("Code=") > -1 Then
                                                                    Dim UserCod As String = Mid(ArrayStr(k), 6)
                                                                    If UserCod Like codObject Then
                                                                        indN = j
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Next k
                                                            If indN > -1 Then
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next j
                                                End If
                                                If indN > -1 Then
                                                    If acLineTypTbl.Has(ArrayLineProperties(1, indN)) = True Then
                                                        LineTypeObj = ArrayLineProperties(1, indN)
                                                    Else
                                                        LineTypeObj = "Continuous"
                                                    End If
                                                    If acLyrTbl.Has(ArrayLineProperties(2, indN)) = True Then
                                                        LayerObject = ArrayLineProperties(2, indN)
                                                    Else
                                                        LayerObject = "0"
                                                    End If
                                                    If dictPropSetDefTemp.Has(ArrayLineProperties(7, indN), acTrans) = True Then
                                                        NamePSSetTable = ArrayLineProperties(7, indN)
                                                        IDPSTAble = FuncPropertySet.FuncReturnIDDictSetProperties(NamePSSetTable)
                                                    End If
                                                    '3. Находим цвет линии
                                                    IndexColorObj = Val(ArrayLineProperties(3, indN))
                                                    If boolParcel = True And boolSubTypeParcel = False Then
                                                        If CodCategory Like "003001000000" Then
                                                            IndexColorObj = 40
                                                        ElseIf CodCategory Like "003003000000" Then
                                                            IndexColorObj = 12
                                                        ElseIf CodCategory Like "003004000000" Then
                                                            IndexColorObj = 230
                                                        ElseIf CodCategory Like "003005000000" Then
                                                            IndexColorObj = 94
                                                        ElseIf CodCategory Like "003007000000" Then
                                                            IndexColorObj = 200
                                                        ElseIf CodCategory Like "003008000000" Then
                                                            IndexColorObj = 123
                                                        End If
                                                    End If
                                                    If IndexColorObj Like "ПоСлою" Then
                                                        IndexColorObj = 256
                                                    ElseIf IndexColorObj Like "ПоБлоку" Then
                                                        IndexColorObj = 0
                                                    Else
                                                        If IndexColorObj < 0 Then
                                                            IndexColorObj = 0
                                                        ElseIf IndexColorObj > 256 Then
                                                            IndexColorObj = 256
                                                        End If
                                                    End If
                                                    '4. Находим вес линии
                                                    WidthrObj = Val(ArrayLineProperties(4, indN))
                                                    If WidthrObj = 0 Then WidthrObj = 20
                                                    '5. формируем масштаб линии
                                                    ScLine = Val(ArrayLineProperties(5, indN))
                                                    If ScLine <= 0 Then ScLine = 1
                                                    Dim PLTSCAle As Integer = Application.GetSystemVariable("PSLTSCALE") 'маштаб автоматически подстраивается под видовой экран
                                                    Dim LTSCAle As Double = Application.GetSystemVariable("LTSCALE") 'коэффициент глобального масштаба
                                                    ScLine = ScLine / LTSCAle 'делаем поправку за глобальный масштаб
                                                    If PLTSCAle = 0 Then 'если опция одинакового показа линии для любого масштаба отключена
                                                        Dim ScaleLocal As Double = M
                                                        ScLine = ScLine * ScaleLocal
                                                    End If
                                                    '5. формируем глобальную ширину
                                                    GlobalContLine = Val(ArrayLineProperties(6, indN))
                                                    OldCodObject = codObject
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'вставляем полилинию или окружность в качестве объекта
                                            If IsArray(NXYZ) = True Then
                                                Dim acadObj As Autodesk.AutoCAD.DatabaseServices.Entity = Nothing
                                                Dim areaObj As Double = 0
                                                Dim boolOstrov As Boolean = False
                                                If NXYZ.GetUpperBound(1) > 0 Then
                                                    'если число вершин больше 1 тогда рисуем полилинию
                                                    Dim IDPLine As ObjectId = FuncCadastr.FuncDrawPolyline(NXYZ)
                                                    If IsNothing(IDPLine) = False Then
                                                        If IDPLine.IsNull = False Then
                                                            'делаем оформление
                                                            Dim Pline As Polyline = acTrans.GetObject(IDPLine, OpenMode.ForWrite)
                                                            Pline.Linetype = LineTypeObj
                                                            Pline.Layer = LayerObject
                                                            Pline.ColorIndex = IndexColorObj
                                                            Pline.LineWeight = WidthrObj
                                                            Pline.LinetypeScale = ScLine
                                                            Pline.ConstantWidth = GlobalContLine
                                                            If Pline.Closed = True Then
                                                                areaObj = Pline.Area
                                                            End If
                                                            acadObj = Pline
                                                        End If
                                                    End If
                                                Else
                                                    '==================================================================================
                                                    'вставляем окружность нужного радиуса
                                                    Dim acCirc As Circle = New Circle()
                                                    acCirc.SetDatabaseDefaults()
                                                    acCirc.Center = New Point3d(Val(NXYZ(3, 0)), Val(NXYZ(2, 0)), 0)
                                                    r = NXYZ(8, 0)
                                                    If r <= 0 Then
                                                        r = 1
                                                    End If
                                                    acCirc.Radius = r
                                                    acBlkTblRec.AppendEntity(acCirc)
                                                    acTrans.AddNewlyCreatedDBObject(acCirc, True)
                                                    acCirc.Linetype = LineTypeObj
                                                    acCirc.Layer = LayerObject
                                                    acCirc.ColorIndex = IndexColorObj
                                                    acCirc.LineWeight = WidthrObj
                                                    acCirc.LinetypeScale = ScLine
                                                    acadObj = acCirc
                                                End If
                                                Erase NXYZ
                                                countcoord = -1
                                                '0-тип объекта, 1-обозначение, 2-Handle
                                                ReDim Preserve ObjectCollections(4, CountObjectCollection)
                                                If boolPart = True Then
                                                    ObjectCollections(0, CountObjectCollection) = "Часть"
                                                    ObjectCollections(1, CountObjectCollection) = NumberSubParcel
                                                ElseIf GlobalID > 0 Then
                                                    ObjectCollections(0, CountObjectCollection) = "Контур"
                                                    ObjectCollections(1, CountObjectCollection) = NumberContour
                                                Else
                                                    ObjectCollections(0, CountObjectCollection) = ""
                                                    ObjectCollections(1, CountObjectCollection) = ""
                                                End If
                                                ObjectCollections(2, CountObjectCollection) = acadObj.Handle.ToString 'ID объекта
                                                ObjectCollections(3, CountObjectCollection) = areaObj
                                                ObjectCollections(4, CountObjectCollection) = NamePSSetTable
                                                CountObjectCollection += 1
                                            End If
                                        ElseIf NameBlock Like "land_record" Then
                                            boolParcel = False
                                            OldCodObject = ""
                                            If IsArray(ObjectCollections) = True Then
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 3)
                                                End If
                                                For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                                    Dim typeObject As String = ObjectCollections(0, k)
                                                    Dim hgObject As String = ObjectCollections(2, k)
                                                    Dim idObject As ObjectId = AcadFunction.FuncReturnObjectID(hgObject)
                                                    LocalID = rng.Next
                                                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObject.GetObject(OpenMode.ForWrite)
                                                    Dim InsertPointText As Point3d = New Point3d
                                                    If TypeOf acEnt Is Polyline Then
                                                        Dim Pline As Polyline = acEnt
                                                        InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                        If InsertPointText.X = 0 And InsertPointText.Y = 0 Then
                                                            InsertPointText = Pline.StartPoint
                                                        End If
                                                    ElseIf TypeOf acEnt Is Circle Then
                                                        Dim Circ As Circle = acEnt
                                                        InsertPointText = Circ.Center
                                                    End If
                                                    NamePSSetTable = ObjectCollections(4, k)
                                                    IDPSTAble = FuncPropertySet.FuncReturnIDDictSetProperties(NamePSSetTable)
                                                    If typeObject Like "Часть" Then
                                                        '==========================================================================================================
                                                        'присваиваем характеристики
                                                        If IsNothing(IDPSTAble) = False Then
                                                            If IDPSTAble.IsNull = False Then
                                                                Dim ArrayData As String(,) = Nothing
                                                                ReDim Preserve ArrayData(2, 15)
                                                                ArrayData(0, 0) = NamePSSetTable
                                                                ArrayData(1, 0) = "Cad_number"
                                                                ArrayData(2, 0) = Cad_Number & "/" & ObjectCollections(1, k)

                                                                ArrayData(0, 1) = NamePSSetTable
                                                                ArrayData(1, 1) = "Part_number"
                                                                ArrayData(2, 1) = ObjectCollections(1, k)

                                                                ArrayData(0, 2) = NamePSSetTable
                                                                ArrayData(1, 2) = "Desk"
                                                                ArrayData(2, 2) = ""

                                                                ArrayData(0, 3) = NamePSSetTable
                                                                ArrayData(1, 3) = "Area"
                                                                ArrayData(2, 3) = areaPart

                                                                ArrayData(0, 4) = NamePSSetTable
                                                                ArrayData(1, 4) = "Inaccuracy"
                                                                ArrayData(2, 4) = ""

                                                                ArrayData(0, 13) = NamePSSetTable
                                                                ArrayData(1, 13) = "Encumbrance"
                                                                ArrayData(2, 13) = ""

                                                                ArrayData(0, 14) = NamePSSetTable
                                                                ArrayData(1, 14) = "LOCALID"
                                                                ArrayData(2, 14) = LocalID

                                                                ArrayData(0, 15) = NamePSSetTable
                                                                ArrayData(1, 15) = "GlobalID"
                                                                ArrayData(2, 15) = GlobalID

                                                                Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                If boolPSObject = True Then
                                                                    Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                End If
                                                            End If
                                                        End If
                                                        Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number & "/" & ObjectCollections(1, k), HtextParcel, LayerObject, idTTextStyle)
                                                    Else
                                                        'контур участка (делаем проверку на внитренний островок)
                                                        Dim boolOstrov As Boolean = False
                                                        If k + 1 < ObjectCollections.GetUpperBound(1) Then
                                                            For k2 = k + 1 To ObjectCollections.GetUpperBound(1) - 1
                                                                If ObjectCollections(0, k2) Like "Часть" Then
                                                                    Continue For
                                                                Else
                                                                    Dim hgObject2 As String = ObjectCollections(2, k2)
                                                                    Dim idObject2 As ObjectId = AcadFunction.FuncReturnObjectID(hgObject2)
                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObject2.GetObject(OpenMode.ForWrite)
                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                        Dim FirstPline As Polyline = acEnt
                                                                        Dim secondPline As Polyline = acEnt2
                                                                        Dim boolCross As Boolean = MathFunction.FuncCrossPointPolygon(FirstPline, secondPline.StartPoint)
                                                                        If boolCross = True Then
                                                                            boolOstrov = True
                                                                            Dim lIDStar As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "LOCALID")
                                                                            If IsNumeric(lIDStar) = False Then
                                                                                lIDStar = 0
                                                                            End If
                                                                            LocalID = lIDStar
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next k2
                                                        End If
                                                        If boolOstrov = True Then
                                                            If IsNothing(IDPSSetTableOstrov) = False Then
                                                                If IDPSSetTableOstrov.IsNull = False Then
                                                                    Dim ArrayData As String(,) = Nothing
                                                                    ReDim Preserve ArrayData(2, 2)
                                                                    ArrayData(0, 0) = NamePSSetTableOstrov
                                                                    ArrayData(1, 0) = "IDKOUNTER"
                                                                    ArrayData(2, 0) = LocalID

                                                                    ArrayData(0, 1) = NamePSSetTableOstrov
                                                                    ArrayData(1, 1) = "LOCALID"
                                                                    ArrayData(2, 1) = rng.Next

                                                                    ArrayData(0, 2) = NamePSSetTableOstrov
                                                                    ArrayData(1, 2) = "ConditionalNumber"
                                                                    ArrayData(2, 2) = k

                                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSSetTableOstrov)
                                                                    If boolPSObject = True Then
                                                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            '==========================================================================================================
                                                            'присваиваем характеристики
                                                            'земельный участок
                                                            If IsNothing(IDPSTAble) = False Then
                                                                If IDPSTAble.IsNull = False Then
                                                                    Dim numKounter As String = ""
                                                                    If ObjectCollections(1, k).Trim.Length > 0 Then
                                                                        numKounter = "(" & ObjectCollections(1, k) & ")"
                                                                    End If
                                                                    Dim ArrayData As String(,) = Nothing
                                                                    ReDim Preserve ArrayData(2, 15)
                                                                    ArrayData(0, 0) = NamePSSetTable
                                                                    ArrayData(1, 0) = "Cad_number"
                                                                    If Cad_Number_Note.Trim.Length > 0 Then
                                                                        ArrayData(2, 0) = Cad_Number_Note & " вх.н. " & Cad_Number & numKounter
                                                                    Else
                                                                        ArrayData(2, 0) = Cad_Number & numKounter
                                                                    End If

                                                                    ArrayData(0, 1) = NamePSSetTable
                                                                    ArrayData(1, 1) = "Short_cad_number"
                                                                    ArrayData(2, 1) = Cad_Number_Short & numKounter

                                                                    ArrayData(0, 2) = NamePSSetTable
                                                                    ArrayData(1, 2) = "Type"
                                                                    ArrayData(2, 2) = SubTypeParcel

                                                                    ArrayData(0, 3) = NamePSSetTable
                                                                    ArrayData(1, 3) = "Area"
                                                                    ArrayData(2, 3) = GlobalArea

                                                                    ArrayData(0, 4) = NamePSSetTable
                                                                    ArrayData(1, 4) = "Inaccuracy"
                                                                    ArrayData(2, 4) = deltaArea

                                                                    ArrayData(0, 5) = NamePSSetTable
                                                                    ArrayData(1, 5) = "OKATO"
                                                                    ArrayData(2, 5) = AddresParcel

                                                                    ArrayData(0, 6) = NamePSSetTable
                                                                    ArrayData(1, 6) = "Adress"
                                                                    ArrayData(2, 6) = readable_address

                                                                    ArrayData(0, 7) = NamePSSetTable
                                                                    ArrayData(1, 7) = "Category"
                                                                    ArrayData(2, 7) = Category

                                                                    ArrayData(0, 8) = NamePSSetTable
                                                                    ArrayData(1, 8) = "UtilizationByDoc"
                                                                    ArrayData(2, 8) = RazrIspByDoc

                                                                    ArrayData(0, 9) = NamePSSetTable
                                                                    ArrayData(1, 9) = "Utilization"
                                                                    ArrayData(2, 9) = RazrIsp

                                                                    ArrayData(0, 10) = NamePSSetTable
                                                                    ArrayData(1, 10) = "CadastralCost"
                                                                    ArrayData(2, 10) = CostParcel

                                                                    ArrayData(0, 11) = NamePSSetTable
                                                                    ArrayData(1, 11) = "Right_type"
                                                                    ArrayData(2, 11) = RightType

                                                                    ArrayData(0, 12) = NamePSSetTable
                                                                    ArrayData(1, 12) = "Patronymic"
                                                                    ArrayData(2, 12) = ownerName

                                                                    ArrayData(0, 13) = NamePSSetTable
                                                                    ArrayData(1, 13) = "Encumbrance"
                                                                    ArrayData(2, 13) = ""

                                                                    ArrayData(0, 14) = NamePSSetTable
                                                                    ArrayData(1, 14) = "LOCALID"
                                                                    ArrayData(2, 14) = LocalID

                                                                    ArrayData(0, 15) = NamePSSetTable
                                                                    ArrayData(1, 15) = "GlobalID"
                                                                    ArrayData(2, 15) = GlobalID

                                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                    If boolPSObject = True Then
                                                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                    End If
                                                                End If
                                                            End If
                                                            Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number_Short, HtextParcel, LayerObject, idTTextStyle)
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                            Erase ObjectCollections
                                            Cad_Number_Note = ""
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        ElseIf NameBlock Like "build_record" Then
                                            boolBuild = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line2
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line2:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number_Short & numKounter, HtextBuild, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 15)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Cad_number"
                                                                        ArrayData(2, 0) = Cad_Number & numKounter

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Short_cad_number"
                                                                        ArrayData(2, 1) = Cad_Number_Short & numKounter

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "Assignation"
                                                                        ArrayData(2, 2) = permitted_use

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "Area"
                                                                        ArrayData(2, 3) = GlobalArea

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "OKATO"
                                                                        ArrayData(2, 4) = AddresParcel

                                                                        ArrayData(0, 5) = NamePSSetTable
                                                                        ArrayData(1, 5) = "Adress"
                                                                        ArrayData(2, 5) = readable_address

                                                                        ArrayData(0, 6) = NamePSSetTable
                                                                        ArrayData(1, 6) = "CadastralCost"
                                                                        ArrayData(2, 6) = CostParcel

                                                                        ArrayData(0, 7) = NamePSSetTable
                                                                        ArrayData(1, 7) = "KeyParameter"
                                                                        ArrayData(2, 7) = baseParameters

                                                                        ArrayData(0, 8) = NamePSSetTable
                                                                        ArrayData(1, 8) = "State"
                                                                        ArrayData(2, 8) = purpose_build

                                                                        ArrayData(0, 9) = NamePSSetTable
                                                                        ArrayData(1, 9) = "Right_type"
                                                                        ArrayData(2, 9) = RightType

                                                                        ArrayData(0, 10) = NamePSSetTable
                                                                        ArrayData(1, 10) = "Patronymic"
                                                                        ArrayData(2, 10) = ownerName

                                                                        ArrayData(0, 11) = NamePSSetTable
                                                                        ArrayData(1, 11) = "Encumbrance"
                                                                        ArrayData(2, 11) = ""

                                                                        ArrayData(0, 12) = NamePSSetTable
                                                                        ArrayData(1, 12) = "LOCALID"
                                                                        ArrayData(2, 12) = LocalID

                                                                        ArrayData(0, 13) = NamePSSetTable
                                                                        ArrayData(1, 13) = "GlobalID"
                                                                        ArrayData(2, 13) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            Cad_Number_Short = ""
                                            purpose_build = ""
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        ElseIf NameBlock Like "municipal_boundary_record" Then
                                            boolboundary = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line3
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line3:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, RegNumbZone, HtextZone, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 4)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Reg_numb_border"
                                                                        ArrayData(2, 0) = RegNumbZone

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Type_boundary"
                                                                        ArrayData(2, 1) = type_boundary

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "DateRegistration"
                                                                        ArrayData(2, 2) = registr_data

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "LOCALID"
                                                                        ArrayData(2, 3) = LocalID

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "GlobalID"
                                                                        ArrayData(2, 4) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextZone, LayerObject, idTTextStyle)
                                        ElseIf NameBlock Like "inhabited_locality_boundary_record" Then
                                            boolcityboundary = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line4
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line4:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, RegNumbZone, HtextZone, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 4)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Reg_numb_border"
                                                                        ArrayData(2, 0) = RegNumbZone

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Type_boundary"
                                                                        ArrayData(2, 1) = type_boundary

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "DateRegistration"
                                                                        ArrayData(2, 2) = registr_data

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "LOCALID"
                                                                        ArrayData(2, 3) = LocalID

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "GlobalID"
                                                                        ArrayData(2, 4) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextZone, LayerObject, idTTextStyle)
                                        ElseIf NameBlock Like "construction_record" Then
                                            boolconstruction = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) >= 0 Then
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                                        'читаем характеристики первого контура
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                        countArrayTempParcel += 1
                                                    Next k
                                                End If
                                            End If
                                            Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                            If IsArray(ArrayTempParcel) = True Then
                                                If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                    GlobalID = rng.Next
                                                Else
                                                    GlobalID = 0
                                                End If
                                                For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                    Dim HgStr As String = ArrayTempParcel(0, k1)
                                                    Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                    Dim numKounter As String = ArrayTempParcel(2, k1)
                                                    Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                    LocalID = rng.Next
                                                    If IsNothing(idObj) = False Then
                                                        If idObj.IsNull = False Then
                                                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                            If TypeOf acEnt Is Polyline Then
                                                                Dim Pline As Polyline = acEnt
                                                                InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                            ElseIf TypeOf acEnt Is Circle Then
                                                                Dim Circ As Circle = acEnt
                                                                InsertPointText = Circ.Center
                                                            End If
                                                            Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number_Short & numKounter, HtextConstruction, LayerObject, idTTextStyle)
                                                            '==========================================================================================================
                                                            'присваиваем характеристики
                                                            'земельный участок
                                                            If IsNothing(IDPSTAble) = False Then
                                                                If IDPSTAble.IsNull = False Then
                                                                    Dim ArrayData As String(,) = Nothing
                                                                    ReDim Preserve ArrayData(2, 11)
                                                                    ArrayData(0, 0) = NamePSSetTable
                                                                    ArrayData(1, 0) = "Cad_number"
                                                                    ArrayData(2, 0) = Cad_Number & numKounter

                                                                    ArrayData(0, 1) = NamePSSetTable
                                                                    ArrayData(1, 1) = "Short_cad_number"
                                                                    ArrayData(2, 1) = Cad_Number_Short & numKounter

                                                                    ArrayData(0, 2) = NamePSSetTable
                                                                    ArrayData(1, 2) = "Assignation"
                                                                    ArrayData(2, 2) = purpose_construction

                                                                    ArrayData(0, 3) = NamePSSetTable
                                                                    ArrayData(1, 3) = "OKATO"
                                                                    ArrayData(2, 3) = AddresParcel

                                                                    ArrayData(0, 4) = NamePSSetTable
                                                                    ArrayData(1, 4) = "Adress"
                                                                    ArrayData(2, 4) = readable_address

                                                                    ArrayData(0, 5) = NamePSSetTable
                                                                    ArrayData(1, 5) = "CadastralCost"
                                                                    ArrayData(2, 5) = CostParcel

                                                                    ArrayData(0, 6) = NamePSSetTable
                                                                    ArrayData(1, 6) = "KeyParameter"
                                                                    ArrayData(2, 6) = baseParameters

                                                                    ArrayData(0, 7) = NamePSSetTable
                                                                    ArrayData(1, 7) = "Right_type"
                                                                    ArrayData(2, 7) = RightType

                                                                    ArrayData(0, 8) = NamePSSetTable
                                                                    ArrayData(1, 8) = "Patronymic"
                                                                    ArrayData(2, 8) = ownerName

                                                                    ArrayData(0, 9) = NamePSSetTable
                                                                    ArrayData(1, 9) = "Encumbrance"
                                                                    ArrayData(2, 9) = ""

                                                                    ArrayData(0, 10) = NamePSSetTable
                                                                    ArrayData(1, 10) = "LOCALID"
                                                                    ArrayData(2, 10) = LocalID

                                                                    ArrayData(0, 11) = NamePSSetTable
                                                                    ArrayData(1, 11) = "GlobalID"
                                                                    ArrayData(2, 11) = GlobalID

                                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                    If boolPSObject = True Then
                                                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next k1
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextConstruction, LayerObject, idTTextStyle)
                                            Cad_Number_Short = ""
                                        ElseIf NameBlock Like "object_under_construction_record" Then
                                            'блок кода для определения внутренних островков
                                            OldCodObject = ""
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) >= 0 Then
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1)
                                                        'читаем характеристики первого контура
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                        countArrayTempParcel += 1
                                                    Next k
                                                End If
                                            End If
                                            Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                            If IsArray(ArrayTempParcel) = True Then
                                                If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                    GlobalID = rng.Next
                                                Else
                                                    GlobalID = 0
                                                End If
                                                For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                    Dim HgStr As String = ArrayTempParcel(0, k1)
                                                    Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                    Dim numKounter As String = ArrayTempParcel(2, k1)
                                                    Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                    LocalID = rng.Next
                                                    If IsNothing(idObj) = False Then
                                                        If idObj.IsNull = False Then
                                                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                            If TypeOf acEnt Is Polyline Then
                                                                Dim Pline As Polyline = acEnt
                                                                InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                            ElseIf TypeOf acEnt Is Circle Then
                                                                Dim Circ As Circle = acEnt
                                                                InsertPointText = Circ.Center
                                                            End If
                                                            Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number_Short & numKounter, HtextBuild, LayerObject, idTTextStyle)
                                                            '==========================================================================================================
                                                            'присваиваем характеристики
                                                            'земельный участок
                                                            If IsNothing(IDPSTAble) = False Then
                                                                If IDPSTAble.IsNull = False Then
                                                                    Dim ArrayData As String(,) = Nothing
                                                                    ReDim Preserve ArrayData(2, 11)
                                                                    ArrayData(0, 0) = NamePSSetTable
                                                                    ArrayData(1, 0) = "Cad_number"
                                                                    ArrayData(2, 0) = Cad_Number & numKounter

                                                                    ArrayData(0, 1) = NamePSSetTable
                                                                    ArrayData(1, 1) = "Short_cad_number"
                                                                    ArrayData(2, 1) = Cad_Number_Short & numKounter

                                                                    ArrayData(0, 2) = NamePSSetTable
                                                                    ArrayData(1, 2) = "Assignation"
                                                                    ArrayData(2, 2) = permitted_use

                                                                    ArrayData(0, 3) = NamePSSetTable
                                                                    ArrayData(1, 3) = "OKATO"
                                                                    ArrayData(2, 3) = AddresParcel

                                                                    ArrayData(0, 4) = NamePSSetTable
                                                                    ArrayData(1, 4) = "Adress"
                                                                    ArrayData(2, 4) = readable_address

                                                                    ArrayData(0, 5) = NamePSSetTable
                                                                    ArrayData(1, 5) = "CadastralCost"
                                                                    ArrayData(2, 5) = CostParcel

                                                                    ArrayData(0, 6) = NamePSSetTable
                                                                    ArrayData(1, 6) = "KeyParameter"
                                                                    ArrayData(2, 6) = baseParameters

                                                                    ArrayData(0, 7) = NamePSSetTable
                                                                    ArrayData(1, 7) = "Right_type"
                                                                    ArrayData(2, 7) = RightType

                                                                    ArrayData(0, 8) = NamePSSetTable
                                                                    ArrayData(1, 8) = "Patronymic"
                                                                    ArrayData(2, 8) = ownerName

                                                                    ArrayData(0, 9) = NamePSSetTable
                                                                    ArrayData(1, 9) = "Encumbrance"
                                                                    ArrayData(2, 9) = ""

                                                                    ArrayData(0, 10) = NamePSSetTable
                                                                    ArrayData(1, 10) = "LOCALID"
                                                                    ArrayData(2, 10) = LocalID

                                                                    ArrayData(0, 11) = NamePSSetTable
                                                                    ArrayData(1, 11) = "GlobalID"
                                                                    ArrayData(2, 11) = GlobalID

                                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                    If boolPSObject = True Then
                                                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next k1
                                            End If
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextBuild, LayerObject, idTTextStyle)
                                            boolunderconstruction = False
                                            Cad_Number_Short = ""
                                            Erase ObjectCollections
                                        ElseIf NameBlock Like "zones_and_territories_record" Or NameBlock Like "zones_and_territories" Then
                                            boolzone = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line6
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line6:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, RegNumbZone, HtextZone, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 7)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Reg_numb_border"
                                                                        ArrayData(2, 0) = RegNumbZone

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Type_boundary"
                                                                        ArrayData(2, 1) = type_boundary

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "Type_zone"
                                                                        ArrayData(2, 2) = type_zona

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "Index"
                                                                        ArrayData(2, 3) = IndexZona

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "Number"
                                                                        ArrayData(2, 4) = NumberZona

                                                                        ArrayData(0, 5) = NamePSSetTable
                                                                        ArrayData(1, 5) = "DateRegistration"
                                                                        ArrayData(2, 5) = registr_data

                                                                        ArrayData(0, 6) = NamePSSetTable
                                                                        ArrayData(1, 6) = "LOCALID"
                                                                        ArrayData(2, 6) = LocalID

                                                                        ArrayData(0, 7) = NamePSSetTable
                                                                        ArrayData(1, 7) = "GlobalID"
                                                                        ArrayData(2, 7) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextZone, LayerObject, idTTextStyle)
                                        ElseIf NameBlock Like "public_easement_record" Then
                                            boolzone = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line16
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line16:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, RegNumbZone, HtextZone, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 7)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Reg_numb_border"
                                                                        ArrayData(2, 0) = RegNumbZone

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Type_boundary"
                                                                        ArrayData(2, 1) = type_boundary

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "Type_zone"
                                                                        ArrayData(2, 2) = type_zona

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "Index"
                                                                        ArrayData(2, 3) = IndexZona

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "Number"
                                                                        ArrayData(2, 4) = NumberZona

                                                                        ArrayData(0, 5) = NamePSSetTable
                                                                        ArrayData(1, 5) = "DateRegistration"
                                                                        ArrayData(2, 5) = registr_data

                                                                        ArrayData(0, 6) = NamePSSetTable
                                                                        ArrayData(1, 6) = "LOCALID"
                                                                        ArrayData(2, 6) = LocalID

                                                                        ArrayData(0, 7) = NamePSSetTable
                                                                        ArrayData(1, 7) = "GlobalID"
                                                                        ArrayData(2, 7) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'кадастровый квартал
                                        ElseIf NameBlock Like "spatial_data" Then
                                            boolspatial_data = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line8
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line8:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, CadastralBlock, HtextKV, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 3)
                                                                        ArrayData(0, 0) = "LAND_CADASTRAL_DIVISION"
                                                                        ArrayData(1, 0) = "Cadastral_number"
                                                                        ArrayData(2, 0) = CadastralBlock

                                                                        ArrayData(0, 1) = "LAND_CADASTRAL_DIVISION"
                                                                        ArrayData(1, 1) = "Area"
                                                                        ArrayData(2, 1) = Area_quarter

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "LOCALID"
                                                                        ArrayData(2, 2) = LocalID

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "GlobalID"
                                                                        ArrayData(2, 3) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextKV, LayerObject, idTTextStyle)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'береговые линии
                                        ElseIf NameBlock Like "coastline_record" Then
                                            boolcoastline = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo Line7
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
Line7:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If
                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ArrayTempParcel(2, k1)
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, RegNumbZone & numKounter, HtextZone, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 6)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Reg_numb_border"
                                                                        ArrayData(2, 0) = RegNumbZone & numKounter

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Type_boundary"
                                                                        ArrayData(2, 1) = type_boundary

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "Name"
                                                                        ArrayData(2, 2) = water_object_name

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "Water_object_type"
                                                                        ArrayData(2, 3) = typewater_object_type

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "DateRegistration"
                                                                        ArrayData(2, 4) = registr_data

                                                                        ArrayData(0, 5) = NamePSSetTable
                                                                        ArrayData(1, 5) = "LOCALID"
                                                                        ArrayData(2, 5) = LocalID

                                                                        ArrayData(0, 6) = NamePSSetTable
                                                                        ArrayData(1, 6) = "GlobalID"
                                                                        ArrayData(2, 6) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                            'Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, kanNumStr, HtextZone, LayerObject, idTTextStyle)
                                            water_object_name = ""
                                            typewater_object_type = ""
                                        ElseIf NameBlock Like "object_part" Then
                                            boolPart = False
                                            OldCodObject = ""
                                            'блок кода для определения внутренних островков
                                            Dim ArrayTempParcel As String(,) = Nothing
                                            Dim countArrayTempParcel As Integer = 0
                                            If IsArray(ObjectCollections) = True Then
                                                'сортируем соллекцию в случае если там более 1 элемента
                                                If ObjectCollections.GetUpperBound(1) > 0 Then
                                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ObjectCollections, 0)
                                                    'заносим в массив последний элемент с наибольшей площадью
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                    'начинаем проход начиная с первого элемента, а сравнивать начинаем начиная с последнего элемента
                                                    For k As Integer = 0 To ObjectCollections.GetUpperBound(1) - 1
                                                        'читаем характеристики первого контура
                                                        Dim boolIns As Boolean = False 'это не внутренний островок
                                                        Dim numKounter As String = ObjectCollections(1, k) 'обозначение контура
                                                        Dim HgStr As String = ObjectCollections(2, k) 'handle контура
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr) 'находим его id
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim FirstPline As Polyline = acEnt
                                                                    'ищем внутреннюю точку
                                                                    Dim CrossPointFirstPline As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(FirstPline)
                                                                    'проверяем вдруг это островок какого либо другого контура
                                                                    For k2 = ObjectCollections.GetUpperBound(1) To k + 1 Step -1
                                                                        Dim numKounter2 As String = ObjectCollections(1, k2)
                                                                        Dim HgStr2 As String = ObjectCollections(2, k2)
                                                                        If numKounter2 = numKounter Then
                                                                            'делаем проверку (вдруг это островок)
                                                                            Dim idObj2 As ObjectId = AcadFunction.FuncReturnObjectID(HgStr2)
                                                                            If IsNothing(idObj2) = False Then
                                                                                If idObj2.IsNull = False Then
                                                                                    Dim acEnt2 As Autodesk.AutoCAD.DatabaseServices.Entity = idObj2.GetObject(OpenMode.ForRead)
                                                                                    If TypeOf acEnt2 Is Polyline Then
                                                                                        Dim SecondPline As Polyline = acEnt2
                                                                                        boolIns = MathFunction.FuncCrossPointPolygon(SecondPline, CrossPointFirstPline)
                                                                                        If boolIns = True Then
                                                                                            'это внутренний островок
                                                                                            Dim boolrec As Boolean = False
                                                                                            For k3 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                                                                Dim tempHgFirst As String = ArrayTempParcel(0, k3)
                                                                                                If tempHgFirst Like HgStr2 Then
                                                                                                    ArrayTempParcel(1, k3) = ArrayTempParcel(1, k3) & " " & HgStr 'островок
                                                                                                    boolrec = True
                                                                                                    GoTo LinePart1
                                                                                                End If
                                                                                            Next k3
                                                                                            If boolrec = False Then
                                                                                                ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                                                ArrayTempParcel(0, countArrayTempParcel) = HgStr2 'записываем handle основного контура
                                                                                                ArrayTempParcel(1, countArrayTempParcel) = HgStr 'островок
                                                                                                ArrayTempParcel(2, countArrayTempParcel) = numKounter2 'обозначение контура
                                                                                                countArrayTempParcel += 1
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Next k2
                                                                    If boolIns = False Then
                                                                        ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                        ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                        ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                        ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                        countArrayTempParcel += 1
                                                                    End If
                                                                Else
                                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                                    ArrayTempParcel(0, countArrayTempParcel) = HgStr 'записываем handle основного контура
                                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                                    ArrayTempParcel(2, countArrayTempParcel) = numKounter 'обозначение контура
                                                                    countArrayTempParcel += 1
                                                                End If
                                                            End If
                                                        End If
LinePart1:
                                                    Next k
                                                ElseIf ObjectCollections.GetUpperBound(1) = 0 Then
                                                    ReDim Preserve ArrayTempParcel(2, countArrayTempParcel)
                                                    ArrayTempParcel(0, countArrayTempParcel) = ObjectCollections(2, ObjectCollections.GetUpperBound(1)) 'записываем handle основного контура
                                                    ArrayTempParcel(1, countArrayTempParcel) = "" 'островок
                                                    ArrayTempParcel(2, countArrayTempParcel) = ObjectCollections(1, ObjectCollections.GetUpperBound(1)) 'обозначение контура
                                                    countArrayTempParcel += 1
                                                End If

                                                Dim InsertPointText As Point3d = New Point3d(0, 0, -1) 'точка вставки текста
                                                If IsArray(ArrayTempParcel) = True Then
                                                    If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                        GlobalID = rng.Next
                                                    Else
                                                        GlobalID = 0
                                                    End If
                                                    For k1 As Integer = 0 To ArrayTempParcel.GetUpperBound(1)
                                                        Dim HgStr As String = ArrayTempParcel(0, k1)
                                                        Dim HgOstrovStr As String = ArrayTempParcel(1, k1)
                                                        Dim numKounter As String = ""
                                                        If ArrayTempParcel.GetUpperBound(1) > 0 Then
                                                            numKounter = "\" & ArrayTempParcel(2, k1)
                                                        End If
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(HgStr)
                                                        LocalID = rng.Next
                                                        If IsNothing(idObj) = False Then
                                                            If idObj.IsNull = False Then
                                                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = idObj.GetObject(OpenMode.ForRead)
                                                                If TypeOf acEnt Is Polyline Then
                                                                    Dim Pline As Polyline = acEnt
                                                                    InsertPointText = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                                                ElseIf TypeOf acEnt Is Circle Then
                                                                    Dim Circ As Circle = acEnt
                                                                    InsertPointText = Circ.Center
                                                                End If
                                                                Dim PosText As Boolean = AcadFunction.FuncDrawMText(InsertPointText, Cad_Number_Short & numKounter, HtextParcel, LayerObject, idTTextStyle)
                                                                '==========================================================================================================
                                                                'присваиваем характеристики
                                                                'земельный участок
                                                                If IsNothing(IDPSTAble) = False Then
                                                                    If IDPSTAble.IsNull = False Then
                                                                        Dim ArrayData As String(,) = Nothing
                                                                        ReDim Preserve ArrayData(2, 15)
                                                                        ArrayData(0, 0) = NamePSSetTable
                                                                        ArrayData(1, 0) = "Cad_number"
                                                                        ArrayData(2, 0) = Cad_Number & numKounter

                                                                        ArrayData(0, 1) = NamePSSetTable
                                                                        ArrayData(1, 1) = "Part_number"
                                                                        ArrayData(2, 1) = ArrayTempParcel(2, k1)

                                                                        ArrayData(0, 2) = NamePSSetTable
                                                                        ArrayData(1, 2) = "Desk"
                                                                        ArrayData(2, 2) = ""

                                                                        ArrayData(0, 3) = NamePSSetTable
                                                                        ArrayData(1, 3) = "Area"
                                                                        ArrayData(2, 3) = GlobalArea

                                                                        ArrayData(0, 4) = NamePSSetTable
                                                                        ArrayData(1, 4) = "Inaccuracy"
                                                                        ArrayData(2, 4) = deltaArea

                                                                        ArrayData(0, 13) = NamePSSetTable
                                                                        ArrayData(1, 13) = "Encumbrance"
                                                                        ArrayData(2, 13) = ""

                                                                        ArrayData(0, 14) = NamePSSetTable
                                                                        ArrayData(1, 14) = "LOCALID"
                                                                        ArrayData(2, 14) = LocalID

                                                                        ArrayData(0, 15) = NamePSSetTable
                                                                        ArrayData(1, 15) = "GlobalID"
                                                                        ArrayData(2, 15) = GlobalID

                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEnt, IDPSTAble)
                                                                        If boolPSObject = True Then
                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        If HgOstrovStr.Trim.Length > 0 Then
                                                            Dim ArrayStr As String() = HgOstrovStr.Trim.Split()
                                                            If IsArray(ArrayStr) = True Then
                                                                For k2 As Integer = 0 To ArrayStr.Length() - 1
                                                                    Dim HgOstrov As String = ArrayStr(k2)
                                                                    If HgOstrov.Length > 2 Then
                                                                        Dim idObjOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HgOstrov)
                                                                        If IsNothing(idObjOstrov) = False Then
                                                                            If idObjOstrov.IsNull = False Then
                                                                                Dim acEntOstrov As Autodesk.AutoCAD.DatabaseServices.Entity = idObjOstrov.GetObject(OpenMode.ForRead)
                                                                                If IsNothing(IDPSSetTableOstrov) = False Then
                                                                                    If IDPSSetTableOstrov.IsNull = False Then
                                                                                        Dim ArrayData As String(,) = Nothing
                                                                                        ReDim Preserve ArrayData(2, 2)
                                                                                        ArrayData(0, 0) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 0) = "IDKOUNTER"
                                                                                        ArrayData(2, 0) = LocalID

                                                                                        ArrayData(0, 1) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 1) = "LOCALID"
                                                                                        ArrayData(2, 1) = rng.Next

                                                                                        ArrayData(0, 2) = NamePSSetTableOstrov
                                                                                        ArrayData(1, 2) = "ConditionalNumber"
                                                                                        ArrayData(2, 2) = k2 + 1

                                                                                        Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(acEntOstrov, IDPSSetTableOstrov)
                                                                                        If boolPSObject = True Then
                                                                                            Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEntOstrov, ArrayData)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            End If
                                            Erase ObjectCollections
                                        ElseIf NameBlock Like "water_object_type" Then
                                            boolwater_object_type = False
                                            Erase ObjectCollections
                                        ElseIf NameBlock Like "water_object_name" Then
                                            boolwater_object_name = False
                                            Erase ObjectCollections
                                        ElseIf NameBlock Like "extract_cadastral_plan_territory" Then
                                            Erase ObjectCollections
                                        ElseIf NameBlock Like "land_records" Then
                                            CodCategory = ""
                                            Erase ObjectCollections
                                        End If
                                End Select
                            End While
                        Catch ex As System.Exception
                            MsgBox(ex.ToString & vbLf & "Ошибка обработки данных. Макро прервано." & vbLf & LineErrNumber)
                            acTrans.Commit()
                            locker.Dispose()
                            Exit Sub
                        End Try
                    Next
                    acTrans.Commit()
                End Using
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Ошибка обработки данных. Макро прервано." & vbLf & LineErrNumber)
                locker.Dispose()
                Exit Sub
            End Try
        End Using
    End Sub

    '=========================================================================================================
    'отрисовка и нумерация полигона на основе выписки из xml-файла (Технокад)
    <CommandMethod("RosreestrImportVp")>
    Public Sub RosreestrImportVp()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем документ
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем масштаб аннотаций (текущий)
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'получаем доступ к шаблону
        Dim PutchTemplate As String = ""
        Dim NameFileTemplate As String = ""
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        Dim parentPatch1 As String = System.IO.Directory.GetParent(PathFile).FullName 'отделяем имя файла от папки
        Dim parentPatch2 As String = System.IO.Directory.GetParent(parentPatch1).FullName 'определяем имя родительской папки для папки где находится сборка
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If System.IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In System.IO.Directory.GetFiles(PutchTemplate)
                Try
                    Dim File1 As String = System.IO.Path.GetFileName(fileDwg)
                    Dim ext As String = System.IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        NameFileTemplate = File1
                        Exit For
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'общие флаги
        Dim flagParcel As Boolean = False
        Dim flagSubParcel As Boolean = False
        Dim flagCounter As Boolean = False
        Dim flagEntParcel As Boolean = False
        Dim flagBuilding As Boolean = False
        Dim flagSubBuilding As Boolean = False
        Dim flagConstruction As Boolean = False
        Dim flagSubConstruction As Boolean = False
        Dim flagUncompleted As Boolean = False
        Dim flagSubUncompleted As Boolean = False
        Dim flagZone As Boolean = False
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальное значение номера точки
        Dim Numb As Integer = 1 'номер
        Dim NXYZ As String(,) = Nothing 'массив для координит точек, их номеров и дельт
        Dim CadastralObjectCollections As ObjectIdCollection = New ObjectIdCollection
        Dim bool_x As Boolean = False
        Dim bool_y As Boolean = False
        Dim bool_ord_nmb As Boolean = False
        Dim bool_num_geopoint As Boolean = False
        Dim bool_delta_geopoint As Boolean = False
        Dim bool_r As Boolean = False
        Dim bool_spatial_element As Boolean = False
        Dim bool_ordinates As Boolean = False

        Dim StrObjectType As String = ""
        Dim flag As Boolean = False
        Dim countCoord As Integer = 0
        'кадастровый квартал
        Dim boolKv As Boolean = False
        Dim KVNumber As String = ""
        'адрес земельного участка
        Dim FlagAddress As Boolean = False 'флаг для считывания адреса если он лож то данные note не читаются
        Dim District1 As String = "" 'наименование района
        Dim TipDistrict As String = "" 'тип 
        Dim City As String = ""
        Dim TipCity As String = ""
        Dim Street As String = ""
        Dim TipStreet As String = ""
        Dim Level1 As String = ""
        Dim TypeLevel1 As String = ""
        Dim Level2 As String = ""
        Dim TypeLevel2 As String = ""
        Dim Level3 As String = ""
        Dim TypeLevel3 As String = ""
        Dim TipBuiild As String = ""
        Dim ValBuild As String = ""
        Dim UrbanBuild As String = ""
        Dim TipUrbanBuild As String = ""
        Dim SovietVillige As String = ""
        Dim TypeSovetVillige As String = ""
        Dim Locality As String = ""
        Dim TipLocality As String = ""
        Dim Apartament As String = ""
        Dim TipApartament As String = ""
        Dim AddresParcel As String = "" 'адрес
        Dim bool_AddresParcelNote As Boolean = False
        Dim AddresParcelNote As String = "" 'неструктурируемый адрес
        'права
        Dim FlagRight As Boolean = False 'флаг для считывания права собственности
        Dim RightType As String = "" 'вид права
        Dim boolRightType As Boolean = False
        Dim boolPerson As Boolean = False
        Dim FamilyName As String = ""
        Dim boolFM As Boolean = False
        Dim FirstName As String = ""
        Dim boolFirstNAme As Boolean = False
        Dim Patronymic As String = ""
        Dim boolPatrony As Boolean = False
        Dim boolRightName As Boolean = False
        'Номера кадастровые
        Dim bool_cad_number As Boolean = False
        Dim NumberParcel As String = "" 'кадастровый номер участка (полный)
        Dim NParcel As String = "" 'кадастровый номер участка (сокращенный)
        Dim NumberContour As String = "" 'номер контура
        'номера ОКС для участка
        Dim boolInnerKN As Boolean = False
        Dim InnerKN As String = ""
        'площади
        Dim boolAreaParcel As Boolean = False
        Dim AreaParcel As String = 0 'площадь участка общая
        Dim boolInaccuracy As Boolean = False
        Dim AreaInaccuracy As String = "" 'допуск по площади
        Dim boolAreaUnit As Boolean = False
        Dim AreaUnit As String = ""
        'характеристика участка
        Dim TypeParcel As String = ""
        Dim bool_commom_data As Boolean = False
        Dim StateiParcel As String = "" 'тип участка 01-ранее учтенный 06-учтенный ....
        Dim DateCreateParcel As String = "" 'дата постановки участка на кадастровый учет
        Dim boolName As Boolean = False 'временный учтенный и т.п
        Dim NameParcel As String = ""
        Dim boolCategory As Boolean = False
        Dim Category As String = "" 'категория земель
        Dim bool_permitted_use As Boolean = False
        Dim boolRazrIsp As Boolean = False
        Dim RazrIsp As String = "" 'разрешенное использование 
        Dim RazrIspByDoc As String = ""
        Dim bool_CadastralCost As Boolean = False
        Dim CadastralCost As String = "" 'кадастровая стоимость
        Dim boolSpecialNote As Boolean = False
        Dim SpecialNote As String = "" 'поле примечание
        'общие характеристики
        Dim bool_Type As Boolean = False
        Dim bool_code As Boolean = False
        Dim codeObject As String = ""
        Dim bool_value As Boolean = False
        Dim valueObject As String = ""
        Dim bool_subtype As Boolean = False
        Dim bool_included_object As Boolean = False
        Dim included_cad_number As String = ""
        Dim bool_record_info As Boolean = False
        Dim bool_registration_date As Boolean = False
        Dim registration_date As String = ""
        Dim bool_quarter_cad_number As Boolean = False
        Dim quarter_cad_number As String = ""
        Dim ExtractNumber As String = ""
        Dim ExtractData As String = ""
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'характеристики ОКС
        Dim boolObjectType As Boolean = False
        Dim boolAssignBuild As Boolean = False
        Dim StrAssignBuild As String = ""
        Dim boolMaterialWall As Boolean = False
        Dim StrMaterialWal As String = ""
        Dim StrYearBuilt As String = "" 'год завершения строительства
        Dim StrYearUsed As String = "" 'год ввода в эксплуатацию
        Dim StrCountFloor As String = "" 'количество этажей
        Dim StrCountUnderFloor As String = "" 'количество подземных этажей
        Dim TypeKeyParameterConstr As String = ""
        Dim TypeValParameter As String = ""
        Dim CadastralCostParcel As String = ""
        Dim boolParentNumb As Boolean = False
        Dim boolAssignNAme As Boolean = False
        Dim AssignConstr As String = ""
        Dim boolR As Boolean = False
        Dim ColObjHatch As ObjectIdCollection = New ObjectIdCollection
        'Характеристики линии по умолчанию
        Dim NameLayerLine As String = "0"
        Dim NameTypeLine As String = "Continuous" 'тип линии
        Dim ColorLine As Integer = 7
        Dim TablePSLine As String = ""
        Dim WLine As Double = 20
        Dim GLine As Double = 0
        Dim ScaleLine As Double = 1
        Dim NoteLine As String = ""
        'Характеристики блока по умолчанию
        Dim NameBlk As String = ""
        Dim LayerBlk As Integer = "0"
        Dim ColorBlk As Integer = 0
        Dim ScalePoint As Scale3d = New Scale3d(1, 1, 1) 'задаем масштаб блоков
        Dim TypeLineBlk As String = "ПоСлою"
        Dim WLineBlk As Double = 20
        Dim SizeAttr As Double = 2
        Dim IDTablePSBlk As ObjectId = Nothing
        Dim DistAttr As Double = 1 '
        Dim NoteBlk As String = ""
        'массив с индивидуальными характеристиками
        Dim ArrayLineProperties As String(,) = Nothing
        Dim CountArrayLineProperties As Integer = 0
        Dim ArrayBlkProperties As String(,) = Nothing
        Dim CountArrayBlkPropertie As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
            Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(acCurDb.LinetypeTableId, OpenMode.ForRead)
            Dim dictPropSetDef = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(acCurDb)
            '==================================================================================================================
            'ЧИТАЕМ ВЫПИСКУ
            'выбор файла
            Dim NameFiles As String() = {""}
            Dim fileDialog1 As OpenFileDialog = New OpenFileDialog("Открыть файл xml", "", "xml", "xml", OpenFileDialogFlags.AllowMultiple)
            Dim FileRez As String = fileDialog1.ShowDialog()
            NameFiles = fileDialog1.GetFilenames
            If IsNothing(NameFiles) = True Then
                Exit Sub
            End If
            '===================================================================================================================
            Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тел функции
                For i As Integer = 0 To NameFiles.GetUpperBound(0)
                    Dim FileName As String = NameFiles(i)
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'читаем файл
                    Dim reader As XmlTextReader = New XmlTextReader(FileName)
                    While reader.Read()
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'земельные участки
                                If NameBlock Like "Parcel" Or NameBlock Like "land_record" Then
                                    flagParcel = True 'признак того что это выписка об участке
                                    CountArrayLineProperties = FuncListLineObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayLineProperties, "ЕГРН-Кадастровые участки")
                                    If IsArray(ArrayLineProperties) = True Then
                                        If ArrayLineProperties.GetUpperBound(1) > 0 Then
                                            '(0-desk, 1-name, 2-layer, 3-color, 4-width, 5-M, 6-GWidth, 7-PS, 8-Note)
                                            For j As Integer = 0 To ArrayLineProperties.GetUpperBound(1)
                                                NoteLine = ArrayLineProperties(8, j)
                                                If IsNothing(NoteLine) = False Then
                                                    If NoteLine.IndexOf("002001001000_") Then
                                                        Dim TempTypeLine = ArrayLineProperties(1, j).Trim
                                                        If acLineTypTbl.Has(TempTypeLine) = True Then
                                                            NameTypeLine = TempTypeLine
                                                            If acLyrTbl.Has(ArrayLineProperties(2, j).Trim) Then
                                                                NameLayerLine = ArrayLineProperties(2, j).Trim
                                                            End If
                                                        Else
                                                            Continue For
                                                        End If
                                                        ColorLine = Val(ArrayLineProperties(3, j).Trim)
                                                        If ColorLine < 0 Or ColorLine > 256 Then
                                                            ColorLine = 7
                                                        End If
                                                        WLine = Val(ArrayLineProperties(4, j).Trim)
                                                        If WLine < 0 Then
                                                            WLine = 20
                                                        End If
                                                        Dim MLine As Double = Val(ArrayLineProperties(5, j).Trim)
                                                        If MLine < 0 Then
                                                            ScaleLine = MLine
                                                        End If
                                                        Dim GLineT As Double = Val(ArrayLineProperties(6, j).Trim)
                                                        If GLineT < 0 Then
                                                            GLine = 0
                                                        End If
                                                        Dim PsStr As String = ArrayLineProperties(7, j).Trim
                                                        If dictPropSetDef.Has(PsStr, acTrans) = True Then
                                                            IDTablePSBlk = FuncPropertySet.FuncReturnIDDictSetProperties(PsStr)
                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                    CountArrayBlkPropertie = FuncListBlkObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayBlkProperties, "ЕГРН-Кадастровые участки")
                                    '=============================================================================================================================================================
                                    'настройки для блоков
                                    If IsArray(ArrayBlkProperties) = True Then
                                        If ArrayBlkProperties.GetUpperBound(1) > 0 Then
                                            Dim boolCopyStyle As Boolean = FuncCopyBlkStylesDestDraw(NameFileTemplate, ArrayBlkProperties, False)
                                            '(0-desk, 1-name, 2-layer, 3-color, 4-размер атрибута, 5-M, 6-WidthLine, 7-PS, 8-Note)
                                            For j As Integer = 0 To ArrayBlkProperties.GetUpperBound(1)
                                                NameBlock = ArrayBlkProperties(1, j).Trim
                                                If acBlkTbl.Has(NameBlock) = True Then
                                                    NameBlk = NameBlock
                                                    If acLyrTbl.Has(ArrayBlkProperties(2, j).Trim) Then
                                                        LayerBlk = ArrayBlkProperties(2, j).Trim
                                                    End If
                                                Else
                                                    Continue For
                                                End If
                                                ColorBlk = Val(ArrayBlkProperties(3, j).Trim)
                                                If ColorBlk < 0 Or ColorBlk > 256 Then
                                                    ColorBlk = 0
                                                End If
                                                SizeAttr = Val(ArrayBlkProperties(4, j).Trim)
                                                If SizeAttr <= 0 Then
                                                    SizeAttr = 2
                                                End If
                                                Dim MBlk As Double = Val(ArrayBlkProperties(5, j).Trim)
                                                If MBlk > 0 Then
                                                    ScalePoint = New Scale3d(MBlk, MBlk, MBlk) 'задаем масштаб блоков
                                                End If
                                                Dim PsStr As String = ArrayBlkProperties(7, j).Trim
                                                If dictPropSetDef.Has(PsStr, acTrans) = True Then
                                                    IDTablePSBlk = FuncPropertySet.FuncReturnIDDictSetProperties(PsStr)
                                                End If
                                                NoteBlk = ArrayBlkProperties(8, j)
                                                DistAttr = Math.Sqrt(SizeAttr ^ 2 + SizeAttr ^ 2) 'дистанция от точки до атрибута
                                                Exit For
                                            Next j
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'здания
                                ElseIf NameBlock Like "Building" Then
                                    flagBuilding = True
                                    CountArrayLineProperties = FuncListLineObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayLineProperties, "ЕГРН-Здания (строения)")
                                    CountArrayBlkPropertie = FuncListBlkObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayBlkProperties, "ЕГРН-Здания (строения)")
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'объекты незавершенного строительства
                                ElseIf NameBlock Like "Uncompleted" Then
                                    flagUncompleted = True
                                    CountArrayLineProperties = FuncListLineObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayLineProperties, "ЕГРН-Здания (строения)")
                                    CountArrayBlkPropertie = FuncListBlkObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayBlkProperties, "ЕГРН-Здания (строения)")
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'сооружения
                                ElseIf NameBlock Like "Construction" Then
                                    flagConstruction = True
                                    CountArrayLineProperties = FuncListLineObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayLineProperties, "ЕГРН-Сооружения")
                                    CountArrayBlkPropertie = FuncListBlkObjectDataTableDestDWG(NameFileTemplate, "Земельный кадастр", ArrayBlkProperties, "ЕГРН-Сооружения")
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'части земельных участков
                                ElseIf NameBlock Like "SubParcel" Then
                                    flagSubParcel = True
                                    'части зданий
                                ElseIf NameBlock Like "SubBuilding" Then
                                    flagSubBuilding = True
                                    'части сооружений
                                ElseIf NameBlock Like "SubConstruction" Then
                                    flagSubConstruction = True
                                    'единое землепользование
                                ElseIf NameBlock Like "EntryParcel" Then
                                    flagEntParcel = True
                                    'контур участка
                                ElseIf NameBlock Like "Contour" Then
                                    flagCounter = True
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                ElseIf NameBlock Like "CadastralBlock" Then
                                    boolKv = True
                                ElseIf NameBlock Like "ObjectType" Then
                                    boolObjectType = True
                                ElseIf NameBlock Like "common_data" Then
                                    bool_commom_data = True
                                ElseIf NameBlock Like "cad_number" And bool_included_object = False Then
                                    bool_cad_number = True
                                ElseIf NameBlock Like "quarter_cad_number" Then
                                    bool_quarter_cad_number = True
                                ElseIf NameBlock Like "type" Then
                                    bool_Type = True
                                ElseIf NameBlock Like "code" Then
                                    bool_code = True
                                ElseIf NameBlock Like "value" Then
                                    bool_value = True
                                ElseIf NameBlock Like "subtype" Then
                                    bool_subtype = True
                                ElseIf NameBlock Like "included_object" Then
                                    bool_included_object = True
                                ElseIf NameBlock Like "Category" Then
                                    boolCategory = True
                                ElseIf NameBlock Like "InnerCadastralNumbers" Then 'для считывания кадастрового номера участка для ОКС
                                    boolInnerKN = True
                                ElseIf NameBlock Like "ParentCadastralNumbers" And flagParcel = False Then
                                    boolInnerKN = True
                                ElseIf NameBlock Like "CadastralNumber" And boolInnerKN = True Then
                                    boolParentNumb = True
                                ElseIf NameBlock Like "cad_number" And bool_included_object = True Then
                                    boolParentNumb = True
                                ElseIf NameBlock Like "permitted_use" Then
                                    bool_permitted_use = True
                                ElseIf NameBlock Like "by_document" And bool_permitted_use = True Then
                                    boolRazrIsp = True
                                ElseIf NameBlock Like "Name" And FlagRight = False And FlagAddress = False Then
                                    If flagParcel = True Or flagBuilding = True Or flagConstruction = True Or flagUncompleted = True Then
                                        boolName = True
                                    End If
                                ElseIf NameBlock Like "Area" Then
                                    boolAreaParcel = True
                                ElseIf NameBlock Like "Inaccuracy" Then
                                    boolInaccuracy = True
                                ElseIf NameBlock Like "Unit" And boolAreaParcel = True Then
                                    boolAreaUnit = True
                                ElseIf NameBlock Like "Address" Then
                                    FlagAddress = True
                                ElseIf NameBlock Like "adrs:Note" Or NameBlock Like "readable_address" Then
                                    bool_AddresParcelNote = True
                                ElseIf NameBlock Like "cost" Then
                                    bool_CadastralCost = True
                                    'здания и строения
                                ElseIf NameBlock Like "AssignationBuilding" Then
                                    boolAssignBuild = True
                                ElseIf NameBlock Like "AssignationName" Then
                                    boolAssignNAme = True
                                    'права
                                ElseIf NameBlock Like "Right" Or NameBlock Like "right_data" Then
                                    FlagRight = True
                                ElseIf NameBlock Like "Type" And FlagRight = True Then
                                    boolRightType = True
                                ElseIf NameBlock Like "right_type" And FlagRight = True Then
                                    boolRightType = True
                                ElseIf NameBlock Like "Name" And FlagRight = True Then
                                    boolRightName = True
                                ElseIf NameBlock Like "Person" Then
                                    boolPerson = True
                                ElseIf NameBlock Like "FamilyName" And boolPerson = True Then
                                    boolFM = True
                                ElseIf NameBlock Like "FirstName" And boolPerson = True Then
                                    boolFirstNAme = True
                                ElseIf NameBlock Like "Patronymic" And boolPerson = True Then
                                    boolPatrony = True
                                ElseIf NameBlock Like "spa:SpatialElement" Or NameBlock Like "SpatialElement" Then
                                    bool_spatial_element = True
                                ElseIf NameBlock Like "spa:SpelementUnit" Or NameBlock Like "SpelementUnit" Or NameBlock Like "ordinates" Then
                                    bool_ordinates = True
                                    Erase NXYZ
                                    countCoord = 0
                                    ReDim Preserve NXYZ(7, countCoord)
                                    NXYZ(7, countCoord) = reader.Value
                                ElseIf NameBlock Like "x" Then
                                    bool_x = True
                                ElseIf NameBlock Like "y" Then
                                    bool_y = True
                                ElseIf NameBlock Like "ord_nmb" Then
                                    bool_ord_nmb = True
                                ElseIf NameBlock Like "num_geopoint" Then
                                    bool_num_geopoint = True
                                ElseIf NameBlock Like "delta_geopoint" Then
                                    bool_delta_geopoint = True
                                ElseIf NameBlock Like "ns3:R" Or NameBlock Like "R" Or NameBlock Like "r" Then
                                    boolR = True
                                ElseIf NameBlock Like "SpecialNote" Then
                                    boolSpecialNote = True
                                End If
                                '====================================================================================================
                                'поля с атрибутами
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        'считываем кадастровые номера
                                        If nameN Like "CadastralNumber" Then 'скойство кадастровый номер
                                            'запись в таблицу кадастрового номера
                                            NumberParcel = reader.Value
                                        ElseIf nameN Like "State" Then
                                            'временный, учтенный,
                                            StateiParcel = FuncCadastr.FuncCadastrStateParcel(reader.Value)
                                        ElseIf nameN Like "DateCreated" Then
                                            DateCreateParcel = reader.Value
                                        ElseIf nameN Like "DateRemoved" Then
                                            DateCreateParcel = reader.Value
                                        ElseIf nameN Like "DateExpiry" Then
                                            DateCreateParcel = reader.Value
                                        ElseIf nameN Like "NumberRecord" Then
                                            NumberContour = reader.Value
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'адрес по КЛАДР
                                        ElseIf nameN Like "Name" And FlagAddress = True Then
                                            If NameBlock Like "adrs:District" Then
                                                District1 = reader.Value
                                            ElseIf NameBlock Like "adrs:City" Then
                                                City = reader.Value
                                            ElseIf NameBlock Like "adrs:UrbanDistrict" Then
                                                UrbanBuild = reader.Value
                                            ElseIf NameBlock Like "adrs:SovietVillage" Then
                                                SovietVillige = reader.Value & " " & SovietVillige
                                            ElseIf NameBlock Like "adrs:Locality" Then
                                                Locality = reader.Value & " " & Locality
                                            ElseIf NameBlock Like "adrs:Street" Then
                                                Street = reader.Value & " " & Street
                                            ElseIf NameBlock Like "adrs:Level1" Then
                                                Level1 = reader.Value
                                            ElseIf NameBlock Like "adrs:Level2" Then
                                                Level2 = reader.Value
                                            ElseIf NameBlock Like "adrs:Level3" Then
                                                Level3 = reader.Value
                                            ElseIf NameBlock Like "adrOut4:Apartment" Then
                                                Apartament = reader.Value
                                            End If
                                        ElseIf nameN Like "Value" And FlagAddress = True Then
                                            If NameBlock Like "adrs:Level1" Then
                                                Level1 = Level1 & " " & reader.Value
                                            ElseIf NameBlock Like "adrs:Level2" Then
                                                Level2 = Level2 & " " & reader.Value
                                            ElseIf NameBlock Like "adrs:Level3" Then
                                                Level3 = Level3 & " " & reader.Value
                                            ElseIf NameBlock Like "adrs:Apartment" Then
                                                Apartament = Apartament & " " & reader.Value
                                            End If
                                        ElseIf nameN Like "Type" And FlagAddress = True Then
                                            If NameBlock Like "adrs:District" Then
                                                District1 = reader.Value & " " & District1
                                            ElseIf NameBlock Like "adrs:City" Then
                                                City = reader.Value & " " & City
                                            ElseIf NameBlock Like "adrs:UrbanDistrict" Then
                                                UrbanBuild = reader.Value & " " & UrbanBuild
                                            ElseIf NameBlock Like "adrs:SovietVillage" Then
                                                SovietVillige = reader.Value & " " & SovietVillige
                                            ElseIf NameBlock Like "adrs:Locality" Then
                                                Locality = reader.Value & " " & Locality
                                            ElseIf NameBlock Like "adrs:Street" Then
                                                Street = reader.Value & " " & Street
                                            ElseIf NameBlock Like "adrs:Level1" Then
                                                Level1 = reader.Value
                                            ElseIf NameBlock Like "adrs:Level2" Then
                                                Level2 = reader.Value
                                            ElseIf NameBlock Like "adrs:Level3" Then
                                                Level3 = reader.Value
                                            ElseIf NameBlock Like "adrs:Apartment" Then
                                                Apartament = reader.Value
                                            End If
                                        ElseIf nameN Like "Utilization" Then
                                            RazrIsp = FuncCadastr.FuncFindLandUtilizationByCod(reader.Value)
                                        ElseIf nameN Like "ByDoc" Then
                                            If NameBlock Like "Utilization" Then
                                                RazrIspByDoc = reader.Value
                                            End If
                                            'заполняем координаты
                                        ElseIf nameN Like "TypeUnit" And NameBlock Like "spa:SpelementUnit" Then
                                            NXYZ(0, countCoord) = reader.Value
                                        ElseIf nameN Like "TypeUnit" And NameBlock Like "SpelementUnit" Then
                                            NXYZ(0, countCoord) = reader.Value
                                        ElseIf nameN Like "SuNmb" And NameBlock Like "spa:SpelementUnit" Then
                                            NXYZ(1, countCoord) = reader.Value
                                        ElseIf nameN Like "X" And NameBlock Like "spa:Ordinate" Then
                                            NXYZ(2, countCoord) = reader.Value
                                        ElseIf nameN Like "X" And NameBlock Like "Ordinate" Then
                                            NXYZ(2, countCoord) = reader.Value
                                        ElseIf nameN Like "Y" And NameBlock Like "spa:Ordinate" Then
                                            NXYZ(3, countCoord) = reader.Value
                                        ElseIf nameN Like "Y" And NameBlock Like "Ordinate" Then
                                            NXYZ(3, countCoord) = reader.Value
                                        ElseIf nameN Like "OrdNmb" And NameBlock Like "spa:Ordinate" Then
                                            NXYZ(4, countCoord) = reader.Value
                                        ElseIf nameN Like "OrdNmb" And NameBlock Like "Ordinate" Then
                                            NXYZ(4, countCoord) = reader.Value
                                        ElseIf nameN Like "DeltaGeopoint" And NameBlock Like "spa:Ordinate" Then
                                            NXYZ(5, countCoord) = reader.Value
                                        ElseIf nameN Like "DeltaGeopoint" And NameBlock Like "Ordinate" Then
                                            NXYZ(5, countCoord) = reader.Value
                                        ElseIf nameN Like "GeopointZacrep" And NameBlock Like "spa:Ordinate" Then
                                            NXYZ(6, countCoord) = reader.Value
                                        ElseIf nameN Like "GeopointZacrep" And NameBlock Like "Ordinate" Then
                                            NXYZ(6, countCoord) = reader.Value
                                            'характеристики здания
                                        ElseIf nameN Like "Wall" Then
                                            StrMaterialWal = FuncCadastr.FuncCadastrMaterialWall(reader.Value)
                                        ElseIf nameN Like "YearBuilt" Then
                                            StrYearBuilt = reader.Value
                                        ElseIf nameN Like "YearUsed" Then
                                            StrYearUsed = reader.Value
                                        ElseIf nameN Like "Floors" Then
                                            StrCountFloor = reader.Value
                                        ElseIf nameN Like "UndergroundFloors" Then
                                            StrCountUnderFloor = reader.Value
                                        ElseIf NameBlock Like "KeyParameter" And nameN Like "Type" Then
                                            TypeKeyParameterConstr = Val(reader.Value)
                                        ElseIf NameBlock Like "KeyParameter" And nameN Like "Value" Then
                                            TypeValParameter = FuncCadastr.FuncTypeKeyParameter(TypeKeyParameterConstr, reader.Value)
                                        ElseIf nameN Like "Value" And NameBlock Like "CadastralCost" Then
                                            CadastralCostParcel = reader.Value
                                        ElseIf nameN Like "Unit" And NameBlock Like "CadastralCost" Then
                                            CadastralCostParcel = CadastralCostParcel & " " & FuncCadastr.FuncFindCadastrUnitByCode(reader.Value)
                                        ElseIf NameBlock Like "DeclarAttribute" And nameN Like "ExtractNumber" Then
                                            ExtractNumber = reader.Value
                                        ElseIf NameBlock Like "DeclarAttribute" And nameN Like "ExtractDate" Then
                                            ExtractData = reader.Value
                                        End If
                                    End While
                                End If
                                    '=====================================================================================================
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Case XmlNodeType.Text
                                If boolKv = True Then
                                    KVNumber = reader.Value
                                    boolKv = False
                                ElseIf bool_quarter_cad_number = True Then
                                    KVNumber = reader.Value
                                    bool_quarter_cad_number = False
                                ElseIf bool_registration_date = True Then
                                    registration_date = reader.Value
                                    bool_registration_date = False
                                ElseIf bool_cad_number = True And bool_commom_data Then
                                    NumberParcel = reader.Value
                                    bool_cad_number = False
                                ElseIf bool_code = True And bool_Type = True And bool_commom_data = True Then
                                    codeObject = reader.Value
                                    bool_code = False
                                ElseIf bool_value = True And bool_Type = True And bool_commom_data = True Then
                                    valueObject = reader.Value
                                    bool_value = False
                                ElseIf bool_value = True And bool_subtype = True And bool_commom_data = True Then
                                    TypeParcel = reader.Value
                                    bool_value = False
                                ElseIf boolParentNumb = True Then
                                    included_cad_number = reader.Value
                                    boolParentNumb = False
                                ElseIf bool_permitted_use = True And boolRazrIsp = True Then
                                    RazrIspByDoc = reader.Value
                                    boolRazrIsp = False
                                ElseIf boolCategory = True And bool_value = False Then
                                    Category = FuncCadastr.FuncFindLandCategoryByCode(reader.Value)
                                    boolCategory = False
                                ElseIf boolCategory = True And bool_value = True Then
                                    Category = reader.Value
                                    bool_value = False
                                ElseIf boolName = True Then
                                    boolName = False
                                    NameParcel = FuncCadastr.FuncCadastrStateParcel(reader.Value)
                                ElseIf boolObjectType = True Then
                                    StrObjectType = FuncCadastr.FuncFindTypeObjectByCod(reader.Value)
                                    boolObjectType = False
                                    'площадь
                                ElseIf boolAreaParcel = True And boolAreaUnit = False Then
                                    AreaParcel = reader.Value
                                ElseIf boolAreaParcel = True And bool_value = True Then
                                    AreaParcel = reader.Value
                                ElseIf boolInaccuracy = True Then
                                    AreaInaccuracy = reader.Value
                                    boolInaccuracy = False
                                ElseIf boolAreaParcel = True And bool_Type = False And boolInaccuracy = True Then
                                    AreaInaccuracy = reader.Value
                                    boolInaccuracy = False
                                ElseIf boolAreaParcel = True And boolAreaUnit = True Then
                                    AreaUnit = FuncCadastr.FuncFindCadastrUnitByCode(reader.Value)
                                    boolAreaUnit = False
                                    boolAreaParcel = False
                                ElseIf bool_AddresParcelNote = True Then
                                    AddresParcelNote = reader.Value
                                    bool_AddresParcelNote = False
                                ElseIf RightType = True And bool_value = True And FlagRight = True Then
                                    RightType = reader.Value
                                    bool_value = False
                                ElseIf RightType = True And boolName = True Then
                                    Patronymic = reader.Value
                                    boolName = False
                                ElseIf boolRightType = True Then
                                    If flagParcel = True Or flagBuilding = True Or flagConstruction = True Or flagUncompleted = True Then
                                        RightType = FuncCadastr.FuncFindRightByCod(reader.Value)
                                    End If
                                    boolRightType = False
                                ElseIf boolFM = True Then
                                    If flagParcel = True Or flagBuilding = True Or flagConstruction = True Or flagUncompleted = True Then
                                        FamilyName = reader.Value
                                        RightType = RightType & " " & FamilyName
                                    End If
                                    boolFM = False
                                ElseIf boolFirstNAme = True Then
                                    If flagParcel = True Or flagBuilding = True Or flagConstruction = True Or flagUncompleted = True Then
                                        FirstName = reader.Value
                                        RightType = RightType & " " & FirstName
                                    End If
                                    boolFirstNAme = False
                                ElseIf boolPatrony = True Then
                                    If flagParcel = True Or flagBuilding = True Or flagConstruction = True Or flagUncompleted = True Then
                                        Patronymic = reader.Value
                                        RightType = RightType & " " & Patronymic
                                    End If
                                    boolPatrony = False
                                ElseIf boolAssignBuild = True Then

                                    boolAssignBuild = False
                                ElseIf boolInnerKN = True And boolParentNumb = True Then
                                    included_cad_number = reader.Value
                                    boolParentNumb = False
                                ElseIf boolAssignNAme = True Then
                                    boolAssignNAme = False
                                    AssignConstr = reader.Value
                                ElseIf boolSpecialNote = True Then
                                    SpecialNote = reader.Value
                                    boolSpecialNote = False
                                ElseIf bool_num_geopoint = True Then
                                    NXYZ(1, countCoord) = reader.Value
                                    bool_num_geopoint = False
                                ElseIf bool_x = True Then
                                    NXYZ(2, countCoord) = reader.Value
                                    bool_x = False
                                ElseIf bool_y = True Then
                                    NXYZ(3, countCoord) = reader.Value
                                    bool_y = False
                                ElseIf bool_ord_nmb = True Then
                                    NXYZ(4, countCoord) = reader.Value
                                    bool_ord_nmb = False
                                ElseIf bool_delta_geopoint = True Then
                                    NXYZ(5, countCoord) = reader.Value
                                    bool_delta_geopoint = False
                                ElseIf bool_r = True Then
                                    NXYZ(7, countCoord) = reader.Value
                                    boolR = False
                                End If
                            Case XmlNodeType.EndElement
                                Dim NameBlock As String = reader.Name 'читаем ветку
                                If NameBlock Like "Right" Then
                                    FlagRight = False
                                ElseIf NameBlock Like "ObjectType" Then
                                    boolObjectType = False
                                ElseIf NameBlock Like "record_info" Then
                                    bool_record_info = False
                                ElseIf NameBlock Like "common_data" Then
                                    bool_commom_data = False
                                ElseIf NameBlock Like "type" Then
                                    bool_Type = False
                                ElseIf NameBlock Like "subtype" Then
                                    bool_subtype = False
                                ElseIf NameBlock Like "included_object" Then
                                    bool_included_object = False
                                ElseIf NameBlock Like "category" Then
                                    boolCategory = False
                                ElseIf NameBlock Like "permitted_use" Then
                                    bool_permitted_use = False
                                ElseIf NameBlock Like "Area" Then
                                    boolAreaParcel = False
                                ElseIf NameBlock Like "AssignationBuilding" Then
                                    boolAssignBuild = False
                                ElseIf NameBlock Like "Address" Then
                                    FlagAddress = False
                                    'формируем строку адреса
                                    If District1.Trim.Length > 0 Then AddresParcel = District1
                                    If City.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & City
                                    If UrbanBuild.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & UrbanBuild
                                    If SovietVillige.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & SovietVillige
                                    If Locality.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Locality
                                    If Street.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Street
                                    If Level1.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Level1
                                    If Level2.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Level2
                                    If Level3.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Level3
                                    If Apartament.Trim.Length > 0 Then AddresParcel = AddresParcel & " " & Apartament
                                    District1 = ""
                                    City = ""
                                    UrbanBuild = ""
                                    SovietVillige = ""
                                    Locality = ""
                                    Street = ""
                                    Level1 = ""
                                    Level2 = ""
                                    Level3 = ""
                                    Apartament = ""
                                    AddresParcelNote = ""
                                ElseIf NameBlock Like "Parcel" Then
                                    flagParcel = False
                                ElseIf NameBlock Like "SubParcel" Then
                                    flagSubParcel = False
                                ElseIf NameBlock Like "Building" Then
                                    flagBuilding = False
                                ElseIf NameBlock Like "SubBuilding" Then
                                    flagSubBuilding = False
                                ElseIf NameBlock Like "Construction" Then
                                    flagConstruction = False
                                ElseIf NameBlock Like "SubConstruction" Then
                                    flagSubConstruction = False
                                ElseIf NameBlock Like "Uncompleted" Then
                                    flagUncompleted = False
                                ElseIf NameBlock Like "InnerCadastralNumbers" Then
                                    boolInnerKN = False
                                    InnerKN = ""
                                ElseIf NameBlock Like "ParentCadastralNumbers" Then
                                    boolInnerKN = False
                                ElseIf NameBlock Like "spa:SpelementUnit" Then
                                    countCoord += 1
                                ElseIf NameBlock Like "SpelementUnit" Then
                                    countCoord += 1
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'конец блока с координатами, рисуем объект
                                ElseIf NameBlock Like "spa:SpatialElement" Or NameBlock Like "SpatialElement" Then
                                    'рисуем участок
                                    Dim PLineObj As ObjectId = Nothing
                                    If IsArray(NXYZ) = True Then
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'рисуем контур участка
                                        If IsArray(NXYZ) = True Then
                                            If NXYZ.GetUpperBound(1) > 0 Then
                                                'рисуем полилинию
                                                Dim IDPLine As ObjectId = FuncCadastr.FuncDrawPolyline(NXYZ)
                                            End If
                                            'добавляем расширенные данные

                                        End If
                                    End If
                                End If
                        End Select
                    End While
                Next
            Catch ex As System.Exception
                MsgBox(ex.ToString & vbLf & "Ошибка обработки данных. Макро прервано.")
                acTrans.Commit()
                Exit Sub
            End Try
            acTrans.Commit()
        End Using

    End Sub

    '===========================================================================================================
    'восстановить свойства из таблицы
    <CommandMethod("STPR-RestablODTlb")>
    Public Sub RestablODTlb()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim count As Integer = 0
        Dim TotalCount As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '========================================================================================
        'загружаем пути к сборкам
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = System.IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = System.IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                '========================================================================================
                '1.загружаем таблицы OD
                Dim counts As Integer = 0
                Dim PutchXMLFile As String = "k:\Civil3DToolsUtility\FileResources\MapClassification\"
                Dim boolPutchXMLFile As Boolean = False
                If System.IO.Directory.Exists(parentPatch2 & "\FileResources\MapClassification\") Then
                    PutchXMLFile = parentPatch2 & "\FileResources\MapClassification\"
                    boolPutchXMLFile = File.Exists(PutchXMLFile & "ClassObject.xml")
                    If boolPutchXMLFile = False Then
                        boolPutchXMLFile = File.Exists("k:\Civil3DToolsUtility\FileResources\MapClassification\ClassObject.xml")
                        If boolPutchXMLFile = True Then PutchXMLFile = "k:\Civil3DToolsUtility\FileResources\MapClassification\"
                    End If
                End If
                Dim NameFile As String = "ClassObject.xml"
                'вставляем PropertySet
                Dim idPSTblZU As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(PutchXMLFile & NameFile, "LAND_CADASTRAL_PARCEL")
                Dim idPSTblOstr As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(PutchXMLFile & NameFile, "CADASTRAL_OSTROV")
                '1. Делаем выбор либо таблица либо объекта
                Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pKeyOpts = New PromptKeywordOptions("")
                pKeyOpts.Message = vbLf & "Выбрать таблицу или участки? "
                pKeyOpts.Keywords.Add("Таблица")
                pKeyOpts.Keywords.Add("Участок")
                pKeyOpts.Keywords.Default = "Таблица"
                Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
                pKeyRes = ed.GetKeywords(pKeyOpts)
                If pKeyRes.Status = PromptStatus.OK Then
                    If pKeyRes.StringResult Like "Таблица" Then
                        '1. Выбираем таблицу с данными
                        Dim SelectTable As PromptEntityResult = ed.GetEntity("Выберите таблицу с данными:  ")
                        If SelectTable.Status = PromptStatus.Cancel Then
                            Throw New System.Exception("Прервано пользователем")
                            Exit Sub 'если нажали ESC то выходим из программы
                        Else
                            Dim acEnt1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectTable.ObjectId, OpenMode.ForRead)
                            If TypeOf acEnt1 Is Autodesk.AutoCAD.DatabaseServices.Table Then
                                Dim tlb As Autodesk.AutoCAD.DatabaseServices.Table = acEnt1
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                If tlb.Columns.Count > 19 Then
                                    Dim arrayData As String(,) = Nothing
                                    For i As Integer = 2 To tlb.Rows.Count - 1
                                        Dim NAmeTbl As String = ""
                                        Dim Str As String = tlb.Cells(i, 18).Value
                                        Dim Str1 As String = tlb.Cells(i, 1).Value
                                        If Str1 Like "Кадастровый участок" Or Str1 Like "Земельный участок" Then
                                            ReDim Preserve arrayData(2, 16)
                                            arrayData(0, 0) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 0) = "LABEL"
                                            arrayData(2, 0) = tlb.Cells(i, 1).Value

                                            arrayData(0, 1) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 1) = "Cad_number"
                                            arrayData(2, 1) = tlb.Cells(i, 2).Value

                                            arrayData(0, 2) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 2) = "Quarter_cad_numberr"
                                            Dim cad_numb As String = tlb.Cells(i, 2).Value
                                            arrayData(2, 2) = Mid(cad_numb, 1, cad_numb.LastIndexOf(":") - 1)

                                            arrayData(0, 3) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 3) = "Type"
                                            arrayData(2, 3) = tlb.Cells(i, 3).Value

                                            arrayData(0, 4) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 4) = "State"
                                            arrayData(2, 4) = tlb.Cells(i, 4).Value

                                            arrayData(0, 5) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 5) = "Area"
                                            arrayData(2, 5) = tlb.Cells(i, 8).Value

                                            arrayData(0, 6) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 6) = "OKATO"
                                            arrayData(2, 6) = tlb.Cells(i, 6).Value

                                            arrayData(0, 7) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 7) = "Adress"
                                            arrayData(2, 7) = tlb.Cells(i, 7).Value

                                            arrayData(0, 8) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 8) = "Category"
                                            arrayData(2, 8) = tlb.Cells(i, 11).Value

                                            arrayData(0, 9) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 9) = "Utilization"
                                            arrayData(2, 9) = tlb.Cells(i, 9).Value

                                            arrayData(0, 10) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 10) = "UtilizationByDoc"
                                            arrayData(2, 10) = tlb.Cells(i, 10).Value

                                            arrayData(0, 11) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 11) = "Patronymic"
                                            arrayData(2, 11) = tlb.Cells(i, 13).Value

                                            arrayData(0, 12) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 12) = "CadastralCost"
                                            arrayData(2, 12) = tlb.Cells(i, 12).Value

                                            arrayData(0, 13) = "LAND_CADASTRAL_PARCEL"
                                            arrayData(1, 13) = "NOTE"
                                            arrayData(2, 13) = tlb.Cells(i, 17).Value
                                            NAmeTbl = "LAND_CADASTRAL_PARCEL"

                                        ElseIf Str1 Like "Часть участка" Or Str1 Like "Часть" Then
                                            If arrayData.GetUpperBound(1) = 16 Then
                                                arrayData(2, 0) = tlb.Cells(i, 1).Value 'LABEL
                                                arrayData(2, 1) = tlb.Cells(i, 2).Value
                                                arrayData(2, 5) = tlb.Cells(i, 8).Value
                                            End If
                                            NAmeTbl = "LAND_CADASTRAL_PARCEL"
                                        End If


                                        If IsNothing(Str) = True Then Continue For
                                        If Str.Trim.Length > 0 Then
                                            Dim ArrayStr As String() = Str.Trim.Split
                                            If ArrayStr.Length > 0 Then
                                                Dim ArrayStrInfo As String() = Nothing
                                                ReDim ArrayStrInfo(15)
                                                For j As Integer = 0 To ArrayStr.Length - 1
                                                    Dim hgstr As String = ArrayStr(j)
                                                    Dim objid As ObjectId = AcadFunction.FuncReturnObjectID(hgstr)
                                                    If IsNothing(objid) = False And objid.IsNull = False Then
                                                        Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(objid, OpenMode.ForWrite)
                                                        If IsDBNull(obj) = False Then
                                                            'считываем тип объекта
                                                            If IsArray(arrayData) = True Then
                                                                If NAmeTbl Like "LAND_CADASTRAL_PARCEL" Then
                                                                    If IsNothing(idPSTblZU) = False Then
                                                                        If idPSTblZU.IsNull = False Then
                                                                            Dim boolAddPSTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(obj, idPSTblZU)
                                                                            If boolAddPSTbl = True Then
                                                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(obj, arrayData)
                                                                            End If
                                                                        End If
                                                                    End If

                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    count += 1
                                                    If count > 99 Then
                                                        TotalCount = TotalCount + count
                                                        MsgBox(TotalCount & "объектов обработано!!!")
                                                        count = 0
                                                    End If
                                                Next
                                            End If
                                        End If
                                        'проверка на внутренние островки
                                        Dim Str2 As String = tlb.Cells(i, 19).Value
                                        If IsNothing(Str2) = True Then Continue For
                                        If Str2.Trim.Length > 0 Then
                                            Dim ArrayStr As String() = Str2.Trim.Split
                                            If ArrayStr.Length > 0 Then
                                                For j As Integer = 0 To ArrayStr.Length - 1
                                                    Dim hgstr As String = ArrayStr(j)
                                                    Dim objid As ObjectId = AcadFunction.FuncReturnObjectID(hgstr)
                                                    If IsNothing(objid) = False And objid.IsNull = False Then
                                                        Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(objid, OpenMode.ForWrite)
                                                        Dim ArrayOstrov As String(,) = Nothing
                                                        If IsDBNull(obj) = False Then
                                                            ReDim Preserve ArrayOstrov(2, 2)
                                                            ArrayOstrov(0, 0) = "CADASTRAL_OSTROV"
                                                            ArrayOstrov(1, 0) = "LABEL"
                                                            ArrayOstrov(2, 0) = "Внутренний островок"

                                                            ArrayOstrov(0, 1) = "CADASTRAL_OSTROV"
                                                            ArrayOstrov(1, 1) = "ConditionalNumber"
                                                            ArrayOstrov(2, 1) = tlb.Cells(i, 2).Value

                                                            ArrayOstrov(0, 2) = "CADASTRAL_OSTROV"
                                                            ArrayOstrov(1, 2) = "IDKOUNTER"
                                                            ArrayOstrov(2, 2) = tlb.Cells(i, 18).Value
                                                        End If
                                                        'записываем OD
                                                        If idPSTblOstr.IsNull = False And IsArray(ArrayOstrov) = True Then
                                                            'добавляем таблицу к примитиву
                                                            Dim boolAddODtbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(obj, idPSTblOstr)
                                                            'записываем данные
                                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(obj, ArrayOstrov)
                                                        End If
                                                    End If
                                                Next j
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                End If
                acTrans.Commit()
            End Using
        End Using
    End Sub

End Class
