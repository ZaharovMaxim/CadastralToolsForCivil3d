Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports System.Environment
Imports System.Xml
Imports System.ComponentModel
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Buffer
Imports System.Windows.Forms

Public Class XRefHighlighter
    Public Shared Sub HighlightXRefEntity(xrefName As String, handleID As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                If bt.Has(xrefName) Then
                    Dim blockId As ObjectId = bt(xrefName)
                    Dim btr As BlockTableRecord = trans.GetObject(blockId, OpenMode.ForRead)

                    If btr.IsFromExternalReference Then
                        For Each id As ObjectId In btr
                            Dim acEnt As Entity = trans.GetObject(id, OpenMode.ForRead)
                            If acEnt.Id.IsNull = False AndAlso acEnt.Handle.ToString() = handleID Then
                                ' Временное выделение объекта
                                Dim objIds As ObjectId() = {id}
                                ed.SetImpliedSelection(objIds)
                                ed.WriteMessage(vbLf & "Объект выделен: " & acEnt.Handle.ToString())

                                ' Показ сообщения
                                ed.WriteMessage(vbLf & "Примитив во внешней ссылке подсвечен")
                                Exit For
                            End If
                        Next
                    End If
                End If

                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Ошибка выделения: " & ex.Message)
                trans.Abort()
            End Try
        End Using
    End Sub
End Class

Public Class TerritoryPlanProject
    Public PutchTemplateDwg As String = "d:\Civil3DToolsUtility\FileResources\Sample\" 'пути для файла шаблонов
    Public putshXMLData As String = ""
    Public Shared custdata As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
    Private WithEvents _editor As Autodesk.AutoCAD.EditorInput.Editor

    Private foundNodes As New List(Of TreeNode)()
    Private currentIndex As Integer = -1

    Public Enum SearchFromEndMode
        EndsWith        ' Точное совпадение с конца
        ContainsFromEnd ' Содержит текст ближе к концу
        LastIndexOf     ' Последнее вхождение
    End Enum

    Public Class LandCategory
        Public Property Name As String
        Public Property Code As String

        Public Overrides Function ToString() As String
            Return $"{Name} ({Code})"
        End Function
    End Class
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'столбцы
    Public Class Customer
        Implements System.ComponentModel.INotifyPropertyChanged
        Public Event PropertyChanged As System.ComponentModel.PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged
        Public Property Имя_свойства As String
        Public Property Значение As String
        'Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    End Class
    Structure CentrePolygon
        Dim x As Double
        Dim y As Double
    End Structure
    'возвращает корневой элемент
    Public Function FindRootTreeViewItem(currentItem As TreeViewItem) As TreeViewItem
        If currentItem Is Nothing Then Return Nothing
        Dim parent As DependencyObject = currentItem.Parent
        ' Поднимаемся вверх по иерархии пока не найдем корень
        While parent IsNot Nothing AndAlso TypeOf parent IsNot Controls.TreeView
            If TypeOf parent Is TreeViewItem Then
                currentItem = DirectCast(parent, TreeViewItem)
            End If
            parent = VisualTreeHelper.GetParent(parent)
        End While
        Return currentItem
    End Function
    'функция возвращает объект из внешнего чертежа по handle
    '===================================================================================================================
    Public Shared Function FuncReturnEntityDeskDrawByHandle(ByVal putchDWG As String, ByVal handleID As String) As Autodesk.AutoCAD.DatabaseServices.Entity
        FuncReturnEntityDeskDrawByHandle = Nothing
        If File.Exists(putchDWG) = True Then
            Try
                Dim sourceDb As Database = New Database(False, True)
                sourceDb.ReadDwgFile(putchDWG, FileShare.Read, True, String.Empty)
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                    'получаем доступ к объекту
                    Dim ln As Long = Convert.ToInt64(handleID, 16)
                    Dim hn As Handle = New Handle(ln)
                    Dim id As ObjectId = sourceDb.GetObjectId(False, hn, 0)
                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(id, OpenMode.ForRead)
                    If acEnt.Id.IsNull = False Then
                        Return acEnt
                    End If
                End Using
            Catch ex As System.Exception
            End Try
        End If
    End Function
    '===================================================================================================================
    'функция читает все подключенные ссылки к файлу dwg (первый элемент в списке - это исходный чертеж к которому подключены ссылки)
    Public Shared Function FuncReadXRef(ByRef arrayXRef As String(,)) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        If IsArray(arrayXRef) = True Then
            Erase arrayXRef
        End If
        Dim countArrayXRef As Integer = 0
        Dim putchDwg As String = doc.Name
        ReDim Preserve arrayXRef(2, countArrayXRef)
        arrayXRef(0, countArrayXRef) = Path.GetFileNameWithoutExtension(putchDwg)
        arrayXRef(1, countArrayXRef) = putchDwg
        arrayXRef(2, countArrayXRef) = ""
        countArrayXRef += 1
        Dim directory As String = Path.GetDirectoryName(putchDwg)
        Using docLock As DocumentLock = doc.LockDocument()
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Получаем таблицу блоков
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    ' Перебираем все записи в таблице блоков
                    For Each blockId As ObjectId In bt
                        Dim btr As BlockTableRecord = trans.GetObject(blockId, OpenMode.ForRead)
                        If btr.IsFromExternalReference Then
                            putchDwg = btr.PathName
                            If putchDwg.StartsWith("..\") OrElse putchDwg.StartsWith(".\") Then
                                putchDwg = directory & "\" & btr.Name & ".dwg"
                                If File.Exists(putchDwg) = False Then Continue For
                            End If
                            Dim boolRecDWG As Boolean = False
                            If IsArray(arrayXRef) = True Then
                                If arrayXRef.GetUpperBound(1) > 1 Then
                                    For i As Integer = 0 To arrayXRef.GetUpperBound(1)
                                        Dim temoPatch As String = arrayXRef(1, i)
                                        If temoPatch Like putchDwg Then
                                            boolRecDWG = True
                                            Exit For
                                        End If
                                    Next i
                                End If
                            End If
                            If boolRecDWG = False Then
                                ' Проверяем, является ли блок внешней ссылкой
                                ReDim Preserve arrayXRef(2, countArrayXRef)
                                arrayXRef(0, countArrayXRef) = btr.Name
                                arrayXRef(1, countArrayXRef) = putchDwg
                                Dim xrefInfo As XrefGraph = db.GetHostDwgXrefGraph(True)
                                Dim countNodes As Integer = xrefInfo.NumNodes
                                For i As Integer = 0 To countNodes - 1
                                    Dim node As XrefGraphNode = xrefInfo.GetXrefNode(i)
                                    If node.Name Like btr.Name Then
                                        arrayXRef(2, countArrayXRef) = node.XrefStatus.ToString
                                        Exit For
                                    End If
                                    'If node.XrefStatus =
                                    'If Not node.XrefStatus.IsResolved Then
                                    'Continue For
                                    'End If
                                Next
                                countArrayXRef += 1
                            End If
                        End If
                    Next
                    trans.Commit()
                Catch ex As Exception
                    trans.Abort()
                    Return False
                End Try
            End Using
        End Using
        Return True
    End Function
    '===================================================================================================================
    'функция читает характеристики объектов из внешних ссылок
    Public Shared Function FuncReadObjectDestDraw(ByVal putchDWG As String, ByRef dictionaryALLObject As Dictionary(Of String, Dictionary(Of String, String(,)))) As Boolean
        If File.Exists(putchDWG) = True Then
            Try
                Dim sourceDb As Database = New Database(False, True)
                sourceDb.ReadDwgFile(putchDWG, FileShare.Read, True, String.Empty)
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                    For Each idObject As ObjectId In sourceBtr
                        If idObject.IsNull = False Then
                            Dim entObject As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(idObject, OpenMode.ForRead) 'открываем объект на чтение
                            Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(entObject)
                            If setIds.Count > 0 Then
                                For Each psId As ObjectId In setIds
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = sourceTr.GetObject(psId, OpenMode.ForRead, False, False)
                                    Dim nameTable As String = pset.PropertySetDefinitionName
                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    Dim arrayData As String(,) = {}
                                    Dim countArrayData As Integer = 0
                                    For j As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(j)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            ReDim Preserve arrayData(1, countArrayData)
                                            arrayData(0, countArrayData) = dataField
                                            arrayData(1, countArrayData) = data
                                            countArrayData += 1
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next j
                                    If arrayData.Length > 0 Then
                                        If dictionaryALLObject.Count > 0 Then
                                            If dictionaryALLObject.ContainsKey(nameTable) = True Then
                                                Dim newDict As Dictionary(Of String, String(,)) = dictionaryALLObject.Item(nameTable)
                                                If newDict.ContainsKey(entObject.Handle.ToString) = False Then
                                                    newDict.Add(entObject.Handle.ToString, arrayData)
                                                    dictionaryALLObject.Item(nameTable) = newDict
                                                End If
                                            Else
                                                Dim newDict As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                newDict.Add(entObject.Handle.ToString, arrayData)
                                                dictionaryALLObject.Add(nameTable, newDict)
                                            End If
                                        Else
                                            Dim newDict As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                            newDict.Add(entObject.Handle.ToString, arrayData)
                                            dictionaryALLObject.Add(nameTable, newDict)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                    sourceTr.Commit()
                End Using
                sourceDb.CloseInput(True)
                FuncReadObjectDestDraw = True
            Catch ex As System.Exception
            End Try
        End If
    End Function
    '===================================================================================================================
    'функция записывает характеристики объектов из внешних ссылок (не рабочая функция, нельзя менять XRef)
    Public Shared Function FuncWriteObjectDestDraw(ByVal xrefName As String, ByVal handleID As String, ByVal dictionaryPropertyPS As Dictionary(Of String, String(,))) As Boolean
        FuncWriteObjectDestDraw = False
        Dim nameTablePS As String = ""
        Dim arrayDataRec As String(,) = {}
        If dictionaryPropertyPS.Count > 0 Then
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                ' Получаем таблицу блоков
                Try
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    If bt.Has(xrefName) Then
                        Dim blockId As ObjectId = bt(xrefName)
                        Dim btr As BlockTableRecord = trans.GetObject(blockId, OpenMode.ForRead)
                        If btr.IsFromExternalReference Then
                            For Each id As ObjectId In btr
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = trans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                                If acEnt.Id.IsNull = False Then
                                    If acEnt.Handle.ToString Like handleID Then
                                        If acEnt.IsWriteEnabled = False Then
                                            acEnt.UpgradeOpen()
                                        End If
                                        Dim boolRec As Boolean = False
                                        Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(acEnt)
                                        If setIds.Count > 0 Then
                                            For Each psId As ObjectId In setIds
                                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = trans.GetObject(psId, OpenMode.ForWrite, False, False)
                                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                                Dim nameTable As String = pset.PropertySetDefinitionName
                                                For i As Integer = 0 To dictionaryPropertyPS.Count - 1
                                                    nameTablePS = dictionaryPropertyPS.ElementAt(i).Key
                                                    arrayDataRec = dictionaryPropertyPS.ElementAt(i).Value
                                                    If nameTable Like nameTablePS Then
                                                        If psetDataColl.Count > 0 Then
                                                            For j As Integer = 0 To psetDataColl.Count - 1
                                                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = psetDataColl.Item(j)
                                                                Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                                Dim data As String = pTempData.GetData() 'значение
                                                                For k As Integer = 0 To arrayDataRec.GetUpperBound(1)
                                                                    Dim nameField As String = arrayDataRec(0, k)
                                                                    Dim value As String = arrayDataRec(1, k)
                                                                    If dataField Like nameField Then
                                                                        If pTempData.DataType = Autodesk.AEC.PropertyData.DataType.Real Then
                                                                            Dim RealData As Double = 0
                                                                            If IsNothing(value) = True Then
                                                                                RealData = 0
                                                                            ElseIf value.Trim.Length = 0 Then
                                                                                RealData = 0
                                                                            Else
                                                                                If IsNumeric(value) = True Then
                                                                                    RealData = Val(value)
                                                                                Else
                                                                                    RealData = 0
                                                                                End If
                                                                            End If
                                                                            pTempData.SetData(RealData)
                                                                        Else
                                                                            If IsNothing(value) = True Then
                                                                                value = ""
                                                                            End If
                                                                            pTempData.SetData(value)
                                                                        End If

                                                                        Exit For
                                                                    End If
                                                                Next k
                                                            Next j
                                                        End If
                                                        boolRec = True
                                                    End If
                                                Next i
                                            Next
                                        End If
                                        If boolRec = True Then
                                            trans.Commit()
                                            'Dim putchDwg As String = btr.PathName
                                            'Dim directory As String = Path.GetDirectoryName(btr.Database.Filename)
                                            'If putchDwg.StartsWith("..\") OrElse putchDwg.StartsWith(".\") Then
                                            '    putchDwg = Directory & "\" & btr.Name & ".dwg"
                                            '    If File.Exists(putchDwg) = False Then Continue For
                                            'End If
                                            'btr.Database.SaveAs(putchDwg, DwgVersion.Current)
                                            Return True
                                        End If
                                    End If
                                End If

                            Next
                        End If
                    End If
                    FuncWriteObjectDestDraw = True
                Catch ex As System.Exception
                End Try
            End Using
        End If
    End Function

    Public Shared Function FuncWriteObjectDestDraw2(ByVal xrefName As String, ByVal handleID As String, ByVal dictionaryPropertyPS As Dictionary(Of String, String(,))) As Boolean
        FuncWriteObjectDestDraw2 = False
        Dim nameTablePS As String = ""
        Dim arrayDataRec As String(,) = {}
        If IsNothing(dictionaryPropertyPS) = True Then Return False
        If dictionaryPropertyPS.Count = 0 Then Return False
        Try
            If dictionaryPropertyPS.Count > 0 Then
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim database As Database = doc.Database
                Dim mXRefTree As XrefGraph = database.GetHostDwgXrefGraph(False)

                Dim xrefDatatbase As Database = Nothing
                Dim xgn As XrefGraphNode = Nothing

                For i As Integer = 0 To mXRefTree.HostDrawing.NumOut - 1
                    xgn = TryCast(mXRefTree.HostDrawing.Out(i), XrefGraphNode)
                    ' Находим имя файла внешней ссылки для редактирования
                    If String.Compare(xgn.Name, xrefName, True) = 0 Then
                        xrefDatatbase = xgn.Database
                        Exit For
                    End If
                Next
                If xrefDatatbase Is Nothing Then
                    Return False
                End If

                Using Filelock As XrefFileLock = XrefFileLock.LockFile(xrefDatatbase.XrefBlockId)
                    xrefDatatbase.RestoreOriginalXrefSymbols()
                    Using tr As Transaction = xrefDatatbase.TransactionManager.StartTransaction()
                        Dim bt As BlockTable = CType(tr.GetObject(xrefDatatbase.BlockTableId, OpenMode.ForRead), BlockTable)
                        Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                        For Each id As ObjectId In ms
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForWrite) 'открываем объект на чтение
                            If acEnt.Id.IsNull = False Then
                                If acEnt.Handle.ToString Like handleID Then
                                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(acEnt)
                                    If setIds.Count > 0 Then
                                        For Each psId As ObjectId In setIds
                                            Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = tr.GetObject(psId, OpenMode.ForWrite, False, False)
                                            Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                            Dim nameTable As String = pset.PropertySetDefinitionName
                                            For i As Integer = 0 To dictionaryPropertyPS.Count - 1
                                                nameTablePS = dictionaryPropertyPS.ElementAt(i).Key
                                                arrayDataRec = dictionaryPropertyPS.ElementAt(i).Value
                                                If IsNothing(arrayDataRec) = True Then Return False
                                                If nameTable Like nameTablePS Then
                                                    If psetDataColl.Count > 0 Then
                                                        For j As Integer = 0 To psetDataColl.Count - 1
                                                            Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = psetDataColl.Item(j)
                                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                            Dim data As String = pTempData.GetData() 'значение
                                                            For k As Integer = 0 To arrayDataRec.GetUpperBound(1)
                                                                Dim nameField As String = arrayDataRec(0, k)
                                                                Dim value As String = arrayDataRec(1, k)
                                                                If dataField Like nameField Then
                                                                    If pTempData.DataType = Autodesk.AEC.PropertyData.DataType.Real Then
                                                                        Dim RealData As Double = 0
                                                                        If IsNothing(value) = True Then
                                                                            RealData = 0
                                                                        ElseIf value.Trim.Length = 0 Then
                                                                            RealData = 0
                                                                        Else
                                                                            If IsNumeric(value) = True Then
                                                                                RealData = Val(value)
                                                                            Else
                                                                                RealData = 0
                                                                            End If
                                                                        End If
                                                                        pTempData.SetData(RealData)
                                                                    Else
                                                                        If IsNothing(value) = True Then
                                                                            value = ""
                                                                        End If
                                                                        pTempData.SetData(value)
                                                                    End If
                                                                    Exit For
                                                                End If
                                                            Next k
                                                        Next j
                                                    End If
                                                End If
                                            Next i
                                        Next
                                    End If
                                End If
                            End If
                        Next
                        tr.Commit()
                    End Using
                    xrefDatatbase.RestoreForwardingXrefSymbols()
                End Using
            End If
        Catch ex As System.Exception
        End Try
        Return True
    End Function

    '===================================================================================================================
    'функция читает характеристики объектов из активного чертежа
    Public Shared Function FuncReadObjectCurrentDraw(ByRef dictionaryALLObject As Dictionary(Of String, Dictionary(Of String, String(,)))) As Boolean
        FuncReadObjectCurrentDraw = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim count As Integer = 0
        Dim count1 As Integer = 0
        Using docLock As DocumentLock = acDoc.LockDocument()
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each idObject As ObjectId In btr
                    Try
                        If idObject.IsNull = False Then
                            Dim entObject As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(idObject, OpenMode.ForRead) 'открываем объект на чтение
                            Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(entObject)
                            If setIds.Count > 0 Then
                                For Each psId As ObjectId In setIds
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                    Dim nameTable As String = pset.PropertySetDefinitionName
                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    Dim arrayData As String(,) = {}
                                    Dim countArrayData As Integer = 0
                                    For j As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(j)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            ReDim Preserve arrayData(1, countArrayData)
                                            arrayData(0, countArrayData) = dataField
                                            arrayData(1, countArrayData) = data
                                            countArrayData += 1
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next j
                                    If arrayData.Length > 0 Then
                                        If dictionaryALLObject.Count > 0 Then
                                            If dictionaryALLObject.ContainsKey(nameTable) = True Then
                                                Dim newDict As Dictionary(Of String, String(,)) = dictionaryALLObject.Item(nameTable)
                                                If newDict.ContainsKey(entObject.Handle.ToString) = False Then
                                                    newDict.Add(entObject.Handle.ToString, arrayData)
                                                    dictionaryALLObject.Item(nameTable) = newDict
                                                End If
                                            Else
                                                Dim newDict As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                                newDict.Add(entObject.Handle.ToString, arrayData)
                                                dictionaryALLObject.Add(nameTable, newDict)
                                            End If
                                        Else
                                            Dim newDict As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                            newDict.Add(entObject.Handle.ToString, arrayData)
                                            dictionaryALLObject.Add(nameTable, newDict)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Catch ex As System.Exception
                        acTrans.Abort()
                    End Try
                Next
                acTrans.Commit()
            End Using
        End Using
        Return True
    End Function

    Private Function IsObjectInXRef(objId As ObjectId, tr As Transaction) As Boolean
        Dim blkRef As BlockReference = TryCast(tr.GetObject(objId, OpenMode.ForRead), BlockReference)

        If blkRef IsNot Nothing Then
            Dim btr As BlockTableRecord = tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead)

            ' Проверяем, является ли блок внешней ссылкой
            Return btr.IsFromExternalReference
        End If

        Return False
    End Function

    Public Shared Function SelectXRefEntityByPoint(xrefName As String) As ObjectId
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Запрос точки выбора у пользователя
        Dim ppo As New PromptPointOptions(vbLf & "Выберите примитив во внешней ссылке: ")
        Dim ppr As PromptPointResult = ed.GetPoint(ppo)

        If ppr.Status <> PromptStatus.OK Then
            Return ObjectId.Null
        End If

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Используем выбор по точке
                'Dim psr As PromptSelectionResult = ed.SelectAtPoint(ppr.Value)

                'If psr.Status = PromptStatus.OK Then
                '    For Each selectedId As ObjectId In psr.Value.GetObjectIds()
                '        Dim ent As Entity = trans.GetObject(selectedId, OpenMode.ForRead)

                '        ' Проверяем, принадлежит ли объект нужной внешней ссылке
                '        If IsEntityInXRef(ent, xrefName, trans) Then
                '            trans.Commit()
                '            ed.WriteMessage(vbLf & "Выбран примитив: " & ent.Handle.ToString())
                '            Return selectedId
                '        End If
                '    Next
                'End If

                trans.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbLf & "Ошибка выбора: " & ex.Message)
                trans.Abort()
            End Try
        End Using

        Return ObjectId.Null
    End Function

    Private Function IsEntityInXRef(ent As Entity, xrefName As String, trans As Transaction) As Boolean
        Try
            Dim ownerId As ObjectId = ent.OwnerId
            Dim owner As BlockTableRecord = trans.GetObject(ownerId, OpenMode.ForRead)

            Return owner.IsFromExternalReference AndAlso owner.Name = xrefName
        Catch
            Return False
        End Try
    End Function


    Public Function ParseLandCategories(xmlText As String) As List(Of LandCategory)
        Dim categories As New List(Of LandCategory)()

        Try
            Dim wrappedXml As String = "<root>" & xmlText & "</root>"
            Dim xDoc As XDocument = XDocument.Parse(wrappedXml)

            For Each element As XElement In xDoc.Root.Elements("Name")
                Dim category As New LandCategory() With {
                .Name = element.Attribute("Value")?.Value,
                .Code = element.Attribute("Code")?.Value
            }

                If category.Name IsNot Nothing AndAlso category.Code IsNot Nothing Then
                    categories.Add(category)
                End If
            Next

        Catch ex As Exception
            Console.WriteLine($"Ошибка при парсинге: {ex.Message}")
        End Try

        Return categories
    End Function


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'запуск палитры
    Public Sub New()
        InitializeComponent()
        Dim arrayXRef As String(,) = {}
        'ветка чертежей
        Dim node0 As TreeViewItem = New TreeViewItem
        node0.Header = "Управление чертежами"
        node0.Name = "FileNameDWG"
        node0.FontSize = 16
        node0.Tag = {"FileNameDWG"}

        'ветка буферные зоны
        Dim node1 As TreeViewItem = New TreeViewItem
        node1.Header = "Буферные зоны"
        node1.Name = "BufferZone"
        node1.FontSize = 16
        node1.Tag = {"BufferZone"}

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
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
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        'Путь к файлам шаблонам с условными знами
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplateDwg = parentPatch2 & "\FileResources\Sample\"
        Else
            MsgBox("Каталог \\FileResources\Sample, не найден!!!")
            Exit Sub
        End If
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\Semantic\") Then
            putshXMLData = parentPatch2 & "\FileResources\Sample\Semantic\"
        End If
        '4.читаем шаблоны условных знаков
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. получем доступ к пользовательскому каталогу
        'временный каталог для автосохранения таблицы
        Dim tempDirTopoTemplate As String = ""
        Dim tempDirSemantic As String = ""
        Dim systemTemporaryDirectory = GetFolderPath(SpecialFolder.ApplicationData)
        Try
            If IO.Directory.Exists(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates") = False Then
                IO.Directory.CreateDirectory(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates")
            End If
            If IO.Directory.Exists(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates\Semantic") = False Then
                IO.Directory.CreateDirectory(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates\Semantic")
            End If
            If IO.Directory.Exists(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                tempDirTopoTemplate = systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates"
            End If
            If IO.Directory.Exists(systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates\Semantic") = True Then
                tempDirSemantic = systemTemporaryDirectory & "\Civil3DToolsUtility\TOPOTemplates\Semantic"
            End If
        Catch ex As System.Security.SecurityException
        End Try
        '=============================================================================================
        'читаем файлы xml
        If IO.Directory.Exists(PutchTemplateDwg) = True Then
            Dim arrayStr As String() = {"FileNameTOPOxml", tempDirTopoTemplate}
            node0.Tag = arrayStr
            Try
                'создаем дерево в TreeView
                For Each userfile As String In Directory.GetFiles(PutchTemplateDwg, "*.xml").Concat(Directory.GetFiles(PutchTemplateDwg, "*.dwg"))
                    Dim nameOriginFile As String = IO.Path.GetFileName(userfile)
                    Dim fileDate1 As Date = File.GetLastWriteTime(userfile)
                    Dim fileDate2 As Date = New Date()
                    If File.Exists(tempDirTopoTemplate & "\" & nameOriginFile) Then
                        fileDate2 = File.GetLastWriteTime(tempDirTopoTemplate & "\" & nameOriginFile)
                    End If
                    Dim destCopyFileXML As String = tempDirTopoTemplate & "\" & nameOriginFile
                    If fileDate1 > fileDate2 Then
                        IO.File.Copy(userfile, destCopyFileXML, True)
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If

        '=============================================================================================
        'читаем файлы xml
        If IO.Directory.Exists(PutchTemplateDwg & "Semantic") = True Then
            Dim arrayStr As String() = {"BufferZone", tempDirSemantic}
            node1.Tag = arrayStr
            Try
                'создаем дерево в TreeView
                For Each fileXML As String In Directory.GetFiles(PutchTemplateDwg & "Semantic", "*.xml")
                    Dim nameOriginFile As String = IO.Path.GetFileName(fileXML)
                    Dim fileDate1 As Date = File.GetLastWriteTime(fileXML)
                    Dim fileDate2 As Date = New Date()
                    If File.Exists(tempDirSemantic & "\" & nameOriginFile) Then
                        fileDate2 = File.GetLastWriteTime(tempDirSemantic & "\" & nameOriginFile)
                    End If
                    Dim destCopyFileXML As String = tempDirSemantic & "\" & nameOriginFile
                    If fileDate1 > fileDate2 Then
                        IO.File.Copy(fileXML, destCopyFileXML, True)
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        TreeViewDWG.Items.Add(node0)
        TreeViewDWG.Items.Add(node1)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменение размера палитры
    Private Sub TerritoryPlanProject_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Try
            Dim s1 As Double = Me.Height 'ширина формы (палитры)
            Dim s2 As Double = Me.Width
            If s1 > 620 Then
                GridDWG.Height = s1 - 10
                TreeViewDWG.Height = s1 - 300
            End If
            If s2 > 350 Then
                GridDWG.Width = s2 - 10
                TreeViewDWG.Width = s2 - 20
                PropertyGridObject.Width = s2 - 20
            Else
                GridDWG.Width = 340
                TreeViewDWG.Width = 320
                PropertyGridObject.Width = 320
            End If
        Catch ex As Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор узла дерева
    Private Sub TreeViewDWG_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles TreeViewDWG.SelectedItemChanged
        Dim TreeViewControl As System.Windows.Controls.TreeView = e.Source
        TreeViewControl.Focus()
        ButtonWritePS.Background = New SolidColorBrush(Color.FromRgb(221, 221, 221))
        Dim itemSelectObject As Object = TreeViewControl.SelectedValue
        If TypeOf itemSelectObject Is DockPanel Then

        Else
            Dim ItemSelect As TreeViewItem = itemSelectObject 'ссылка на выбранный элемент
            If IsNothing(ItemSelect) = False Then
                Dim typeItem As String = ItemSelect.Name 'тип выбранной ветки
                If typeItem Like "FileNameDWG" Then
                    Dim arrayXRef As String(,) = {}
                    Dim boolReadXRef As Boolean = FuncReadXRef(arrayXRef)
                    If arrayXRef.Length > 0 Then
                        For i As Integer = 0 To arrayXRef.GetUpperBound(1)
                            Dim nameDrawing As String = arrayXRef(0, i)
                            Dim putchDrawing As String = arrayXRef(1, i)
                            Dim statusXRef As String = arrayXRef(2, i)
                            Dim Node0 As TreeViewItem = New TreeViewItem
                            Node0.Header = nameDrawing
                            Node0.Tag = {"xRefDwg", nameDrawing, putchDrawing, statusXRef}
                            Node0.Name = "xRefDwg"
                            Node0.FontSize = 14
                            ItemSelect.Items.Add(Node0)
                        Next i
                    End If
                ElseIf typeItem Like "BufferZone" Then
                    Dim arrayTagItemSelect As String() = ItemSelect.Tag 'считываем скрытые данные
                    Dim putchXML As String = arrayTagItemSelect(1)
                    If Directory.Exists(putchXML) = True Then
                        For Each fileXML As String In Directory.GetFiles(putchXML, "*.xml")
                            Dim reader As XmlReader = XmlReader.Create(fileXML)
                            Try
                                While reader.Read()
                                    If reader.NodeType = XmlNodeType.Element Then
                                        If reader.Name Like "TypeBuffer" Then
                                            If reader.HasAttributes = True Then
                                                While reader.MoveToNextAttribute()
                                                    Dim name As String = reader.Name
                                                    Dim value As String = reader.Value
                                                    If name Like "Desk" Then
                                                        If value.Trim.Length > 0 Then
                                                            Dim Node0 As TreeViewItem = New TreeViewItem
                                                            Node0.Header = value
                                                            Dim arrayRec As String() = {"TypeBuffer", name, value, fileXML}
                                                            Node0.Tag = arrayRec
                                                            Node0.Name = "TypeBuffer"
                                                            Node0.FontSize = 14
                                                            ItemSelect.Items.Add(Node0)
                                                        End If
                                                    End If
                                                End While
                                            End If
                                        End If
                                    End If
                                End While
                            Catch ex As System.Xml.XmlException
                            End Try
                            reader.Close()
                        Next
                    End If
                ElseIf typeItem Like "TypeBuffer" Then
                    Dim arrayTagItemSelect As String() = ItemSelect.Tag 'считываем скрытые данные
                    Dim nameType As String = arrayTagItemSelect(2)
                    Dim putchXML As String = arrayTagItemSelect(3)
                    If File.Exists(putchXML) = True Then
                        Dim reader As XmlReader = XmlReader.Create(putchXML)
                        Dim boolReadFile As Boolean = False
                        While reader.Read()
                            Try
                                If reader.NodeType = XmlNodeType.Element Then
                                    If reader.Name Like "TypeBuffer" Then
                                        If reader.HasAttributes = True Then
                                            While reader.MoveToNextAttribute()
                                                Dim name As String = reader.Name
                                                Dim value As String = reader.Value
                                                If name Like "Desk" Then
                                                    If value.Trim.Length > 0 Then
                                                        If value Like nameType Then
                                                            boolReadFile = True
                                                        Else
                                                            boolReadFile = False
                                                        End If
                                                    End If
                                                End If
                                            End While
                                        End If
                                    Else
                                        If boolReadFile = True Then
                                            If reader.Name Like "NamePipeLine" Then
                                                If reader.HasAttributes = True Then
                                                    Dim arrayRec As String() = {"NamePipeLine", "", "", putchXML}
                                                    Dim countArrayRec As Integer = 1
                                                    While reader.MoveToNextAttribute()
                                                        Dim name As String = reader.Name
                                                        Dim value As String = reader.Value
                                                        If name Like "Desk" Then
                                                            arrayRec(1) = reader.Value
                                                        ElseIf name Like "Document" Then
                                                            arrayRec(2) = reader.Value
                                                        End If
                                                    End While
                                                    If arrayRec(1).Trim.Length > 0 Then
                                                        Dim Node0 As TreeViewItem = New TreeViewItem
                                                        Node0.Header = arrayRec(1)
                                                        Node0.Tag = arrayRec
                                                        Node0.Name = arrayRec(0)
                                                        Node0.FontSize = 12
                                                        ItemSelect.Items.Add(Node0)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Catch ex As System.Xml.XmlException
                            End Try
                        End While
                        reader.Close()
                    End If
                ElseIf typeItem Like "NamePipeLine" Then
                    Dim arrayTagItemSelect As String() = ItemSelect.Tag 'считываем скрытые данные
                    Dim nameType As String = arrayTagItemSelect(1)
                    Dim putchXML As String = arrayTagItemSelect(3)
                    If File.Exists(putchXML) = True Then
                        Dim reader As XmlReader = XmlReader.Create(putchXML)
                        Dim boolReadFile As Boolean = False
                        While reader.Read()
                            Try
                                If reader.NodeType = XmlNodeType.Element Then
                                    If reader.Name Like "NamePipeLine" Then
                                        If reader.HasAttributes = True Then
                                            While reader.MoveToNextAttribute()
                                                Dim name As String = reader.Name
                                                Dim value As String = reader.Value
                                                If name Like "Desk" Then
                                                    If value.Trim.Length > 0 Then
                                                        If value Like nameType Then
                                                            boolReadFile = True
                                                        Else
                                                            boolReadFile = False
                                                        End If
                                                    End If
                                                End If
                                            End While
                                        End If
                                    Else
                                        If boolReadFile = True Then
                                            If reader.Name Like "PipeLine" Then
                                                If reader.HasAttributes = True Then
                                                    Dim arrayRec As String() = {"PipeLine", "", "", putchXML}
                                                    Dim countArrayRec As Integer = 1
                                                    While reader.MoveToNextAttribute()
                                                        Dim name As String = reader.Name
                                                        Dim value As String = reader.Value
                                                        If name Like "Desk" Then
                                                            arrayRec(1) = reader.Value
                                                        ElseIf name Like "Offset" Then
                                                            arrayRec(2) = reader.Value
                                                        End If
                                                    End While
                                                    If arrayRec(1).Trim.Length > 0 Then
                                                        Dim Node0 As TreeViewItem = New TreeViewItem
                                                        Node0.Header = arrayRec(1)
                                                        Node0.Tag = arrayRec
                                                        Node0.Name = arrayRec(0)
                                                        Node0.FontSize = 12
                                                        ItemSelect.Items.Add(Node0)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Catch ex As System.Xml.XmlException
                            End Try
                        End While
                        reader.Close()
                    End If
                ElseIf typeItem Like "PipeLine" Then
                    Dim arrayTagItemSelect As String() = ItemSelect.Tag 'считываем скрытые данные
                    Dim offst As Double = Val(arrayTagItemSelect(2))
                    custdata.Clear()
                    Dim Acustim As Customer = New Customer()
                    Acustim.Имя_свойства = "Размер зоны"
                    Acustim.Значение = offst
                    custdata.Add(Acustim)
                    PropertyGridObject.HeadersVisibility = False
                    PropertyGridObject.Tag = arrayTagItemSelect
                    PropertyGridObject.ItemsSource = custdata
                ElseIf typeItem Like "xRefDwg" Then
                    Dim arrayTagItemSelect As String() = ItemSelect.Tag 'считываем скрытые данные
                    If ItemSelect.Items.Count = 0 Then
                        Dim nameDrawing As String = arrayTagItemSelect(1)
                        Dim putchDrawing As String = arrayTagItemSelect(2)
                        Dim statusXRef As String = arrayTagItemSelect(3)

                        Dim dictionarySelectObject As Dictionary(Of String, Dictionary(Of String, String(,))) = New Dictionary(Of String, Dictionary(Of String, String(,)))
                        If statusXRef.Trim.Length > 0 Then
                            If File.Exists(putchDrawing) = True Then
                                Dim boolSelect As Boolean = FuncReadObjectDestDraw(putchDrawing, dictionarySelectObject)
                            End If
                        Else
                            Dim boolSelect As Boolean = FuncReadObjectCurrentDraw(dictionarySelectObject)
                        End If
                        If dictionarySelectObject.Count > 0 Then
                            For i As Integer = 0 To dictionarySelectObject.Count - 1
                                Dim nameTable As String = dictionarySelectObject.ElementAt(i).Key
                                Dim dictionaryData As Dictionary(Of String, String(,)) = dictionarySelectObject.ElementAt(i).Value
                                Dim Node0 As TreeViewItem = New TreeViewItem
                                Node0.Header = nameTable
                                Dim arrayTag As String() = {nameTable, nameDrawing, putchDrawing, statusXRef}
                                Node0.Tag = arrayTag
                                Node0.FontSize = 12
                                ItemSelect.Items.Add(Node0)
                                If dictionaryData.Count > 0 Then
                                    For j As Integer = 0 To dictionaryData.Count - 1
                                        Dim handleID As String = dictionaryData.ElementAt(j).Key
                                        Dim arrayData As String(,) = dictionaryData.ElementAt(j).Value
                                        Dim nameNode1 As String = ""
                                        Dim nameNode2 As String = ""
                                        If IsArray(arrayData) = True Then
                                            For k As Integer = 0 To arrayData.GetUpperBound(1)
                                                Dim nameField As String = arrayData(0, k)
                                                Dim value As String = arrayData(1, k)
                                                If nameField Like "LABEL" Then
                                                    If value.Trim.Length > 0 And nameNode1.Trim.Length = 0 Then
                                                        nameNode1 = value
                                                    End If
                                                ElseIf nameField Like "Name" Or nameField Like "NAME" Then
                                                    If value.Trim.Length > 0 Then
                                                        nameNode1 = value
                                                    End If
                                                ElseIf nameField Like "Type_line" Then
                                                    If value.Trim.Length > 0 Then
                                                        nameNode1 = value
                                                    End If
                                                End If
                                                If nameField Like "Cad_number" Then
                                                    If value.Trim.Length > 0 Then
                                                        nameNode2 = value
                                                    End If
                                                ElseIf nameField Like "Reg_numb_border" Then
                                                    If value.Trim.Length > 0 Then
                                                        nameNode2 = value
                                                    End If
                                                ElseIf nameField Like "ConditionalNumber" Then
                                                    If value.Trim.Length > 0 Then
                                                        nameNode2 = value
                                                    End If
                                                End If
                                            Next k
                                            Dim Node1 As TreeViewItem = New TreeViewItem
                                            Node1.Header = nameNode1 & nameNode2
                                            Node0.Items.Add(Node1)
                                            Dim newDictData As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                            newDictData.Add(handleID, arrayData)
                                            Node1.Tag = newDictData
                                            Node1.Name = "Object"
                                            Node0.FontSize = 12
                                        End If
                                    Next j
                                End If
                            Next i

                        End If
                    End If
                ElseIf typeItem Like "Object" Then
                    'заполняем таблицу
                    Dim parentItem As TreeViewItem = ItemSelect.Parent
                    Dim arrayTagItemSelect As String() = parentItem.Tag 'считываем скрытые данные
                    Dim nameDrawing As String = arrayTagItemSelect(1)
                    Dim putchDrawing As String = arrayTagItemSelect(2)
                    Dim statusXRef As String = arrayTagItemSelect(3)
                    Dim dictDataTag As Dictionary(Of String, String(,)) = ItemSelect.Tag
                    Dim handleID As String = ""
                    If dictDataTag.Count > 0 Then
                        Dim arrayData As String(,) = dictDataTag.ElementAt(0).Value
                        handleID = dictDataTag.ElementAt(0).Key
                        If IsArray(arrayData) = True Then
                            custdata.Clear()
                            For i As Integer = 1 To arrayData.GetUpperBound(1)
                                Dim nameField As String = arrayData(0, i)
                                Dim valStr As String = arrayData(1, i)
                                Dim Acustim As Customer = New Customer()
                                Acustim.Имя_свойства = nameField
                                Acustim.Значение = valStr
                                custdata.Add(Acustim)
                            Next i
                        End If
                    End If
                    PropertyGridObject.HeadersVisibility = False
                    Dim pI As TreeViewItem = ItemSelect.Parent
                    Dim arrayDataTag As String() = {pI.Header, handleID, nameDrawing, putchDrawing, statusXRef}
                    PropertyGridObject.Tag = arrayDataTag
                    PropertyGridObject.ItemsSource = custdata
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка наити объект
    Private Sub ButtonZoom_Click(sender As Object, e As RoutedEventArgs) Handles ButtonZoom.Click
        Dim TreeViewControl As System.Windows.Controls.TreeView = TreeViewDWG
        If IsNothing(TreeViewControl) = False Then
            Dim itemSelectObject As Object = TreeViewControl.SelectedValue
            If IsNothing(itemSelectObject) = False Then
                If TypeOf itemSelectObject Is TreeViewItem Then
                    Dim ItemSelect As TreeViewItem = itemSelectObject
                    'Dim parentItem As TreeViewItem = FindRootTreeViewItem(ItemSelect)
                    Dim parentItem As TreeViewItem = ItemSelect.Parent
                    Dim arrayTagItemSelect As String() = parentItem.Tag 'считываем скрытые данные
                    Dim typeXRef As String = arrayTagItemSelect(3)
                    Dim putchDwg As String = arrayTagItemSelect(1)
                    Dim dictDataTag As Dictionary(Of String, String(,)) = ItemSelect.Tag
                    If dictDataTag.Count > 0 Then
                        Dim idHandle As String = dictDataTag.ElementAt(0).Key
                        If typeXRef.Trim.Length > 0 Then
                            If File.Exists(putchDwg) = True Then
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = FuncReturnEntityDeskDrawByHandle(putchDwg, idHandle)
                                Dim bound As Extents3d = acEnt.Bounds
                                SubZoom.Zoom(bound.MinPoint, bound.MaxPoint, New Point3d(), 1)
                                'XRefHighlighter.HighlightXRefEntity(arrayTagItemSelect(1), idHandle)
                            End If
                        Else
                            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                            Dim db As Database = acDoc.Database
                            Dim ed As Editor = acDoc.Editor
                            Using tr As Transaction = db.TransactionManager.StartTransaction()
                                Dim idEnt As ObjectId = AcadFunction.FuncReturnObjectID(idHandle)
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idEnt, OpenMode.ForRead)
                                Dim bound As Extents3d = acEnt.Bounds
                                'Dim centerPoint As Point3d = New Point3d((bound.MaxPoint.X + bound.MinPoint.X) / 2, (bound.MaxPoint.Y + bound.MinPoint.Y) / 2, 0)
                                'Dim minPoint1 As Point3d = New Point3d(centerPoint.X - 10, centerPoint.Y - 10, 0)
                                'Dim maxPoint1 As Point3d = New Point3d(centerPoint.X + 10, centerPoint.Y + 10, 0)
                                SubZoom.Zoom(bound.MaxPoint, bound.MinPoint, New Point3d(), 1)
                                tr.Commit()
                                'SubZoom.Zoom(bound.MinPoint, bound.MaxPoint, New Point3d(), 1)
                            End Using
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    'кнопка сформировать буферную зону для одиночного объекта
    Private Sub ButtonSelectObject_Click(sender As Object, e As RoutedEventArgs) Handles ButtonSelectObject.Click
        Dim collItem As ItemCollection = PropertyGridObject.Items
        If IsNothing(collItem) = False Then
            If collItem.Count > 0 Then
                Dim dimBufferZona As Double = 0
                For i As Integer = 0 To collItem.Count - 1
                    Dim userItem As Object = collItem.Item(i)
                    If TypeOf userItem Is Customer Then
                        Dim custData As Customer = userItem
                        If custData.Имя_свойства Like "Размер зоны" Then
                            dimBufferZona = Val(custData.Значение)
                            If dimBufferZona <= 0 Then
                                MsgBox("Размер буферной зоны задан не корректно!!!")
                                Exit Sub
                            End If
                        Else
                            MsgBox("Для задания объекту буферной зоны необходимо быть в контексте Буферные зоны!!!")
                            Exit Sub
                        End If

                    End If
                Next i
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                If doc Is Nothing Then Return
                Dim ed As Editor = doc.Editor
                Dim db As Database = doc.Database
                Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
                Dim idsEnt As List(Of ObjectId) = New List(Of ObjectId)
                Do
                    Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите полилинию: ")
                    If Not (rs.Status = PromptStatus.OK) Then
                        Exit Do
                    End If
                    If rs.Status = PromptStatus.OK Then
                        Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                        ids.Reverse()
                        ids.Add(rs.ObjectId)
                        For Each id As ObjectId In ids
                            Dim idStr As String = id.ObjectClass.Name
                            If idStr Like "AcDbPolyline" Then
                                idsEnt.Add(id)
                            End If
                            'End If
                            'If Not (id.Database.OriginalFileName Like db.Filename) Then
                            'End If
                        Next
                    End If
                Loop
                Dim netBuffer As BufferZoneCreator = New BufferZoneCreator
                Dim arrayPolyNS As List(Of Polygon) = New List(Of Polygon)
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                Dim docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Using docLock
                    Try
                        For i As Integer = 0 To idsEnt.Count - 1
                            Using acEnt As Entity = TryCast(idsEnt.Item(i).Open(OpenMode.ForRead), Entity)
                                If Not IsDBNull(acEnt) Then
                                    '===================================================================================================
                                    'читаем расширенные данные
                                    If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                        Dim acPolyLine As Polyline = acEnt
                                        If acPolyLine.Closed = True Then
                                            Dim lineNS As Polygon = netTopo.CreatePolygon(acPolyLine)
                                            Dim coordLine As List(Of Coordinate) = lineNS.Coordinates.ToList
                                            Dim polyNS As Polygon = netBuffer.CreatePolygonBuffer(coordLine, dimBufferZona)
                                            arrayPolyNS.Add(polyNS)
                                        Else
                                            Dim lineNS As LineString = netTopo.CreateLineString(acPolyLine)
                                            Dim coordLine As List(Of Coordinate) = lineNS.Coordinates.ToList
                                            Dim polyNS As Polygon = netBuffer.CreateLineBuffer(coordLine, dimBufferZona)
                                            arrayPolyNS.Add(polyNS)
                                        End If
                                    End If
                                End If
                            End Using            '   
                        Next
                        Dim tr As Transaction = db.TransactionManager.StartTransaction()
                        Using tr
                            If arrayPolyNS.Count > 0 Then
                                If arrayPolyNS.Count = 1 Then
                                    Dim polyGeo As Polygon = arrayPolyNS.Item(0)
                                    Dim numGeo As Integer = polyGeo.NumGeometries
                                    For i As Integer = 0 To numGeo - 1
                                        Dim tempGeo As NetTopologySuite.Geometries.Geometry = polyGeo.GetGeometryN(i)
                                        Dim coordPoly As List(Of Coordinate) = tempGeo.Coordinates.ToList
                                        Dim arrayCoord As Double(,) = {}
                                        Dim countArrayCoord As Integer = 0
                                        Dim startPoint As Point2d = New Point2d(-1, -1)
                                        For j As Integer = 0 To coordPoly.Count - 1
                                            Dim coordPoint As Coordinate = coordPoly.Item(j)
                                            ReDim Preserve arrayCoord(1, countArrayCoord)
                                            arrayCoord(0, countArrayCoord) = coordPoint.X
                                            arrayCoord(1, countArrayCoord) = coordPoint.Y
                                            If coordPoint.X = startPoint.X And coordPoint.Y = startPoint.Y Then
                                                If arrayCoord.GetUpperBound(1) > 0 Then
                                                    Dim polyID1 As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                                    arrayCoord = {}
                                                    countArrayCoord = 0
                                                    startPoint = New Point2d(-1, -1)
                                                End If
                                            End If
                                            countArrayCoord += 1
                                        Next j
                                        If arrayCoord.GetUpperBound(1) > 0 Then
                                            Dim polyID1 As ObjectId = AcadFunction.FuncDrawPolyline(arrayCoord)
                                            arrayCoord = {}
                                            countArrayCoord = 0
                                            startPoint = New Point2d(-1, -1)
                                        End If
                                    Next
                                End If
                            End If
                            tr.Commit()
                        End Using
                    Catch ex As System.Exception
                    End Try
                End Using
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'записать значение
    Private Sub ButtonWritePS_Click(sender As Object, e As RoutedEventArgs) Handles ButtonWritePS.Click
        Dim arrayProperty As Object = PropertyGridObject.Tag
        If IsArray(arrayProperty) = True Then
            Dim arrayPropertyTag As String() = arrayProperty
            If arrayPropertyTag.Length = 5 Then
                Dim nameTablePS As String = arrayPropertyTag(0)
                Dim handleId As String = arrayPropertyTag(1)
                Dim nameDrawing As String = arrayPropertyTag(2)
                Dim putchDrawing As String = arrayPropertyTag(3)
                Dim statusXRef As String = arrayPropertyTag(4)
                If nameTablePS.Trim.Length > 0 And handleId.Trim.Length > 0 Then
                    Dim itemColl As ItemCollection = PropertyGridObject.Items
                    If itemColl.Count > 0 Then
                        Dim arrayData As String(,) = {}
                        Dim countArrayData As Integer = 0
                        If statusXRef.Trim.Length > 0 Then
                            For Each userItem As Object In itemColl
                                If TypeOf userItem Is Customer Then
                                    Dim userCustData As Customer = userItem
                                    ReDim Preserve arrayData(1, countArrayData)
                                    arrayData(0, countArrayData) = userCustData.Имя_свойства
                                    arrayData(1, countArrayData) = userCustData.Значение
                                    countArrayData += 1
                                End If
                            Next
                            If arrayData.Length > 0 Then
                                Dim dictWriteData As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
                                dictWriteData.Add(nameTablePS, arrayData)
                                Dim boolRecData As Boolean = FuncWriteObjectDestDraw2(nameDrawing, handleId, dictWriteData)
                                If boolRecData = True Then
                                    ButtonWritePS.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                                End If
                            End If
                        Else
                            For Each userItem As Object In itemColl
                                If TypeOf userItem Is Customer Then
                                    Dim userCustData As Customer = userItem
                                    ReDim Preserve arrayData(2, countArrayData)
                                    arrayData(0, countArrayData) = nameTablePS
                                    arrayData(1, countArrayData) = userCustData.Имя_свойства
                                    arrayData(2, countArrayData) = userCustData.Значение
                                    countArrayData += 1
                                End If
                            Next
                            If arrayData.Length > 0 Then
                                'находим таблицу
                                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                                Dim db As Database = acDoc.Database
                                Try
                                    Using docLock As DocumentLock = acDoc.LockDocument()
                                        Using tr As Transaction = db.TransactionManager.StartTransaction()
                                            Dim idEnt As ObjectId = AcadFunction.FuncReturnObjectID(handleId)
                                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idEnt, OpenMode.ForWrite)
                                            If acEnt.Id.IsNull = False Then
                                                Dim boolPSTables As Boolean = FuncPropertySet.FuncFindPSetTable(acEnt, nameTablePS)
                                                If boolPSTables = True Then
                                                    Dim boolADDDATA As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, arrayData)
                                                    If boolADDDATA = True Then
                                                        ButtonWritePS.Background = New SolidColorBrush(Color.FromRgb(144, 238, 144))
                                                    End If
                                                End If
                                            End If
                                            tr.Commit()
                                        End Using
                                    End Using
                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                End Try
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор
    Private Sub PropertyGridObject_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles PropertyGridObject.MouseRightButtonDown
        Dim selectedItem As Object = PropertyGridObject.SelectedItem
        Dim newCustData As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
        If IsNothing(selectedItem) = False Then
            Dim userCustomer As Customer = selectedItem
            Dim nameField As String = userCustomer.Имя_свойства
            If IsNothing(nameField) = False Then
                If nameField.Trim.Length > 0 Then
                    Dim putchXmlFile As String = putshXMLData & "FieldData.xml"
                    If File.Exists(putchXmlFile) = True Then
                        Dim doc As XmlDocument = New XmlDocument()
                        doc.Load(putchXmlFile)
                        Dim nodes As XmlNodeList = doc.GetElementsByTagName(nameField)
                        If nodes.Count > 0 Then
                            For Each node As XmlNode In nodes
                                Dim text As String = node.InnerXml
                                Dim categories As List(Of LandCategory) = ParseLandCategories(text)
                                Dim arraySelect As String() = {}
                                If categories.Count > 0 Then
                                    Dim countArraySelect As Integer = 0
                                    For i As Integer = 0 To categories.Count - 1
                                        ReDim Preserve arraySelect(countArraySelect)
                                        arraySelect(countArraySelect) = categories.Item(i).Name
                                        countArraySelect += 1
                                    Next
                                End If
                                If arraySelect.Length > 0 Then
                                    Dim selectForm As SelectTemplateFiles = New SelectTemplateFiles
                                    selectForm.ListBox1.DataSource = arraySelect
                                    selectForm.ShowDialog()
                                    If selectForm.boolShow = True Then
                                        Dim NewValue As String = selectForm.ListBox1.Text
                                        userCustomer.Значение = NewValue
                                        custdata.Clear()
                                        PropertyGridObject.SelectedItem = userCustomer
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'найти узел
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim searchText As String = txtSearch.Text.Trim()
        If String.IsNullOrEmpty(searchText) Then Return
        FindAndSelectNode(searchText)
    End Sub
    Private Sub FindAndSelectNode(searchText As String)
        ' Снимаем предыдущее выделение
        ClearSelection()
        Dim foundItem As TreeViewItem = FindTreeViewItemAdvancedFromEnd(TreeViewDWG, searchText, SearchFromEndMode.EndsWith)
        If foundItem IsNot Nothing Then
            foundItem.IsSelected = True
            foundItem.BringIntoView()
            foundItem.Focus()
        Else
            MsgBox($"Узел с текстом '{searchText}' не найден", "Поиск")
        End If
    End Sub

    Private Shared Function FindTreeViewItemAdvancedFromEnd(parent As ItemsControl, searchText As String,
                                               mode As SearchFromEndMode,
                                               Optional comparisonType As StringComparison = StringComparison.OrdinalIgnoreCase) As TreeViewItem

        For i As Integer = 0 To parent.Items.Count - 1
            Dim treeViewItem As TreeViewItem = TryCast(parent.ItemContainerGenerator.ContainerFromIndex(i), TreeViewItem)

            If treeViewItem IsNot Nothing Then
                Dim headerText As String = treeViewItem.Header.ToString()
                Dim isMatch As Boolean = False

                Select Case mode
                    Case SearchFromEndMode.EndsWith
                        ' Точное совпадение с конца
                        isMatch = headerText.EndsWith(searchText, comparisonType)

                    Case SearchFromEndMode.ContainsFromEnd
                        ' Текст содержится ближе к концу строки (последняя треть)
                        Dim lastThirdStart As Integer = Math.Max(0, headerText.Length - headerText.Length \ 3)
                        isMatch = headerText.IndexOf(searchText, lastThirdStart, comparisonType) >= 0

                    Case SearchFromEndMode.LastIndexOf
                        ' Последнее вхождение текста
                        Dim lastIndex As Integer = headerText.LastIndexOf(searchText, comparisonType)
                        Dim threshold As Integer = headerText.Length \ 2 ' Во второй половине строки
                        isMatch = lastIndex >= threshold
                End Select

                If isMatch Then
                    Return treeViewItem
                End If

                ' Рекурсивный поиск в дочерних элементах
                treeViewItem.IsExpanded = True
                Dim foundChild As TreeViewItem = FindTreeViewItemAdvancedFromEnd(treeViewItem, searchText, mode, comparisonType)
                If foundChild IsNot Nothing Then
                    Return foundChild
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Shared Function FindVisualChild(Of T As DependencyObject)(parent As DependencyObject) As T
        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(parent, i)
            If child IsNot Nothing AndAlso TypeOf child Is T Then
                Return DirectCast(child, T)
            Else
                Dim childOfChild As T = FindVisualChild(Of T)(child)
                If childOfChild IsNot Nothing Then
                    Return childOfChild
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Sub ClearSelection()
        For Each item As Object In TreeViewDWG.Items
            Dim treeViewItem As TreeViewItem = TreeViewDWG.ItemContainerGenerator.ContainerFromItem(item)
            If treeViewItem IsNot Nothing Then
                ClearChildSelection(treeViewItem)
            End If
        Next
    End Sub

    Private Sub ClearChildSelection(parent As TreeViewItem)
        parent.IsSelected = False
        For Each item As Object In parent.Items
            Dim treeViewItem As TreeViewItem = parent.ItemContainerGenerator.ContainerFromItem(item)
            If treeViewItem IsNot Nothing Then
                ClearChildSelection(treeViewItem)
            End If
        Next
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As RoutedEventArgs)
        txtSearch.Clear()
        ClearSelection()
    End Sub

    Private Sub PropertyGridObject_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles PropertyGridObject.SelectedCellsChanged

    End Sub

    Private Sub PropertyGridObject_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs) Handles PropertyGridObject.PreviewMouseDown
        ButtonWritePS.Background = New SolidColorBrush(Color.FromRgb(221, 221, 221))
    End Sub
End Class
