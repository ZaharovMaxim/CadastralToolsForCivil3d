Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class ObjectCopyProperties
    '=============================================================================================
    'рисование объектов по аналогу
    <CommandMethod("ObjCopyProperties")>
    Public Sub ObjCopyProperties()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim NameLayer As String = "0" 'слой
        Dim LineUWight As Double = 20 'толщина линии
        Dim LineScale As Double = 1 'масштаб 
        Dim LineColor As Integer = 7 'цвет
        Dim LineType As String = "" 'тип линии
        Dim BlkScaleXY As Double = 1
        Dim BlkScaleZ As Double = 1
        Dim RotationBlk As Double = 0
        Dim GlobalWidtPline As Double = 0 'глобальная ширина
        Dim NameObject As String = "" 'имя блока, код точки сого, имя линии, содержимое текста
        Dim ArrayODData As String(,) = Nothing
        Dim ArrayTblName As String() = Nothing
        Dim boolBlk As Boolean = False
        Dim boolLine As Boolean = False
        Dim boolCoGo As Boolean = False
        '1. Делаем выбор либо таблица либо объекта
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pKeyOpts = New PromptKeywordOptions("")
        pKeyOpts.Message = vbLf & "Укажите объекты для копирования? "
        pKeyOpts.Keywords.Add("Свойствапримитива")
        pKeyOpts.Keywords.Add("Наборхарактеристик")
        pKeyOpts.Keywords.Add("Все")
        pKeyOpts.Keywords.Default = "Все"
        Dim pKeyRes As PromptResult = ed.GetKeywords(pKeyOpts)
        If pKeyRes.Status = PromptStatus.OK Then
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim SelectLine As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект в качестве образца (Блок, COGO-точку, Полилинию, Характерную линию): ")
                If SelectLine.Status = PromptStatus.Cancel Then
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                'ссылка на объект-аналог
                Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectLine.ObjectId, OpenMode.ForRead)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'исходные данные
                NameLayer = obj.Layer 'слой
                LineUWight = obj.LineWeight 'толщина линии
                LineScale = obj.LinetypeScale 'масштаб 
                LineColor = obj.ColorIndex 'цвет
                LineType = obj.Linetype 'тип линии
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'читаем расширенные данные
                Dim boolReadTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(obj, ArrayTblName)
                If IsArray(ArrayTblName) = True Then
                    Dim boolPSReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(obj, ArrayODData)
                End If
                'считываем пользовательские свойства
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'читаем свойства
                If TypeOf obj Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = obj
                    GlobalWidtPline = Pline.ConstantWidth
                    boolLine = True
                ElseIf TypeOf obj Is Polyline2d Then
                    Dim Pline As Polyline2d = obj
                    GlobalWidtPline = Pline.ConstantWidth
                    boolLine = True
                ElseIf TypeOf obj Is CogoPoint Then
                    Dim COGOPoint As CogoPoint = obj
                    NameObject = COGOPoint.RawDescription
                    BlkScaleXY = COGOPoint.ScaleXY
                    BlkScaleZ = COGOPoint.ScaleZ
                    RotationBlk = COGOPoint.MarkerRotation
                    boolCoGo = True
                ElseIf TypeOf obj Is BlockReference Then
                    Dim Blk As BlockReference = obj
                    BlkScaleXY = Blk.ScaleFactors.X
                    BlkScaleZ = Blk.ScaleFactors.Z
                    RotationBlk = Blk.Rotation
                    If Blk.IsDynamicBlock = True Then
                        Dim blockDyn As BlockTableRecord = acTrans.GetObject(Blk.DynamicBlockTableRecord, OpenMode.ForWrite)
                        NameObject = blockDyn.Name
                    Else
                        NameObject = Blk.Name
                    End If
                    boolBlk = True
                End If
            End Using
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объекты
            Dim rng As Random = New Random()
            Do
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim SelectLineObj As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Блок, COGO-Точку или Полилинию: ")
                    If SelectLineObj.Status = PromptStatus.Cancel Then
                        acTrans.Commit()
                        Exit Sub
                    End If
                    Dim obj1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectLineObj.ObjectId, OpenMode.ForWrite)
                    If pKeyRes.StringResult Like "Свойствапримитива" Or pKeyRes.StringResult Like "Все" Then
                        obj1.Layer = NameLayer
                        obj1.LineWeight = LineUWight   'толщина линии
                        obj1.LinetypeScale = LineScale   'масштаб 
                        obj1.ColorIndex = LineColor   'цвет
                        obj1.Linetype = LineType   'тип линии
                        If TypeOf obj1 Is Autodesk.AutoCAD.DatabaseServices.Polyline And boolLine = True Then
                            Dim Pline As Autodesk.AutoCAD.DatabaseServices.Polyline = obj1
                            Pline.ConstantWidth = GlobalWidtPline
                        ElseIf TypeOf obj1 Is Autodesk.AutoCAD.DatabaseServices.Polyline2d And boolLine = True Then
                            Dim Pline As Polyline2d = obj1
                            Pline.ConstantWidth = GlobalWidtPline
                        ElseIf TypeOf obj1 Is BlockReference And boolBlk = True Then
                            Dim Blk As BlockReference = obj1
                            If NameObject.Trim.Length > 0 Then
                                Dim NewBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(Blk.Position, NameObject, NameLayer, LineColor, BlkScaleXY, LineUWight, RotationBlk, BlkScaleZ)
                                If IsNothing(NewBlk) = False Then
                                    If NewBlk.Id.IsNull = False Then
                                        If NewBlk.IsWriteEnabled = False Then
                                            NewBlk.UpgradeOpen()
                                        End If
                                        obj1.Erase()
                                    End If
                                End If
                            End If
                        ElseIf TypeOf obj1 Is CogoPoint Then
                            Dim PtCOGO As CogoPoint = obj1
                            If NameObject.Trim.Length > 0 Then
                                PtCOGO.RawDescription = NameObject
                                PtCOGO.ScaleXY = BlkScaleXY
                                PtCOGO.ScaleZ = BlkScaleZ
                                PtCOGO.MarkerRotation = RotationBlk
                                PtCOGO.ApplyDescriptionKeys()
                            End If
                        End If
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'считываем PS Таблицы
                    If pKeyRes.StringResult Like "Наборхарактеристик" Or pKeyRes.StringResult Like "Все" Then
                        Dim ArrayPSTableObj As String() = Nothing
                        Dim boolReadTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(obj1, ArrayPSTableObj)
                        'удаляем лишние таблицы
                        If IsArray(ArrayPSTableObj) = True Then
                            For i As Integer = 0 To ArrayPSTableObj.Length - 1
                                Dim NameTlb As String = ArrayPSTableObj(i)
                                If IsNothing(NameTlb) Then
                                    Dim boolDelPS As Boolean = FuncPropertySet.FuncDeleteDataSetTables(obj1, NameTlb)
                                End If
                            Next
                        End If
                        'добавляем таблицу в примитив
                        If IsArray(ArrayTblName) = True Then
                            For i As Integer = 0 To ArrayTblName.Length - 1
                                Dim NameTableIns As String = ArrayTblName(i)
                                If IsNothing(NameTableIns) = False Then
                                    Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameTableIns)
                                    Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(obj1, CreateTBLPSID)
                                End If
                            Next
                            If IsArray(ArrayODData) = True Then
                                For i As Integer = 0 To ArrayODData.GetUpperBound(1)
                                    If ArrayODData(1, i).Trim Like "LOCALID" Then
                                        ArrayODData(2, i) = rng.Next.ToString
                                    End If
                                Next i
                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(obj1, ArrayODData)
                            End If
                        End If
                    End If
                    acTrans.Commit()
                End Using
            Loop
        End If
    End Sub
End Class
