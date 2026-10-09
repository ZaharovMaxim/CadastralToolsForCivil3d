Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports Autodesk.Civil.ApplicationServices
Imports NetTopologySuite
Public Class RestoreObject
    '=============================================================================================
    'восстановить поле Area в метке участка
    <CommandMethod("RestoreAreaObject")>
    Public Sub InsertUserCadastralPoints()
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
        'Dim ArrayTablesPS As String() = {"PROJECT_PARCEL", "PROJECT_BUILDING", "PROJECT_ZONA", "URBAN_LINE", "PROJECT_LINE"}
        Dim ArrayParcels As String(,) = Nothing
        Dim CountArrayParcels As Integer = 0
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    Dim valStr As String = FuncPropertySet.FuncReadValueSetProperties(ent, "LOCALID")
                    If IsNothing(valStr) = False Then
                        If valStr.Length > 0 Then
                            ReDim Preserve ArrayParcels(2, CountArrayParcels)
                            ArrayParcels(0, CountArrayParcels) = valStr
                            ArrayParcels(1, CountArrayParcels) = ent.Handle.ToString
                            CountArrayParcels += 1
                        End If
                    End If
                End If
            Next

            If IsArray(ArrayParcels) = False Then
                MsgBox("Объекты не найдены!!! Макро прервано")
                Exit Sub
            End If

            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.MLeader Then
                                    Dim acMLeader As Autodesk.AutoCAD.DatabaseServices.MLeader = acEnt
                                    'массив с расширенными данными
                                    Dim valStrLocalIdMLeader As String = FuncPropertySet.FuncReadValueSetProperties(acMLeader, "IDRNG")
                                    If IsNothing(valStrLocalIdMLeader) = False Then
                                        If valStrLocalIdMLeader.Length > 0 Then
                                            For i As Integer = 0 To ArrayParcels.GetUpperBound(1)
                                                Dim valLocalidParcel As String = ArrayParcels(0, i)
                                                If valLocalidParcel Like valStrLocalIdMLeader Then
                                                    Dim idPline As ObjectId = AcadFunction.FuncReturnObjectID(ArrayParcels(1, i))
                                                    If IsNothing(idPline) = False Then
                                                        If idPline.IsNull = False Then
                                                            If acMLeader.IsWriteEnabled = False Then
                                                                acMLeader.UpgradeOpen()
                                                            End If
                                                            Dim acMtext As String = acMLeader.MText.Contents
                                                            '============================================================================================
                                                            'разбираем строку
                                                            Dim id As ObjectId = acMLeader.GetField() 'смотрим старый атрибут
                                                            Dim fd As Field = tr.GetObject(id, OpenMode.ForWrite) 'читаем его
                                                            Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
                                                            Dim pos2 As Integer = CodStrFiled.IndexOf("Area")
                                                            If pos2 > -1 Then
                                                                Dim fstr As String = FuncCadastr.FuncCreateFieldSQRTPolyline(idPline)
                                                                Dim pos3 As Integer = CodStrFiled.IndexOf("Площадь")
                                                                Dim pos4 As Integer = CodStrFiled.IndexOf("кв")
                                                                If pos3 > -1 And pos4 > 1 Then
                                                                    Dim str1 As String = Mid(CodStrFiled, 1, pos3 + 7).Trim
                                                                    Dim str2 As String = Mid(CodStrFiled, pos4 + 1).Trim
                                                                    fd.SetFieldCode(str1 & " " & fstr & " " & str2)
                                                                    fd.Evaluate()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next i
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
            End If
            tr.Commit()
        End Using
    End Sub
    '==================================================================================================
    'перенумеровать участки
    <CommandMethod("RenumberParcels")>
    Public Sub RenumberParcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim netSuiteTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '=====================================================================================================
            'выбираем трассу
            Dim alignPoly As Polyline = New Polyline
            Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
            pKeyOpts.Message = vbLf & "Выберите объект: "
            pKeyOpts.Keywords.Add("Трасса")
            pKeyOpts.Keywords.Add("Полилиния")
            pKeyOpts.Keywords.Add("Контура")
            pKeyOpts.AllowNone = False
            Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)
            If pKeyRes.StringResult Like "Трасса" Then
                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу: ")
                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                    Dim align As Autodesk.Civil.DatabaseServices.Alignment = ent1
                    Dim idPline As ObjectId = align.GetPolyline
                    alignPoly = idPline.GetObject(OpenMode.ForRead)
                End If
            ElseIf pKeyRes.StringResult Like "Полилиния" Then
                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите полилинию: ")
                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                If TypeOf ent1 Is Polyline Then
                    alignPoly = ent1
                End If
            ElseIf pKeyRes.StringResult Like "Контура" Then
                Dim ppr1 As PromptPointResult = ed.GetPoint("Вуберите точку 1")
                Dim pt1 As Point3d = ppr1.Value
                Dim ppr2 As PromptPointResult = ed.GetPoint("Вуберите точку 2")
                Dim pt2 As Point3d = ppr2.Value
                alignPoly.AddVertexAt(0, New Point2d(pt1.X, pt1.Y), 0, 0, 0)
                alignPoly.AddVertexAt(1, New Point2d(pt2.X, pt2.Y), 0, 0, 0)
            End If
            Dim numberParcel As Integer = ed.GetInteger("Введите начальный номер участка: ").Value
            '=========================================================================================================
            'выбираем участки
            Dim arrayDistanceParcel As String(,) = {}
            Dim countArrayDistanceParcel As Integer = 0
            Dim acSSPrompt As PromptSelectionResult = Nothing
            acSSPrompt = acDoc.Editor.GetSelection
            'смотрим выбор объектов
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                            Dim nameTablePS As String() = {}
                            Dim boolFindTable As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acPoly, nameTablePS)
                            If nameTablePS.Length > 0 Then
                                Dim netPolygon As NetTopologySuite.Geometries.Polygon = netSuiteTopo.CreatePolygon(acPoly)
                                Dim inPoint As NetTopologySuite.Geometries.Point = netSuiteTopo.GetInteriorPoint(netPolygon)
                                Dim pt3d As Point3d = New Point3d(inPoint.X, inPoint.Y, 0)
                                Dim pt As Point3d = alignPoly.GetClosestPointTo(pt3d, True)
                                Dim dist As Double = alignPoly.GetDistAtPoint(pt)
                                ReDim Preserve arrayDistanceParcel(2, countArrayDistanceParcel)
                                arrayDistanceParcel(0, countArrayDistanceParcel) = dist
                                arrayDistanceParcel(1, countArrayDistanceParcel) = acPoly.Handle.ToString
                                arrayDistanceParcel(2, countArrayDistanceParcel) = nameTablePS(0)
                                countArrayDistanceParcel += 1
                            End If

                        End If
                    End If
                Next
            End If
            If arrayDistanceParcel.GetUpperBound(1) > 0 Then
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayDistanceParcel, 0)
            End If
            For i As Integer = 0 To arrayDistanceParcel.GetUpperBound(1)
                Dim hgPoly As String = arrayDistanceParcel(1, i)
                Dim idPoly As ObjectId = AcadFunction.FuncReturnObjectID(hgPoly)
                Dim poly As Polyline = tr.GetObject(idPoly, OpenMode.ForWrite)
                Dim nameTable As String = arrayDistanceParcel(2, i)
                Dim conditionalParcel As String = FuncPropertySet.FuncReadValueSetProperties(poly, "ConditionalNumber")
                Dim ArrayOBJ1 As String(,) = Nothing
                ReDim Preserve ArrayOBJ1(2, 0)
                ArrayOBJ1(0, 0) = nameTable
                ArrayOBJ1(1, 0) = "ConditionalNumber"
                ArrayOBJ1(2, 0) = numberParcel
                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(poly, ArrayOBJ1)
                If boolWriteData = True Then
                    numberParcel += 1
                End If
                'ищем мультивыноску
                'Dim localId As String = FuncPropertySet.FuncReadValueSetProperties(poly, "LOCALID")
                'If localId.Trim.Length > 2 Then
                '    Dim acMleader As MLeader = FuncCadastr.FuncFindMLeaderByLocalId(localId)
                '    If IsNothing(acMleader) = False Then
                '        Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(poly, "Cad_number")
                '        cad_number = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(cad_number)
                '        Dim OboznParcel As String = cad_number & ":ЗУ" & numberParcel


                '        Dim pos1 As Integer = OboznParcel.IndexOf(Chr(10))
                '        If pos1 > -1 Then
                '            OboznParcel = Mid(OboznParcel, pos1 + 1)
                '        End If
                '        Dim htxt As Double = acMleader.MText.TextHeight
                '        Dim idStyle As ObjectId = acMleader.MText.TextStyleId
                '        Dim acMTextl As MText = New MText()
                '        acMTextl.SetDatabaseDefaults()
                '        acMTextl.Contents = NumberZuTemp
                '        acMTextl.TextHeight = htxt
                '        acMTextl.TextStyleId = idStyle
                '        acMleader.MText = acMTextl
                '    End If
                'End If
            Next i
            tr.Commit()
        End Using
    End Sub
    '========================================================================================================
    'заменить участок на новый
    <CommandMethod("ReplaseLandUserObject")>
    Public Sub ReplaseLandUserObject()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'таблицы xml
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        ReDim ArrayTablesXML(1, 6)
        ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
        ArrayTablesXML(1, 0) = "Участок постоянного отвода"
        ArrayTablesXML(0, 1) = "TEMPORARY_PARCEL"
        ArrayTablesXML(1, 1) = "Участок временного отвода"
        ArrayTablesXML(0, 2) = "CADASTR_OSTROV"
        ArrayTablesXML(1, 2) = "Внутренний островок"
        ArrayTablesXML(0, 3) = "MPLAN_PARCEL"
        ArrayTablesXML(1, 3) = "Участок для постановки на ГКН"
        ArrayTablesXML(0, 4) = "PROJECT_BUILDING"
        ArrayTablesXML(1, 4) = "Здание или сооружение проектируемое"
        ArrayTablesXML(0, 5) = "PROJECT_ZONA"
        ArrayTablesXML(1, 5) = "Зона с особыми условиями"
        ArrayTablesXML(0, 6) = "PROJECT_LINE"
        ArrayTablesXML(1, 6) = "Сооружение линейного типа"
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
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

Line1:
        Dim per As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект который требуется заменить (полилиния): ")
        Dim per1 As PromptEntityResult = ed.GetEntity(vbLf & "Выберите новый объект (полилиния): ")
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Try
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'выбор полигона которому принадлежит этот контур
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект первоначальный
                Dim obj1 As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(per1.ObjectId, OpenMode.ForWrite) 'ссылка на объект первоначальный
                Dim acPoly As Polyline = obj1
                Dim NewIdParcel As String = obj1.Id.ToString
                'находим центройд объекта
                Dim CoordText As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(acPoly) 'находим центр полигона
                'вычленяем номер контура
                Dim pos1 As Integer = NewIdParcel.IndexOf("(")
                Dim pos2 As Integer = NewIdParcel.IndexOf(")")
                If pos1 > -1 And pos2 > -1 Then
                    'вычленяем префикс
                    NewIdParcel = Mid(NewIdParcel, pos1 + 2, pos2 - pos1 - 1)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'считываем PropertySet
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ArrayPSData As String(,) = Nothing
                Dim boolPropSet As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(obj, ArrayPSData, ArrayTablesXML)
                'читаем таблицы с новой линии
                Dim ArrayPSTblTemp As String() = Nothing
                Dim boolPropSetTBL As Boolean = FuncPropertySet.FuncReadAllDataSetTables(obj1, ArrayPSTblTemp, ArrayTablesXML)
                If IsArray(ArrayPSTblTemp) = True Then
                    'удаляем таблицы
                    Dim boolPropSetTBLDel As Boolean = FuncPropertySet.FuncDeleteDataSetTables1(obj1, ArrayPSTblTemp)
                End If
                If IsArray(ArrayPSData) = True Then
                    'записываем в PropertySet
                    Dim PropertySetID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayPSData(0, 0))
                    Dim boolPropertySetAdd As Boolean = FuncPropertySet.FuncAddDataSetProperties(obj1, PropertySetID)
                    Dim boolWr2 As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(obj1, ArrayPSData)
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '=================================================================================================
                'считываем индивидуальные характеристики выбранного контура
                Dim LayerObj As String = "0"
                Dim LWeight As Double = 20
                Dim LColor As Integer = 10
                Dim GWignt As Double = 0
                Dim TypeLineObj As ObjectId = Nothing
                If TypeOf obj Is Polyline Then
                    Dim Pline As Polyline = obj
                    LayerObj = Pline.Layer
                    LWeight = Pline.LineWeight
                    TypeLineObj = Pline.LinetypeId
                    LColor = Pline.ColorIndex
                    GWignt = Pline.ConstantWidth
                End If
                'заносим характеристики в новую линию
                Dim Pline2 As Polyline = Nothing
                If TypeOf obj1 Is Polyline Then
                    Pline2 = obj1
                    Pline2.Layer = LayerObj
                    Pline2.LineWeight = LWeight
                    Pline2.LinetypeId = TypeLineObj
                    Pline2.ColorIndex = LColor
                    Pline2.ConstantWidth = GWignt
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'находим локальный ID выбранного контура
                Dim idLocal As String = ""
                If idLocal.Trim.Length = 0 Then
                    If IsArray(ArrayPSData) = True Then
                        For i As Integer = 0 To ArrayPSData.GetUpperBound(1)
                            Dim NameField As String = ArrayPSData(1, i)
                            Dim val As String = ArrayPSData(2, i)
                            If NameField Like "LOCALID" Then
                                idLocal = val
                            End If
                        Next i
                    End If
                End If
                If idLocal.Trim.Length = 0 Then
                    MsgBox("Локальный ID не назначен!!!")
                    Exit Sub
                End If
                '=======================================================================================================================
                'ищем внутренние островки для старого контура
                Dim ArrayOstrov As String(,) = Nothing
                Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLocal, ArrayOstrov)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем обозначение участка по локальному id
                Dim ConditionNumber As String = ""
                Dim MLeaderObj As MLeader = Nothing
                Dim MtextObj As MText = Nothing
                MLeaderObj = FuncCadastr.FuncFindMLeaderByLocalId(idLocal)
                If IsNothing(MLeaderObj) = False Then
                    MtextObj = MLeaderObj.MText
                    ConditionNumber = MLeaderObj.MText.Contents
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'вычленяем префикс участка
                If MtextObj.HasFields = True Then
                    ConditionNumber = MtextObj.getMTextWithFieldCodes 'смотрим старый атрибут
                    Dim pos5 As Integer = ConditionNumber.IndexOf("ObjId ")
                    Dim pos6 As Integer = ConditionNumber.IndexOf(">%).Area")
                    If pos5 > -1 And pos6 > -1 Then
                        Dim TempID As String = Mid(ConditionNumber, pos5 + 7, pos6 - pos5 - 6)
                        ConditionNumber = ConditionNumber.Replace(TempID, NewIdParcel)
                    End If
                End If
                If IsNothing(MLeaderObj) = False Then
                    If MLeaderObj.IsWriteEnabled = False Then MLeaderObj.UpgradeOpen()
                    Dim htxt As Double = MLeaderObj.MText.TextHeight
                    Dim idStyle As ObjectId = MLeaderObj.MText.TextStyleId
                    Dim rot As Double = MLeaderObj.MText.Rotation
                    Dim acMTextl As MText = New MText()
                    acMTextl.SetDatabaseDefaults()
                    acMTextl.Contents = ConditionNumber
                    acMTextl.TextHeight = htxt
                    acMTextl.TextStyleId = idStyle
                    acMTextl.Rotation = rot
                    If MLeaderObj.IsWriteEnabled = False Then MLeaderObj.UpgradeOpen()
                    MLeaderObj.MText = acMTextl
                    MLeaderObj.SetFirstVertex(0, CoordText)
                Else
                    If MtextObj.IsWriteEnabled = False Then MtextObj.UpgradeOpen()
                    MtextObj.Contents = ConditionNumber
                End If
                If obj.IsWriteEnabled = False Then obj.UpgradeOpen()
                obj.Erase()
                tr.Commit()
            Catch ex As System.Exception
                tr.Commit()
            End Try
        End Using
    End Sub
    '========================================================================================================
    'заменить LOCALID
    <CommandMethod("ReplaseLocalIDObect")>
    Public Sub ReplaseLocalIDObect()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
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
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Try
                '=========================================================================================================
                'выбираем участки
                Dim rng As Random = New Random()
                Dim acSSPrompt As PromptSelectionResult = Nothing
                acSSPrompt = acDoc.Editor.GetSelection
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                Dim ArrayData As String(,) = Nothing
                                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, ArrayData)
                                If IsArray(ArrayData) = True Then
                                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Dim NameField As String = ArrayData(1, i)
                                        If NameField Like "LOCALID" Then
                                            ArrayData(2, i) = rng.Next.ToString
                                        End If
                                    Next i
                                    Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acEnt, ArrayData)
                                End If
                            End If
                        End If
                    Next
                End If
            Catch ex As System.Exception
            End Try
            tr.Commit()
        End Using
    End Sub
    '=============================================================================================
    'восстановить метки для участков
    <CommandMethod("RestoreLabelObject")>
    Public Sub RestoreLabelObject()
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
        Dim formRectoreLabel As FormRestoreLabel = New FormRestoreLabel
        formRectoreLabel.ShowDialog()
        If formRectoreLabel.boolShow = False Then
            Exit Sub
        End If
        Dim layerMLeader As String = formRectoreLabel.ComboBox3.Text
        Dim styleMLeader As String = formRectoreLabel.ComboBox2.Text
        Dim idStyle As ObjectId = AcadFunction.FuncReturnIDMLeaderStylebyName(styleMLeader)
        Dim hText As Double = formRectoreLabel.NumericUpDown1.Value
        Dim templateXML As String = formRectoreLabel.ComboBox4.Tag & formRectoreLabel.ComboBox4.Text & ".xml"
        Dim templateDWG As String = formRectoreLabel.ComboBox4.Tag & formRectoreLabel.ComboBox4.Text & ".dwg"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("CADASTRAL_OBJECT")
            If idTablePs.IsNull = True Then
                If File.Exists(templateXML) = True Then
                    idTablePs = FuncPropertySet.FuncCreateTablePSByXML(templateXML, "CADASTRAL_OBJECT")
                End If
            End If
            If IsNothing(idStyle) = True Then
                If File.Exists(templateDWG) = True Then
                    Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(templateDWG, styleMLeader, False)
                End If
            ElseIf idStyle.IsNull = True Then
                If File.Exists(templateDWG) = True Then
                    Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(templateDWG, styleMLeader, False)
                End If
            End If
            Dim acTextTbl As TextStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead)
            Dim idTextStyle As ObjectId = db.Textstyle
            If styleMLeader Like "CADASTRAL_OBJECT" Then
                If acTextTbl.Has("CADASTRAL_OBJECT") = True Then
                    idTextStyle = acTextTbl.Item("CADASTRAL_OBJECT")
                End If
            End If
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            'выбираем объекты
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                'обозначение участка
                                Dim valConditionalNumber As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "ConditionalNumber")
                                Dim LocalID As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "LOCALID")
                                If formRectoreLabel.ComboBox1.Text Like "00:00:0000000:0000:ЗУ" Then
                                    Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number")
                                    cad_number = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(cad_number)
                                    valConditionalNumber = cad_number & ":ЗУ" & valConditionalNumber
                                ElseIf formRectoreLabel.ComboBox1.Text Like "00:00:0000000:0000/чзу" Then
                                    Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number")
                                    cad_number = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(cad_number)
                                    valConditionalNumber = cad_number & "/чзу" & valConditionalNumber
                                Else
                                    valConditionalNumber = formRectoreLabel.ComboBox1.Text & valConditionalNumber
                                End If
                                'площадь участка
                                Dim strSqr As String = ""
                                If valConditionalNumber.Trim.Length > 0 Then
                                    '==================================================================================================
                                    If formRectoreLabel.RadioButton2.Checked = True Then
                                        If TypeOf acEnt Is Polyline Then
                                            '=============================================================================================================================
                                            'ищем внутренние островки
                                            Dim ArrayOstrov As String(,) = Nothing
                                            Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov)
                                            If IsArray(ArrayOstrov) = True Then
                                                Dim ArrayIDOstrov As ObjectId() = Nothing
                                                Dim CountArrayIDOstrov As Integer = 0
                                                For j1 As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                                                    Dim hgOstrov As String = ArrayOstrov(1, j1)
                                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                                    If IsNothing(idOstrov) = False Then
                                                        If idOstrov.IsNull = False Then
                                                            ReDim Preserve ArrayIDOstrov(CountArrayIDOstrov)
                                                            ArrayIDOstrov(CountArrayIDOstrov) = idOstrov
                                                            CountArrayIDOstrov += 1
                                                        End If
                                                    End If
                                                Next j1
                                                strSqr = "Площадь " & FuncCadastr.FuncCreateFieldSQRTPolyline(acEnt.Id, ArrayIDOstrov) & " кв.м."
                                            Else
                                                'формируем площадь участка (делаем поле)
                                                Dim idObj As String = acEnt.Id.ToString
                                                Dim pos1 As Integer = idObj.IndexOf("(")
                                                Dim pos2 As Integer = idObj.IndexOf(")")
                                                If pos1 > -1 And pos2 > 1 Then
                                                    idObj = Mid(idObj, pos1 + 2, pos2 - (pos1 + 1))
                                                    strSqr = "Площадь " & "%<\AcObjProp Object(%<\_ObjId " & idObj & ">%).Area \f " & """" & "%lu2%pr0" & """" & ">%" & " кв.м."
                                                End If
                                            End If
                                        End If
                                    End If
                                    Dim MLeader1 As MLeader = New MLeader
                                    '==================================================================================================================================
                                    'ищем мультивыноску
                                    Dim tempMleader As MLeader = FuncCadastr.FuncFindMLeaderByLocalId(LocalID)
                                    If IsNothing(tempMleader) = False Then
                                        MLeader1 = tr.GetObject(tempMleader.ObjectId, OpenMode.ForWrite)
                                        If MLeader1.IsWriteEnabled = True Then
                                            Dim acMText1 As MText = MLeader1.MText
                                            acMText1.SetDatabaseDefaults()
                                            Dim LABEL As String = ""
                                            If formRectoreLabel.RadioButton1.Checked = True Then
                                                acMText1.Contents = valConditionalNumber
                                                LABEL = "Обозначение участка"
                                            ElseIf formRectoreLabel.RadioButton2.Checked = True Then
                                                acMText1.Contents = valConditionalNumber & vbLf & strSqr
                                                LABEL = "Обозначение участка и площадь"
                                            End If
                                            MLeader1.MText = acMText1
                                            Dim ArrayOBJ1 As String(,) = Nothing
                                            ReDim Preserve ArrayOBJ1(2, 2)
                                            ArrayOBJ1(0, 0) = "CADASTRAL_OBJECT"
                                            ArrayOBJ1(1, 0) = "LABEL"
                                            ArrayOBJ1(2, 0) = LABEL
                                            ArrayOBJ1(0, 1) = "CADASTRAL_OBJECT"
                                            ArrayOBJ1(1, 1) = "Number"
                                            ArrayOBJ1(2, 1) = "1"
                                            ArrayOBJ1(0, 2) = "CADASTRAL_OBJECT"
                                            ArrayOBJ1(1, 2) = "IDRNG"
                                            ArrayOBJ1(2, 2) = LocalID
                                            If IsNothing(idTablePs) = False Then
                                                If idTablePs.IsNull = False Then
                                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idTablePs)
                                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MLeader1, ArrayOBJ1)
                                                End If
                                            End If

                                        End If
                                    Else
                                        MLeader1 = New MLeader
                                        Dim CoordText As Point3d = MathFunction.FuncFindCenterObjectByGeometricExtents(acEnt) 'находим центр объекта
                                        Dim MaxPoint As Point3d = acEnt.GeometricExtents.MaxPoint
                                        If TypeOf acEnt Is Polyline Then
                                            Dim acPoly As Polyline = acEnt
                                            CoordText = MathFunction.FuncFindCrossCenterPointPolyline(acPoly) 'находим центр полигона
                                        End If
                                        'восстанавливаем подпись
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем мультивыноску с тестом
                                        MLeader1.SetDatabaseDefaults()
                                        'MLeader1.PostMLeaderToDb(acCurDb)
                                        MLeader1.Layer = layerMLeader
                                        If IsNothing(idStyle) = False Then
                                            If idStyle.IsNull = False Then
                                                MLeader1.MLeaderStyle = idStyle
                                            End If
                                        End If
                                        Dim leaderIndex As Integer = MLeader1.AddLeader()
                                        Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                                        MLeader1.AddFirstVertex(leaderLineIndex, CoordText)
                                        MLeader1.AddLastVertex(leaderLineIndex, MaxPoint)
                                        MLeader1.ContentType = ContentType.MTextContent
                                        Dim acMText1 As MText = New MText()
                                        acMText1.SetDatabaseDefaults()
                                        Dim LABEL As String = ""
                                        If formRectoreLabel.RadioButton1.Checked = True Then
                                            acMText1.Contents = valConditionalNumber
                                            LABEL = "Обозначение участка"
                                        Else
                                            acMText1.Contents = valConditionalNumber & vbLf & strSqr
                                            LABEL = "Обозначение участка и площадь"
                                        End If
                                        If hText > 0 Then
                                            acMText1.Height = hText
                                        End If
                                        acMText1.TextStyleId = idTextStyle
                                        MLeader1.MText = acMText1
                                        acBlkTblRec.AppendEntity(MLeader1)
                                        tr.AddNewlyCreatedDBObject(MLeader1, True)
                                        'делаем MText идентификатор с помощью XRecord
                                        Dim ArrayOBJ1 As String(,) = Nothing
                                        ReDim Preserve ArrayOBJ1(2, 2)
                                        ArrayOBJ1(0, 0) = "CADASTRAL_OBJECT"
                                        ArrayOBJ1(1, 0) = "LABEL"
                                        ArrayOBJ1(2, 0) = LABEL
                                        ArrayOBJ1(0, 1) = "CADASTRAL_OBJECT"
                                        ArrayOBJ1(1, 1) = "Number"
                                        ArrayOBJ1(2, 1) = "1"
                                        ArrayOBJ1(0, 2) = "CADASTRAL_OBJECT"
                                        ArrayOBJ1(1, 2) = "IDRNG"
                                        ArrayOBJ1(2, 2) = LocalID
                                        If IsNothing(idTablePs) = False Then
                                            If idTablePs.IsNull = False Then
                                                Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(MLeader1, idTablePs)
                                                Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(MLeader1, ArrayOBJ1)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                tr.Commit()
            End If
        End Using
    End Sub

End Class
