Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Imports NetTopologySuite.Geometries
Public Class ListProjectParcels
    '=========================================================================================================
    'Отчет по образуемым земельным участкам (Текстовая часть проекта межевания)
    <CommandMethod("TPartProjectPM")>
    Public Sub TPartProjectPM()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim FormReport As FormListProjectParcels = New FormListProjectParcels
        FormReport.ShowDialog()
        If FormReport.boolShow = False Then
            Exit Sub
        End If
        Dim PutshFileTemplates As String = FormReport.ComboBox1.Tag & FormReport.ComboBox1.Text
        If IO.File.Exists(PutshFileTemplates) = False Then
            MsgBox("Файл-шаблон не найден. Макро прервано!!!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 1)
        ArrayTablesXML(0, 0) = "PROJECT_PARCEL"
        ArrayTablesXML(1, 0) = "Участок постоянного отвода"
        ArrayTablesXML(0, 1) = "TEMPORARY_PARCEL"
        ArrayTablesXML(1, 1) = "Участок временного отвода"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'округления вершин и площади
        Dim RoundZn As Integer = 2
        Dim AreaRoundZn As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '======================================================================================================
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            '=====================================================================================================
            'поворотные точки
            Dim GeoPointCoolection As ObjectIdCollection = New ObjectIdCollection
            Dim acSetBlk As SelectionSet = Nothing
            Dim ValType1(0) As TypedValue
            ValType1.SetValue(New TypedValue(DxfCode.BlockName, "PROJECT_POINT"), 0)
            Dim FilterParcel As SelectionFilter = New SelectionFilter(ValType1)
            Dim BlkSelect As PromptSelectionResult = ed.SelectAll(FilterParcel)
            If BlkSelect.Status = PromptStatus.OK Then
                acSetBlk = BlkSelect.Value
            End If
            If IsNothing(acSetBlk) = False Then
                If acSetBlk.Count > 0 Then
                    For Each id As ObjectId In acSetBlk.GetObjectIds
                        GeoPointCoolection.Add(id)
                    Next
                End If
            End If
            ValType1.SetValue(New TypedValue(DxfCode.BlockName, "PB*"), 0)
            FilterParcel = New SelectionFilter(ValType1)
            BlkSelect = ed.SelectAll(FilterParcel)
            If BlkSelect.Status = PromptStatus.OK Then
                acSetBlk = BlkSelect.Value
            End If
            If IsNothing(acSetBlk) = False Then
                If acSetBlk.Count > 0 Then
                    For Each id As ObjectId In acSetBlk.GetObjectIds
                        GeoPointCoolection.Add(id)
                    Next
                End If
            End If
            '=============================================================================================
            'забираем все мультивыноски
            Dim ArrayMLeader As String(,) = Nothing
            Dim NameStyleMLeader As String() = {"ППТ-Образуемые участки", "CADASTRAL_OBJECT"}
            Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, NameStyleMLeader)
            '==================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем все кадастровые участки на чертеже
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim CountArrayCadastralParcels As Integer = 0
            Dim ArrayCadastralParcels As String(,) = Nothing
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                        Dim nameOD As String = ""
                        Dim Kod As Integer = 0
                        Dim boolFindTablePSParcel As Boolean = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_PARCEL")
                        If boolFindTablePSParcel = True Then
                            nameOD = "LAND_CADASTRAL_PARCEL"
                            Kod = 0
                        End If
                        If boolFindTablePSParcel = False Then
                            boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_DIVISION")
                            If boolFindTablePSParcel = True Then
                                nameOD = "LAND_CADASTRAL_DIVISION"
                                Kod = 1
                            End If
                        End If
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayCadastralParcels(6, CountArrayCadastralParcels)
                            ArrayCadastralParcels(0, CountArrayCadastralParcels) = Kod
                            ArrayCadastralParcels(1, CountArrayCadastralParcels) = ent.Handle.ToString
                            ArrayCadastralParcels(2, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.X - 10
                            ArrayCadastralParcels(3, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.Y - 10
                            ArrayCadastralParcels(4, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.X + 10
                            ArrayCadastralParcels(5, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.Y + 10
                            ArrayCadastralParcels(6, CountArrayCadastralParcels) = nameOD
                            CountArrayCadastralParcels += 1
                            Continue For
                        End If
                    End If
                End If
            Next
            If IsArray(ArrayCadastralParcels) = True Then
                If ArrayCadastralParcels.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayCadastralParcels, 0)
                End If
            End If
            'выбор объектов
            Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков на всем пространстве чертежа (handle, nameOD)
            Dim countArrayAllPoligon As Integer = 0
            'выбор объектов вручную
            If FormReport.RadioButton3.Checked = True Then
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                            ReDim Preserve ArrayAllPoligon(31, countArrayAllPoligon)
                                            ArrayAllPoligon(0, countArrayAllPoligon) = "0"
                                            ArrayAllPoligon(1, countArrayAllPoligon) = acEnt.Handle.ToString
                                            ArrayAllPoligon(31, countArrayAllPoligon) = ArrayNameTableOD(i)
                                            countArrayAllPoligon += 1
                                        Next i
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось найти участки!!!")
                    Exit Sub
                End If
                '===============================================================================================
                'выбор объектов автоматически
            ElseIf FormReport.RadioButton4.Checked = True Then
                For Each id As ObjectId In btr
                    If id.IsNull = False Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                        If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            Dim ArrayNameTableOD As String() = Nothing
                            Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                            If IsArray(ArrayNameTableOD) = True Then
                                ReDim Preserve ArrayAllPoligon(31, countArrayAllPoligon)
                                ArrayAllPoligon(0, countArrayAllPoligon) = "0"
                                ArrayAllPoligon(1, countArrayAllPoligon) = acEnt.Handle.ToString
                                ArrayAllPoligon(31, countArrayAllPoligon) = ArrayNameTableOD(0)
                                countArrayAllPoligon += 1
                            End If
                        End If
                    End If
                Next
                '===============================================================================================
                'выбор объектов внутри контура
            ElseIf FormReport.RadioButton6.Checked = True Then
                Dim CounterPline As Polyline = New Polyline
                Dim EntSelectCounter As PromptEntityResult = acDoc.Editor.GetEntity("Выберите полилинию: ")
                If EntSelectCounter.Status = PromptStatus.OK Then
                    Dim EntCounter As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(EntSelectCounter.ObjectId, OpenMode.ForRead)
                    If TypeOf EntCounter Is Polyline Then
                        Dim CounterPolyline As Polyline = EntCounter
                        For Each id As ObjectId In btr
                            If id.IsNull = False Then
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        Dim Pline As Polyline = acEnt
                                        Dim CenterPoint3d As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(Pline)
                                        Dim centerPoint As Point2d = CenterPoint3d.Convert2d(New Plane)
                                        Dim boolCr As Boolean = MathFunction.FuncCrossPointPolyline(CounterPline, centerPoint)
                                        If boolCr = True Then
                                            ReDim Preserve ArrayAllPoligon(31, countArrayAllPoligon)
                                            ArrayAllPoligon(0, countArrayAllPoligon) = "0"
                                            ArrayAllPoligon(1, countArrayAllPoligon) = acEnt.Handle.ToString
                                            ArrayAllPoligon(31, countArrayAllPoligon) = ArrayNameTableOD(0)
                                            countArrayAllPoligon += 1
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            '=============================================================================================================================================
            If IsArray(ArrayAllPoligon) = False Then
                MsgBox("Участки не найдены. Макро прервано!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем обозначения участков
            '(0-короткий номер участка, 1-handle,2- полное обозначение, 3-кадастровый номер, 4-Адрес, 5-категория планируемая, 6-разр.исп
            For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                Dim PlineHandle As String = ArrayAllPoligon(1, i)
                Dim NameOD As String = ArrayTablesXML(0, 0)
                Dim idProjectParcel As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                If IsNothing(idProjectParcel) = False Then
                    If idProjectParcel.IsNull = False Then
                        Dim PlineObj As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
                        'достаем полилинию
                        Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idProjectParcel, OpenMode.ForRead)
                        If TypeOf Ent Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                            PlineObj = Ent
                        Else
                            Continue For
                        End If
                        'общие параметры
                        Dim arrayData As String(,) = Nothing 'массив с расширенными данными
                        Dim LocalID As String = ""
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ищем глобальный и локальный id
                        Dim boolPSecord As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(PlineObj, arrayData, ArrayTablesXML)
                        If IsArray(arrayData) = True Then
                            For j As Integer = 0 To arrayData.GetUpperBound(1)
                                Dim NField As String = arrayData(1, j)
                                Dim Zn As String = arrayData(2, j)
                                If IsNothing(Zn) = True Then
                                    Zn = ""
                                Else
                                    Zn = Zn.Trim()
                                End If
                                If NField Like "ConditionalNumber" Then
                                    ArrayAllPoligon(2, i) = Zn
                                ElseIf NField Like "Cad_number" Then
                                    ArrayAllPoligon(3, i) = Zn
                                ElseIf NField Like "CadastralNumber" Then
                                    ArrayAllPoligon(3, i) = Zn
                                ElseIf NField Like "Adress" Then
                                    ArrayAllPoligon(4, i) = Zn
                                ElseIf NField Like "Category" Then
                                    ArrayAllPoligon(5, i) = Zn
                                ElseIf NField Like "Utilization" Then
                                    ArrayAllPoligon(6, i) = Zn
                                ElseIf NField Like "Right" Then
                                    ArrayAllPoligon(7, i) = Zn
                                ElseIf NField Like "Patronymic" Then
                                    ArrayAllPoligon(8, i) = Zn
                                ElseIf NField Like "Encumbrance" Then
                                    ArrayAllPoligon(9, i) = Zn
                                ElseIf NField Like "Method_create" Then
                                    ArrayAllPoligon(10, i) = Zn
                                ElseIf NField Like "Confiscation" Then
                                    ArrayAllPoligon(11, i) = Zn
                                ElseIf NField Like "Public_land" Then
                                    ArrayAllPoligon(12, i) = Zn
                                ElseIf NField Like "GLOBALID" Then
                                    ArrayAllPoligon(13, i) = Zn
                                ElseIf NField Like "LOCALID" Then
                                    ArrayAllPoligon(14, i) = Zn
                                    LocalID = Zn
                                ElseIf NField Like "NOTE" Then
                                    ArrayAllPoligon(15, i) = Zn
                                ElseIf NField Like "Assignation" Then
                                    ArrayAllPoligon(16, i) = Zn
                                ElseIf NField Like "Grounds_establishing" Then
                                    ArrayAllPoligon(17, i) = Zn
                                ElseIf NField Like "Limitation" Then
                                    ArrayAllPoligon(18, i) = Zn
                                End If
                            Next j
                        Else
                            Continue For
                        End If
                        '====================================================================
                        'ищем обозначение участка по его локальному id
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ИЩЕМ ОБОЗНАЧЕНИЕ УЧАСТКА (ЛИБО ТЕКСТ ЛИБО МУЛЬТИВЫНОска)
                        Dim ConditionParcel = ArrayAllPoligon(2, i)
                        Dim CenterParcel As Point3d = New Point3d(-1, -1, -1)
                        If LocalID.Trim.Length > 0 Then
                            'ищем мультивыноску
                            Dim acMleader As MLeader = Nothing
                            acMleader = FuncCadastr.FuncFindMLeaderByLocalId(LocalID)
                            If IsNothing(acMleader) = False Then
                                Dim PosFirstVertex As Point3d = acMleader.GetFirstVertex(0)
                                Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(PlineObj, PosFirstVertex.Convert2d(New Plane))
                                If PointCrosPoligon = True Then
                                    ConditionParcel = acMleader.MText.Contents
                                Else
                                    If IsArray(ArrayMLeader) = True Then
                                        For l1 As Integer = 0 To ArrayMLeader.GetUpperBound(1)
                                            Dim xVert As Double = Val(ArrayMLeader(1, l1))
                                            Dim yVert As Double = Val(ArrayMLeader(2, l1))
                                            Dim ptVert As Point2d = New Point2d(xVert, yVert)
                                            PointCrosPoligon = MathFunction.FuncCrossPointPolyline(PlineObj, ptVert)
                                            If PointCrosPoligon = True Then
                                                ConditionParcel = FuncCadastr.FuncFindNumberCadastralObject(ArrayMLeader(0, l1))
                                                CenterParcel = PosFirstVertex
                                                Exit For
                                            End If
                                        Next l1
                                    End If
                                End If
                            Else
                                If IsArray(ArrayMLeader) = True Then
                                    For l1 As Integer = 0 To ArrayMLeader.GetUpperBound(1)
                                        Dim xVert As Double = Val(ArrayMLeader(1, l1))
                                        Dim yVert As Double = Val(ArrayMLeader(2, l1))
                                        Dim ptVert As Point2d = New Point2d(xVert, yVert)
                                        Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(PlineObj, ptVert)
                                        If PointCrosPoligon = True Then
                                            ConditionParcel = FuncCadastr.FuncFindNumberCadastralObject(ArrayMLeader(0, l1))
                                            CenterParcel = New Point3d(xVert, yVert, 0)
                                            Exit For
                                        End If
                                    Next l1
                                End If
                            End If
                        End If
                        '===============================================================================
                        'отделяем номер участка от его префикса
                        ConditionParcel = FuncCadastr.FuncFindNumberCadastralObject(ConditionParcel)
                        Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionParcel)
                        ArrayAllPoligon(0, i) = numbParcel
                        ConditionParcel = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionParcel)
                        ArrayAllPoligon(2, i) = ConditionParcel
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ищем исходный участок из которого был образован новый
                        Dim cad_number As String = ArrayAllPoligon(3, i) 'кадастровый номер исходного зу
                        Dim boolcad_number = False
                        If IsNothing(cad_number) = False Then
                            Dim ArrayCadNumber As String() = cad_number.Split(":")
                            If ArrayCadNumber.Length > 3 Then
                                If IsNumeric(ArrayCadNumber(3)) Then
                                    boolcad_number = True
                                End If
                            End If
                        End If
                        If CenterParcel.X = -1 And CenterParcel.Y = -1 And CenterParcel.Z = -1 Then
                            CenterParcel = MathFunction.FuncFindCrossCenterPointPolyline(PlineObj)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'ищем информацию об участке кадастровом
                        If IsArray(ArrayCadastralParcels) = True Then
                            For k1 As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                                Dim minX As Double = ArrayCadastralParcels(2, k1)
                                Dim minY As Double = ArrayCadastralParcels(3, k1)
                                Dim maxX As Double = ArrayCadastralParcels(4, k1)
                                Dim maxY As Double = ArrayCadastralParcels(5, k1)
                                If CenterParcel.X > minX And CenterParcel.Y > minY Then
                                    If CenterParcel.X < maxX And CenterParcel.Y < maxY Then
                                        Dim idEnt As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(1, k1))
                                        If IsNothing(idEnt) = False Then
                                            If idEnt.IsNull = False Then
                                                Dim acEntParcel As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idEnt, OpenMode.ForRead)
                                                If boolcad_number = True Then
                                                    Dim cadastrnumb As String = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Cad_number")
                                                    If Not (cad_number Like cadastrnumb) Then
                                                        Continue For
                                                    End If
                                                Else
                                                    If TypeOf acEntParcel Is Polyline Then
                                                        Dim PlineCadParcel As Polyline = acEntParcel
                                                        Dim boolcross As Boolean = MathFunction.FuncCrossPointPolyline(PlineCadParcel, CenterParcel.Convert2d(New Plane))
                                                        If boolcross = False Then
                                                            Continue For
                                                        End If
                                                    End If
                                                End If
                                                '====================================================================================================
                                                'записываем данные с исходного земельного участка
                                                ArrayAllPoligon(30, i) = ArrayCadastralParcels(0, k1)
                                                ArrayAllPoligon(29, i) = acEntParcel.Handle.ToString
                                                If ArrayCadastralParcels(0, k1) = 0 Then
                                                    ArrayAllPoligon(19, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Cad_number")
                                                    Dim ParcelArea As Double = Val(FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Area"))
                                                    ArrayAllPoligon(20, i) = ParcelArea
                                                    ArrayAllPoligon(21, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Adress")
                                                    ArrayAllPoligon(22, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Category")
                                                    ArrayAllPoligon(23, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Utilization")
                                                    If ArrayAllPoligon(23, i).Trim.Length = 0 Then
                                                        ArrayAllPoligon(23, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "UtilizationByDoc")
                                                    End If
                                                    ArrayAllPoligon(24, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Right")
                                                    ArrayAllPoligon(25, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Owner")
                                                    ArrayAllPoligon(26, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Encumbrance")
                                                    ArrayAllPoligon(28, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "NOTE")
                                                Else
                                                    ArrayAllPoligon(19, i) = FuncPropertySet.FuncReadValueSetProperties(acEntParcel, "Cadastral_number")
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            Next k1
                        End If
                    End If
                End If
            Next i
            '===================================================================================================
            If ArrayAllPoligon.GetUpperBound(1) > 0 Then
                'сортируем участки
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
            End If
            '========================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'тело программы
            'последовательно достаем участки из массива
            If IsArray(ArrayAllPoligon) = True Then
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                If FormReport.RadioButton1.Checked = True Then
                    Dim WApp As Word.Application = Nothing
                    Dim WDoc As Word.Document = Nothing
                    Dim boolWDVisibl As Boolean = False
                    Try
                        Dim rng As Random = New Random
                        Dim Numbrng As Integer = rng.Next
                        If FormReport.ComboBox1.Text.Trim Like "ПМТ_Текстовая часть.docx" Then
                            Dim DestDoc As String = TempDir & "ПМТ_Текстовая часть_" & Numbrng & ".docx"
                            Try
                                If File.Exists(PutshFileTemplates) = True Then
                                    File.Copy(PutshFileTemplates, DestDoc, True)
                                End If
                            Catch ex As UnauthorizedAccessException
                            Catch ex As ArgumentException
                            Catch ex As PathTooLongException
                            Catch ex As DirectoryNotFoundException
                            Catch ex As FileNotFoundException
                            Catch ex As IOException
                            Catch ex As NotSupportedException
                            End Try
                            If File.Exists(DestDoc) = True Then
                                WApp = CreateObject("Word.Application")
                                WDoc = WApp.Documents.Open(DestDoc)
                            Else
                                MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                                WApp.Quit()
                                WApp = Nothing
                                Exit Sub
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'начинаем перебирать участки
                            WApp.Visible = True
                            Dim arrayGlobalID As String() = Nothing
                            Dim countarrayGlobalID As Integer = 0
                            For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                                Dim ProjectAreaParcel As Double = 0
                                Dim CountCounter As Integer = 0
                                Dim StrNumberPoint As String = ""
                                Dim GlobalID As String = ArrayAllPoligon(13, i)
                                Dim LocalID As String = ArrayAllPoligon(14, i)
                                Dim hgParcel As String = ArrayAllPoligon(1, i)
                                Dim parcelCadasterArea As String = ArrayAllPoligon(20, i)
                                If GlobalID.Trim.Length > 2 Then
                                    If IsArray(arrayGlobalID) = True Then
                                        For j As Integer = 0 To arrayGlobalID.Length - 1
                                            If GlobalID Like arrayGlobalID(j) Then
                                                Continue For
                                            End If
                                        Next j
                                    End If
                                    ReDim arrayGlobalID(countarrayGlobalID)
                                    arrayGlobalID(countarrayGlobalID) = GlobalID
                                    countarrayGlobalID += 1
                                    'ищем входящие номера
                                    Dim ArrayCounterTemp As String(,) = Nothing
                                    Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(GlobalID.Trim, ArrayCounterTemp, ArrayTablesXML(0, 0))
                                    If IsArray(ArrayCounterTemp) = True Then
                                        'перебор по входящим контурам
                                        For j As Integer = 0 To ArrayCounterTemp.GetUpperBound(1)
                                            Dim HandlePLine As String = ArrayCounterTemp(0, j) 'нангл контура
                                            Dim idLoc As String = ArrayCounterTemp(1, j) 'локальный ид
                                            'получаем доступ к полилинии для извлечения координат
                                            Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                            If IsNothing(idCounter) = True Then
                                                Continue For
                                            End If
                                            If idCounter.IsNull = True Then
                                                Continue For
                                            End If
                                            Dim PlineObj As Polyline = tr.GetObject(idCounter, OpenMode.ForRead)
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'вычисляем площадь контура
                                            Dim Pt3dColliction As Point3dCollection = New Point3dCollection()
                                            For k = 0 To PlineObj.NumberOfVertices - 1
                                                If k = PlineObj.NumberOfVertices - 1 Then
                                                    If PlineObj.GetPoint3dAt(0) = PlineObj.GetPoint3dAt(k) Then
                                                        Exit For
                                                    End If
                                                End If
                                                Pt3dColliction.Add(PlineObj.GetPoint3dAt(k))
                                            Next
                                            Dim SQRParcel As Double = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                            ProjectAreaParcel = ProjectAreaParcel + SQRParcel
                                            CountCounter += 1
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            Dim TempNumber As String = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                            If TempNumber.Trim.Length > 0 Then
                                                If StrNumberPoint.Trim.Length = 0 Then
                                                    StrNumberPoint = TempNumber
                                                Else
                                                    StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                End If
                                            End If
                                            'ищем островки
                                            Dim ArrayOstrov1 As String(,) = Nothing
                                            Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLoc, ArrayOstrov1)
                                            If IsArray(ArrayOstrov1) = True Then
                                                For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                                    'получаем доступ к полилинии для извлечения координат
                                                    HandlePLine = ArrayOstrov1(1, k)
                                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                                    If IsNothing(idOstrov) = True Then
                                                        Continue For
                                                    End If
                                                    If idOstrov.IsNull = True Then
                                                        Continue For
                                                    End If
                                                    PlineObj = tr.GetObject(idOstrov, OpenMode.ForRead)
                                                    Pt3dColliction = New Point3dCollection()
                                                    For k1 = 0 To PlineObj.NumberOfVertices - 1
                                                        If k1 = PlineObj.NumberOfVertices - 1 Then
                                                            If PlineObj.GetPoint3dAt(0) = PlineObj.GetPoint3dAt(k1) Then
                                                                Exit For
                                                            End If
                                                        End If
                                                        Pt3dColliction.Add(PlineObj.GetPoint3dAt(k1))
                                                    Next
                                                    Dim SQROstrov As Double = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                                    ProjectAreaParcel = ProjectAreaParcel - SQROstrov
                                                    TempNumber = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                                    If TempNumber.Trim.Length > 0 Then
                                                        If StrNumberPoint.Trim.Length = 0 Then
                                                            StrNumberPoint = TempNumber
                                                        Else
                                                            StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                        Next j
                                    End If
                                Else
                                    CountCounter = 1
                                    Dim idParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgParcel)
                                    If IsNothing(idParcel) = True Then
                                        Continue For
                                    End If
                                    If idParcel.IsNull = True Then
                                        Continue For
                                    End If
                                    Dim PlineObj As Polyline = tr.GetObject(idParcel, OpenMode.ForRead)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'вычисляем площадь контура
                                    Dim Pt3dColliction As Point3dCollection = New Point3dCollection()
                                    For k = 0 To PlineObj.NumberOfVertices - 1
                                        If k = PlineObj.NumberOfVertices - 1 Then
                                            If PlineObj.GetPoint3dAt(0) = PlineObj.GetPoint3dAt(k) Then
                                                Exit For
                                            End If
                                        End If
                                        Pt3dColliction.Add(PlineObj.GetPoint3dAt(k))
                                    Next
                                    Dim SQRParcel As Double = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                    ProjectAreaParcel = ProjectAreaParcel + SQRParcel
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim TempNumber As String = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                    If TempNumber.Trim.Length > 0 Then
                                        If StrNumberPoint.Trim.Length = 0 Then
                                            StrNumberPoint = TempNumber
                                        Else
                                            StrNumberPoint = StrNumberPoint & "," & TempNumber
                                        End If
                                    End If
                                    'ищем островки
                                    Dim ArrayOstrov1 As String(,) = Nothing
                                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov1)
                                    If IsArray(ArrayOstrov1) = True Then
                                        For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                            'получаем доступ к полилинии для извлечения координат
                                            Dim HandlePLine As String = ArrayOstrov1(1, k)
                                            Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                            If IsNothing(idOstrov) = True Then
                                                Continue For
                                            End If
                                            If idOstrov.IsNull = True Then
                                                Continue For
                                            End If
                                            PlineObj = tr.GetObject(idOstrov, OpenMode.ForRead)
                                            Pt3dColliction = New Point3dCollection()
                                            For k1 = 0 To PlineObj.NumberOfVertices - 1
                                                If k1 = PlineObj.NumberOfVertices - 1 Then
                                                    If PlineObj.GetPoint3dAt(0) = PlineObj.GetPoint3dAt(k1) Then
                                                        Exit For
                                                    End If
                                                End If
                                                Pt3dColliction.Add(PlineObj.GetPoint3dAt(k1))
                                            Next
                                            Dim SQROstrov As Double = MathFunction.FuncSQRTPolyline(Pt3dColliction, RoundZn, AreaRoundZn)
                                            ProjectAreaParcel = ProjectAreaParcel - SQROstrov
                                            TempNumber = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                            If TempNumber.Trim.Length > 0 Then
                                                If StrNumberPoint.Trim.Length = 0 Then
                                                    StrNumberPoint = TempNumber
                                                Else
                                                    StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                                If IsNothing(StrNumberPoint) = False Then
                                    If StrNumberPoint.Trim.Length > 0 Then
                                        StrNumberPoint = FuncFormatZN.FuncConvertStrNumberPoint(StrNumberPoint)
                                    End If
                                End If
                                '==============================================================================================================================
                                'флаги для записи ЗУ в определенную таблицу
                                Dim boolObrazovanieZU As Boolean = False
                                Dim boolObrazovanieRazdIshZU As Boolean = False
                                Dim boolObrazovanieRazdZU As Boolean = False
                                Dim boolJoinZU As Boolean = False
                                Dim boolNo As Boolean = False

                                Dim boolIzytieZU As Boolean = False
                                Dim boolRezervZU As Boolean = False

                                Dim boolPunlicLandZU As Boolean = False

                                Dim boolServitut As Boolean = False

                                Dim boolForestZu As Boolean = False
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'способ образования
                                Dim strObrazovanieZu As String = ArrayAllPoligon(10, i)
                                If IsNothing(strObrazovanieZu) = False Then
                                    If strObrazovanieZu.IndexOf("осударствен") > -1 Then
                                        boolObrazovanieZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                    ElseIf strObrazovanieZu.IndexOf("муниципал") > -1 Then
                                        boolObrazovanieZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                    ElseIf Mid(strObrazovanieZu, 1) Like "1" Then
                                        boolObrazovanieZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                    ElseIf strObrazovanieZu.IndexOf("исходн") > -1 Then
                                        boolObrazovanieRazdIshZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка с сохранением исходного в измененных границах"
                                    ElseIf Mid(strObrazovanieZu, 1) Like "2" Then
                                        boolObrazovanieRazdIshZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка с сохранением исходного в измененных границах"
                                    ElseIf strObrazovanieZu.IndexOf("раздел") > -1 Then
                                        boolObrazovanieRazdZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка"
                                    ElseIf Mid(strObrazovanieZu, 1) Like "3" Then
                                        boolObrazovanieRazdZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка"
                                    ElseIf strObrazovanieZu.IndexOf("объединен") > -1 Then
                                        boolJoinZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем объединения земельных участков"
                                    ElseIf Mid(strObrazovanieZu, 1) Like "4" Then
                                        boolJoinZU = True
                                        ArrayAllPoligon(10, i) = "Образование земельного участка путем объединения земельных участков"
                                    ElseIf Mid(strObrazovanieZu, 1) Like "Нет" Then
                                        boolNo = True
                                    End If
                                End If

                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'способ изъятия
                                Dim strRezervZu As String = ArrayAllPoligon(11, i)
                                If IsNothing(strRezervZu) = False Then
                                    If strRezervZu.IndexOf("зъят") > -1 Then
                                        boolIzytieZU = True
                                    ElseIf strRezervZu.IndexOf("езервиров") > -1 Then
                                        boolRezervZU = True
                                    ElseIf Mid(strRezervZu, 1) Like "2" Then
                                        boolRezervZU = True
                                    ElseIf Mid(strRezervZu, 1) Like "1" Then
                                        boolIzytieZU = True
                                    End If
                                End If

                                Dim strPublicZu As String = ArrayAllPoligon(6, i)
                                If IsNothing(strPublicZu) = False Then
                                    If strPublicZu.IndexOf("общего") > -1 Then
                                        boolPunlicLandZU = True
                                    End If
                                End If

                                If ArrayAllPoligon(31, i) Like "TEMPORARY_PARCEL" Then
                                    boolServitut = True
                                End If

                                If IsNothing(ArrayAllPoligon(22, i)) = False Then
                                    If ArrayAllPoligon(22, i) Like "Земли лесного фонда" Then
                                        boolForestZu = True
                                    End If
                                End If
                                '=============================================================================================================
                                'заполняем таблицы
                                For Each tabl As Word.Table In WDoc.Tables
                                    If tabl.Title Like "Образование ЗУ из государственной собственности" And boolObrazovanieZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'номера характерных точек
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = StrNumberPoint
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Kategory As String = ArrayAllPoligon(5, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = Kategory 'категория земель
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" 'категория земель
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 5).Range.Text = "-" 'категория земель
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utiliza As String = ArrayAllPoligon(6, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 6).Range.Text = Utiliza  'планируемый вид разрешенного использования
                                            Else
                                                tabl.Cell(countTableword, 6).Range.Text = "-" 'планируемый вид разрешенного использования
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 6).Range.Text = "-" 'планируемый вид разрешенного использования
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 7).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Способы образования земельных участков
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim SposobObr As String = ArrayAllPoligon(10, i)
                                        tabl.Cell(countTableword, 8).Range.Text = SposobObr
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(15, i)
                                        tabl.Cell(countTableword, 9).Range.Text = Note
                                        boolObrazovanieZU = False
                                        Continue For
                                    End If
                                    If tabl.Title Like "Образование ЗУ путем раздела с сохранением исходного" And boolObrazovanieRazdIshZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'номера характерных точек
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = StrNumberPoint
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Кадастровые номера земельных участков, из которых образуются земельные участки
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim cad_number As String = ArrayAllPoligon(3, i)
                                        If IsNothing(cad_number) = False Then
                                            If cad_number.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                            Else
                                                tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель существующая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Kategory As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = Kategory 'категория земель сушествующая
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель сушествующая
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель сушествующая
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель планируемая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Kategory = ArrayAllPoligon(5, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 8).Range.Text = Kategory 'категория земель проектная
                                            Else
                                                tabl.Cell(countTableword, 8).Range.Text = "-" 'категория земель проектная
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 8).Range.Text = "-" 'категория земель проектная
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование существующее
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utiliza As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 9).Range.Text = Utiliza  'планируемый вид разрешенного использования существующее
                                            Else
                                                tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование планируемое
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Utiliza = ArrayAllPoligon(6, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 10).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                            Else
                                                tabl.Cell(countTableword, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Способы образования земельных участков
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim SposobObr As String = ArrayAllPoligon(10, i)
                                        tabl.Cell(countTableword, 11).Range.Text = SposobObr
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim dSQTR As Integer = Math.Round(parcelCadasterArea - ProjectAreaParcel, 0)
                                        Dim Note As String = ArrayAllPoligon(15, i) & "Площадь земельного участка с кадастровым номером " & cad_number & ", сохраняемого в измененных границах – " & dSQTR & "кв.м."
                                        tabl.Cell(countTableword, 12).Range.Text = Note
                                        boolObrazovanieRazdIshZU = False
                                        Continue For
                                    End If
                                    If tabl.Title Like "Образование ЗУ путем раздела" And boolObrazovanieRazdZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'номера характерных точек
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = StrNumberPoint
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Кадастровые номера земельных участков, из которых образуются земельные участки
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim cad_number As String = ArrayAllPoligon(3, i)
                                        If IsNothing(cad_number) = False Then
                                            If cad_number.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                            Else
                                                tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель существующая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Kategory As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = Kategory 'категория земель сушествующая
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель сушествующая
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель сушествующая
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель планируемая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Kategory = ArrayAllPoligon(5, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 8).Range.Text = Kategory 'категория земель проектная
                                            Else
                                                tabl.Cell(countTableword, 8).Range.Text = "-" 'категория земель проектная
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 8).Range.Text = "-" 'категория земель проектная
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование существующее
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utiliza As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 9).Range.Text = Utiliza  'планируемый вид разрешенного использования существующее
                                            Else
                                                tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование планируемое
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Utiliza = ArrayAllPoligon(6, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 10).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                            Else
                                                tabl.Cell(countTableword, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Способы образования земельных участков
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim SposobObr As String = ArrayAllPoligon(10, i)
                                        tabl.Cell(countTableword, 11).Range.Text = SposobObr
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim dSQTR As Integer = Math.Round(parcelCadasterArea - ProjectAreaParcel, 0)
                                        Dim Note As String = ArrayAllPoligon(15, i) & "Площадь земельного участка с кадастровым номером " & cad_number & ", сохраняемого в результате раздела – " & dSQTR & "кв.м."
                                        tabl.Cell(countTableword, 12).Range.Text = Note
                                        boolObrazovanieRazdZU = False
                                        Continue For
                                    End If
                                    If tabl.Title Like "Изъятие или резервирование земельных участков" And boolNo = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'номера характерных точек
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = StrNumberPoint
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 4).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 4).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 4).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 5).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель существующая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Kategory As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 6).Range.Text = Kategory 'категория земель сушествующая
                                            Else
                                                tabl.Cell(countTableword, 6).Range.Text = "-" 'категория земель сушествующая
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 6).Range.Text = "-" 'категория земель сушествующая
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель планируемая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Kategory = ArrayAllPoligon(5, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = Kategory 'категория земель проектная
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель проектная
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель проектная
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование существующее
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utiliza As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 8).Range.Text = Utiliza  'планируемый вид разрешенного использования существующее
                                            Else
                                                tabl.Cell(countTableword, 8).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 8).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование планируемое
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Utiliza = ArrayAllPoligon(6, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 9).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                            Else
                                                tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 9).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(15, i)
                                        tabl.Cell(countTableword, 10).Range.Text = Note
                                        boolNo = False
                                        Continue For
                                    End If
                                    If tabl.Title Like "Образование ЗУ путем объединения" And boolJoinZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'номера характерных точек
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = StrNumberPoint
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Кадастровые или условные номера земельных участков, подлежащих объединению
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Dim cad_number As String = ArrayAllPoligon(3, i)
                                        'If IsNothing(cad_number) = False Then
                                        '    If cad_number.Trim.Length > 0 Then
                                        '        tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                        '    Else
                                        '        tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        '    End If
                                        'Else
                                        '    tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        'End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель планируемая
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Kategory As String = ArrayAllPoligon(5, i)
                                        If IsNothing(Kategory) = False Then
                                            If Kategory.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = Kategory 'категория земель проектная
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель проектная
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 7).Range.Text = "-" 'категория земель проектная
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование планируемое
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utiliza As String = ArrayAllPoligon(6, i)
                                        If IsNothing(Utiliza) = False Then
                                            If Utiliza.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 8).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                            Else
                                                tabl.Cell(countTableword, 8).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 8).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Способы образования земельных участков
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim SposobObr As String = ArrayAllPoligon(10, i)
                                        tabl.Cell(countTableword, 9).Range.Text = SposobObr
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(15, i)
                                        tabl.Cell(countTableword, 10).Range.Text = Note
                                        boolJoinZU = False
                                        Continue For
                                    End If
                                    If tabl.Title Like "Перечень изымаемых ЗУ" And boolIzytieZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 3).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Category As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Category) = True Then
                                            Category = ArrayAllPoligon(5, i)
                                        ElseIf Category.Trim.Length = 0 Then
                                            Category = ArrayAllPoligon(5, i)
                                        End If

                                        If IsNothing(Category) = False Then
                                            If Category.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = Category 'категория земель исходного участка
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utilization As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utilization) = True Then
                                            Utilization = ArrayAllPoligon(6, i)
                                        ElseIf Utilization.Trim.Length = 0 Then
                                            Utilization = ArrayAllPoligon(6, i)
                                        End If

                                        If IsNothing(Utilization) = False Then
                                            If Utilization.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 6).Range.Text = Utilization 'разрешенное использование
                                            Else
                                                tabl.Cell(countTableword, 6).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(11, i)
                                        tabl.Cell(countTableword, 8).Range.Text = Note
                                        boolIzytieZU = False
                                        Continue For
                                    End If

                                    If tabl.Title Like "Перечень резервируемых ЗУ" And boolRezervZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 3).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Category As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Category) = True Then
                                            Category = ArrayAllPoligon(5, i)
                                        ElseIf Category.Trim.Length = 0 Then
                                            Category = ArrayAllPoligon(5, i)
                                        End If

                                        If IsNothing(Category) = False Then
                                            If Category.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = Category 'категория земель исходного участка
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utilization As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utilization) = True Then
                                            Utilization = ArrayAllPoligon(6, i)
                                        ElseIf Utilization.Trim.Length = 0 Then
                                            Utilization = ArrayAllPoligon(6, i)
                                        End If

                                        If IsNothing(Utilization) = False Then
                                            If Utilization.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 6).Range.Text = Utilization 'разрешенное использование
                                            Else
                                                tabl.Cell(countTableword, 6).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(11, i)
                                        tabl.Cell(countTableword, 8).Range.Text = Note
                                        boolRezervZU = False
                                        Continue For
                                    End If

                                    If tabl.Title Like "ЗУ общего пользования" And boolPunlicLandZU = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 3).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 3).Range.Text = "-" 'адрес
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(15, i)
                                        tabl.Cell(countTableword, 6).Range.Text = Note
                                        boolPunlicLandZU = False
                                        Continue For
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'таблица с лесными ЗУ
                                    If (tabl.Title Like "Перечень и сведения о лесных земельных участках") And boolForestZu = True And boolServitut = False Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Кадастровые номера земельных участков, из которых образуются земельные участки
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim cad_number As String = ArrayAllPoligon(3, i)
                                        If IsNothing(cad_number) = False Then
                                            If cad_number.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                            Else
                                                tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                    End If

                                    If tabl.Title Like "Публичные сервитуты_сервитуты" And boolServitut = True Then
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'записываем номер участка
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                        If CountCounter > 1 Then
                                            ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                        End If
                                        If IsNothing(ConditionParcel) = False Then
                                            If ConditionParcel.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 2).Range.Text = ConditionParcel 'условный номер участка
                                            Else
                                                tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        tabl.Cell(countTableword, 3).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'Кадастровые номера земельных участков, из которых образуются земельные участки
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim cad_number As String = ArrayAllPoligon(3, i)
                                        If IsNothing(cad_number) = False Then
                                            If cad_number.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                            Else
                                                tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                            End If
                                        Else
                                            tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'местоположение
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim AddressP As String = ArrayAllPoligon(4, i)
                                        If IsNothing(AddressP) = False Then
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                            Else
                                                AddressP = ArrayAllPoligon(21, i)
                                                If IsNothing(AddressP) = True Then AddressP = ""
                                                If AddressP.Trim.Length > 0 Then
                                                    tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                                Else
                                                    tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                                End If
                                            End If
                                        Else
                                            AddressP = ArrayAllPoligon(21, i)
                                            If IsNothing(AddressP) = True Then AddressP = ""
                                            If AddressP.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 5).Range.Text = AddressP 'адрес
                                            Else
                                                tabl.Cell(countTableword, 5).Range.Text = "-" 'адрес
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'категория земель
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Category As String = ArrayAllPoligon(22, i)
                                        If IsNothing(Category) = False Then
                                            If Category.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 6).Range.Text = Category 'категория земель исходного участка
                                            Else
                                                tabl.Cell(countTableword, 6).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'разрешенное использование
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Utilization As String = ArrayAllPoligon(23, i)
                                        If IsNothing(Utilization) = False Then
                                            If Utilization.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 7).Range.Text = Utilization 'разрешенное использование
                                            Else
                                                tabl.Cell(countTableword, 7).Range.Text = "-" '
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'примечание
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim Note As String = ArrayAllPoligon(15, i)
                                        If IsNothing(Note) = False Then
                                            If Note.Trim.Length > 0 Then
                                                tabl.Cell(countTableword, 9).Range.Text = Note 'Примечане
                                            Else
                                                tabl.Cell(countTableword, 9).Range.Text = "-" '
                                            End If
                                        End If
                                        boolServitut = False
                                        Continue For
                                    End If
                                Next
                            Next
                            WDoc = Nothing
                            WApp.Quit()
                            WApp = Nothing
                        End If
                    Catch
                        MsgBox("Не удалось записать данные в выбранную таблицу!!!")
                        WDoc = Nothing
                        WApp = Nothing
                        Exit Sub
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Excel документ
                ElseIf FormReport.RadioButton2.Checked = True Then
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
                    Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
                    Dim CellRng As Excel.Range = Nothing 'массив ячеек
                    Dim boolWDVisibl As Boolean = False
                    Try
                        Dim rng As Random = New Random
                        Dim Numbrng As Integer = rng.Next
                        If FormReport.ComboBox1.Text.Trim Like "ПМТ_Текстовая часть.xlsx" Then
                            Dim DestDoc As String = TempDir & "ПМТ_Текстовая часть_" & Numbrng & ".xlsx"
                            Try
                                If File.Exists(PutshFileTemplates) = True Then
                                    File.Copy(PutshFileTemplates, DestDoc, True)
                                End If
                            Catch ex As UnauthorizedAccessException
                            Catch ex As ArgumentException
                            Catch ex As PathTooLongException
                            Catch ex As DirectoryNotFoundException
                            Catch ex As FileNotFoundException
                            Catch ex As IOException
                            Catch ex As NotSupportedException
                            End Try
                            If File.Exists(DestDoc) = True Then
                                xlApp = CreateObject("Excel.Application")
                                xlBook = xlApp.Workbooks.Add()
                            Else
                                MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                                xlApp.Quit()
                                xlApp = Nothing
                                Exit Sub
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'начинаем перебирать участки
                            xlApp.Visible = True
                            Dim arrayGlobalID As String() = Nothing
                            Dim countarrayGlobalID As Integer = 0
                            For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                                Dim ProjectAreaParcel As Double = 0
                                Dim CountCounter As Integer = 0
                                Dim StrNumberPoint As String = ""
                                Dim GlobalID As String = ArrayAllPoligon(13, i)
                                Dim LocalID As String = ArrayAllPoligon(14, i)
                                Dim hgParcel As String = ArrayAllPoligon(1, i)
                                If GlobalID.Trim.Length > 2 Then
                                    If IsArray(arrayGlobalID) = True Then
                                        For j As Integer = 0 To arrayGlobalID.Length - 1
                                            If GlobalID Like arrayGlobalID(j) Then
                                                Continue For
                                            End If
                                        Next j
                                    End If
                                    ReDim arrayGlobalID(countarrayGlobalID)
                                    arrayGlobalID(countarrayGlobalID) = GlobalID
                                    countarrayGlobalID += 1
                                    'ищем входящие номера
                                    Dim ArrayCounterTemp As String(,) = Nothing
                                    Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(GlobalID.Trim, ArrayCounterTemp, ArrayTablesXML(0, 0))
                                    If IsArray(ArrayCounterTemp) = True Then
                                        For j As Integer = 0 To ArrayCounterTemp.GetUpperBound(1)
                                            Dim HandlePLine As String = ArrayCounterTemp(0, j) 'нангл контура
                                            Dim idLoc As String = ArrayCounterTemp(1, j) 'локальный ид
                                            'получаем доступ к полилинии для извлечения координат
                                            Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                            If IsNothing(idCounter) = True Then
                                                Continue For
                                            End If
                                            If idCounter.IsNull = True Then
                                                Continue For
                                            End If
                                            Dim PlineObj As Polyline = tr.GetObject(idCounter, OpenMode.ForRead)
                                            ProjectAreaParcel = ProjectAreaParcel + PlineObj.Area
                                            CountCounter += 1
                                            Dim TempNumber As String = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                            If TempNumber.Trim.Length > 0 Then
                                                If StrNumberPoint.Trim.Length = 0 Then
                                                    StrNumberPoint = TempNumber
                                                Else
                                                    StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                End If
                                            End If
                                            'ищем островки
                                            Dim ArrayOstrov1 As String(,) = Nothing
                                            Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLoc, ArrayOstrov1)
                                            If IsArray(ArrayOstrov1) = True Then
                                                For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                                    'получаем доступ к полилинии для извлечения координат
                                                    HandlePLine = ArrayOstrov1(1, k)
                                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                                    If IsNothing(idOstrov) = True Then
                                                        Continue For
                                                    End If
                                                    If idOstrov.IsNull = True Then
                                                        Continue For
                                                    End If
                                                    PlineObj = tr.GetObject(idOstrov, OpenMode.ForRead)
                                                    ProjectAreaParcel = ProjectAreaParcel - PlineObj.Area
                                                    TempNumber = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                                    If TempNumber.Trim.Length > 0 Then
                                                        If StrNumberPoint.Trim.Length = 0 Then
                                                            StrNumberPoint = TempNumber
                                                        Else
                                                            StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                        Next j
                                    End If
                                Else
                                    CountCounter = 1
                                    Dim idParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgParcel)
                                    If IsNothing(idParcel) = True Then
                                        Continue For
                                    End If
                                    If idParcel.IsNull = True Then
                                        Continue For
                                    End If
                                    Dim PlineObj As Polyline = tr.GetObject(idParcel, OpenMode.ForRead)
                                    ProjectAreaParcel = PlineObj.Area
                                    Dim TempNumber As String = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                    If TempNumber.Trim.Length > 0 Then
                                        If StrNumberPoint.Trim.Length = 0 Then
                                            StrNumberPoint = TempNumber
                                        Else
                                            StrNumberPoint = StrNumberPoint & "," & TempNumber
                                        End If
                                    End If
                                    'ищем островки
                                    Dim ArrayOstrov1 As String(,) = Nothing
                                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov1)
                                    If IsArray(ArrayOstrov1) = True Then
                                        For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                            'получаем доступ к полилинии для извлечения координат
                                            Dim HandlePLine As String = ArrayOstrov1(1, k)
                                            Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                            If IsNothing(idOstrov) = True Then
                                                Continue For
                                            End If
                                            If idOstrov.IsNull = True Then
                                                Continue For
                                            End If
                                            PlineObj = tr.GetObject(idOstrov, OpenMode.ForRead)
                                            ProjectAreaParcel = ProjectAreaParcel - PlineObj.Area
                                            TempNumber = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, PlineObj, 0.05)
                                            If TempNumber.Trim.Length > 0 Then
                                                If StrNumberPoint.Trim.Length = 0 Then
                                                    StrNumberPoint = TempNumber
                                                Else
                                                    StrNumberPoint = StrNumberPoint & "," & TempNumber
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                                StrNumberPoint = FuncFormatZN.FuncConvertStrNumberPoint(StrNumberPoint)
                                '==============================================================================================================================
                                'флаги для записи ЗУ в определенную таблицу
                                Dim boolObrazovanieZU As Boolean = False
                                Dim boolObrazovanieRazdIshZU As Boolean = False
                                Dim boolObrazovanieRazdZU As Boolean = False
                                Dim boolJoinZU As Boolean = False

                                Dim boolIzytieZU As Boolean = False
                                Dim boolRezervZU As Boolean = False

                                Dim boolPunlicLandZU As Boolean = False

                                Dim boolServitut As Boolean = False

                                Dim boolForestZu As Boolean = False
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'способ образования
                                Dim strObrazovanieZu As String = ArrayAllPoligon(10, i).Trim
                                If strObrazovanieZu.IndexOf("осударствен") > -1 Then
                                    boolObrazovanieZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                ElseIf strObrazovanieZu.IndexOf("муниципал") > -1 Then
                                    boolObrazovanieZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                ElseIf Mid(strObrazovanieZu, 1) Like "1" Then
                                    boolObrazovanieZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка из земель, находящихся в государственной или муниципальной собственности"
                                ElseIf strObrazovanieZu.IndexOf("исходн") > -1 Then
                                    boolObrazovanieRazdIshZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка с сохранением исходного в измененных границах"
                                ElseIf Mid(strObrazovanieZu, 1) Like "2" Then
                                    boolObrazovanieRazdIshZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка с сохранением исходного в измененных границах"
                                ElseIf strObrazovanieZu.IndexOf("раздел") > -1 Then
                                    boolObrazovanieRazdZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка"
                                ElseIf Mid(strObrazovanieZu, 1) Like "3" Then
                                    boolObrazovanieRazdZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем раздела земельного участка"
                                ElseIf strObrazovanieZu.IndexOf("объединен") > -1 Then
                                    boolJoinZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем объединения земельных участков"
                                ElseIf Mid(strObrazovanieZu, 1) Like "4" Then
                                    boolJoinZU = True
                                    ArrayAllPoligon(10, i) = "Образование земельного участка путем объединения земельных участков"
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'способ изъятия
                                Dim strRezervZu As String = ArrayAllPoligon(11, i)
                                If strRezervZu.IndexOf("зъятие") > -1 Then
                                    boolIzytieZU = True
                                ElseIf strRezervZu.IndexOf("езервиров") > -1 Then
                                    boolRezervZU = True
                                ElseIf Mid(strRezervZu, 1) Like "2" Then
                                    boolRezervZU = True
                                ElseIf Mid(strRezervZu, 1) Like "1" Then
                                    boolIzytieZU = True
                                End If

                                Dim strPublicZu As String = ArrayAllPoligon(12, i)
                                If strPublicZu.IndexOf("Да") > -1 Then
                                    boolPunlicLandZU = True
                                End If

                                If ArrayTablesXML(0, 0) Like "TEMPORARY_PARCEL" Then
                                    boolServitut = True
                                End If
                                '=============================================================================================================
                                'заполняем таблицы
                                'находим нужный лист(Ведомость объемов работ)
                                Dim countRows As Integer = 4
                                If boolObrazovanieZU = True And FormReport.CheckBox1.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист1")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'номера характерных точек
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 3).Range.Text = StrNumberPoint
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Kategory As String = ArrayAllPoligon(5, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 5).Range.Text = Kategory 'категория земель
                                        Else
                                            xlsheet.Cell(countRows, 5).Range.Text = "-" 'категория земель
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 5).Range.Text = "-" 'категория земель
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Utiliza As String = ArrayAllPoligon(6, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 6).Range.Text = Utiliza  'планируемый вид разрешенного использования
                                        Else
                                            xlsheet.Cell(countRows, 6).Range.Text = "-" 'планируемый вид разрешенного использования
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 6).Range.Text = "-" 'планируемый вид разрешенного использования
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 7).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 7).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 7).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Способы образования земельных участков
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim SposobObr As String = ArrayAllPoligon(10, i)
                                    xlsheet.Cell(countRows, 8).Range.Text = SposobObr
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 9).Range.Text = Note
                                    boolObrazovanieZU = False
                                    Continue For
                                End If
                                If boolObrazovanieRazdIshZU = True And FormReport.CheckBox2.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист2")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'номера характерных точек
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 3).Range.Text = StrNumberPoint
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Кадастровые номера земельных участков, из которых образуются земельные участки
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim cad_number As String = ArrayAllPoligon(3, i)
                                    If IsNothing(cad_number) = False Then
                                        If cad_number.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                        Else
                                            xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 5).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель существующая
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Kategory As String = ArrayAllPoligon(22, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 7).Range.Text = Kategory 'категория земель сушествующая
                                        Else
                                            xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель сушествующая
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель сушествующая
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель планируемая
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Kategory = ArrayAllPoligon(5, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 8).Range.Text = Kategory 'категория земель проектная
                                        Else
                                            xlsheet.Cell(countRows, 8).Range.Text = "-" 'категория земель проектная
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 8).Range.Text = "-" 'категория земель проектная
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование существующее
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Utiliza As String = ArrayAllPoligon(23, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 9).Range.Text = Utiliza  'планируемый вид разрешенного использования существующее
                                        Else
                                            xlsheet.Cell(countRows, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование планируемое
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Utiliza = ArrayAllPoligon(6, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 10).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                        Else
                                            xlsheet.Cell(countRows, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Способы образования земельных участков
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim SposobObr As String = ArrayAllPoligon(10, i)
                                    xlsheet.Cell(countRows, 11).Range.Text = SposobObr
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 12).Range.Text = Note
                                    boolObrazovanieRazdIshZU = False
                                    Continue For
                                End If
                                If boolObrazovanieRazdZU = True And FormReport.CheckBox3.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист3")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'номера характерных точек
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 3).Range.Text = StrNumberPoint
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Кадастровые номера земельных участков, из которых образуются земельные участки
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim cad_number As String = ArrayAllPoligon(3, i)
                                    If IsNothing(cad_number) = False Then
                                        If cad_number.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                        Else
                                            xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 5).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель существующая
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Kategory As String = ArrayAllPoligon(22, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 7).Range.Text = Kategory 'категория земель сушествующая
                                        Else
                                            xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель сушествующая
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель сушествующая
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель планируемая
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Kategory = ArrayAllPoligon(5, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 8).Range.Text = Kategory 'категория земель проектная
                                        Else
                                            xlsheet.Cell(countRows, 8).Range.Text = "-" 'категория земель проектная
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 8).Range.Text = "-" 'категория земель проектная
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование существующее
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Utiliza As String = ArrayAllPoligon(23, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 9).Range.Text = Utiliza  'планируемый вид разрешенного использования существующее
                                        Else
                                            xlsheet.Cell(countRows, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 9).Range.Text = "-" 'планируемый вид разрешенного использования существующее
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование планируемое
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Utiliza = ArrayAllPoligon(6, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 10).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                        Else
                                            xlsheet.Cell(countRows, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 10).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Способы образования земельных участков
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim SposobObr As String = ArrayAllPoligon(10, i)
                                    xlsheet.Cell(countRows, 11).Range.Text = SposobObr
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 12).Range.Text = Note
                                    boolObrazovanieRazdZU = False
                                    Continue For
                                End If
                                If boolJoinZU = True And FormReport.CheckBox4.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист4")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'номера характерных точек
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 3).Range.Text = StrNumberPoint
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Кадастровые или условные номера земельных участков, подлежащих объединению
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Dim cad_number As String = ArrayAllPoligon(3, i)
                                    'If IsNothing(cad_number) = False Then
                                    '    If cad_number.Trim.Length > 0 Then
                                    '        tabl.Cell(countTableword, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                    '    Else
                                    '        tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                    '    End If
                                    'Else
                                    '    tabl.Cell(countTableword, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                    'End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 5).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 6).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'категория земель планируемая
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Kategory As String = ArrayAllPoligon(5, i)
                                    If IsNothing(Kategory) = False Then
                                        If Kategory.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 7).Range.Text = Kategory 'категория земель проектная
                                        Else
                                            xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель проектная
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 7).Range.Text = "-" 'категория земель проектная
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'разрешенное использование планируемое
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Utiliza As String = ArrayAllPoligon(6, i)
                                    If IsNothing(Utiliza) = False Then
                                        If Utiliza.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 8).Range.Text = Utiliza  'планируемый вид разрешенного использования планируемый
                                        Else
                                            xlsheet.Cell(countRows, 8).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 8).Range.Text = "-" 'планируемый вид разрешенного использования планируемый
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Способы образования земельных участков
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim SposobObr As String = ArrayAllPoligon(10, i)
                                    xlsheet.Cell(countRows, 9).Range.Text = SposobObr
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 10).Range.Text = Note
                                    boolJoinZU = False
                                    Continue For
                                End If
                                If boolIzytieZU = True And FormReport.CheckBox5.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист5")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 3).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 6).Range.Text = Note
                                    boolIzytieZU = False
                                    Continue For
                                End If

                                If boolRezervZU = True And FormReport.CheckBox6.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист6")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 3).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 6).Range.Text = Note
                                    boolRezervZU = False
                                    Continue For
                                End If

                                If boolPunlicLandZU = True And FormReport.CheckBox7.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист7")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 3).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 3).Range.Text = "-" 'адрес
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 4).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Примечание
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim Note As String = ArrayAllPoligon(15, i)
                                    xlsheet.Cell(countRows, 6).Range.Text = Note
                                    boolPunlicLandZU = False
                                    Continue For

                                End If
                                If boolServitut = True And FormReport.CheckBox8.Checked = True Then
                                    Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист8")
                                    xlsheet.Activate()
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'записываем номер участка
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim ConditionParcel As String = ArrayAllPoligon(2, i)
                                    If CountCounter > 1 Then
                                        ConditionParcel = ConditionParcel & vbLf & "Количество контуров - " & CountCounter
                                    End If
                                    If IsNothing(ConditionParcel) = False Then
                                        If ConditionParcel.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 2).Range.Text = ConditionParcel 'условный номер участка
                                        Else
                                            xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 2).Range.Text = "-" 'условный номер участка
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    xlsheet.Cell(countRows, 3).Range.Text = Math.Round(ProjectAreaParcel, 0) 'площадь
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'Кадастровые номера земельных участков, из которых образуются земельные участки
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim cad_number As String = ArrayAllPoligon(3, i)
                                    If IsNothing(cad_number) = False Then
                                        If cad_number.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 4).Range.Text = cad_number 'кадастровый номер исходного ЗУ
                                        Else
                                            xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 4).Range.Text = "-" 'кадастровый номер исходного ЗУ
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'местоположение
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim AddressP As String = ArrayAllPoligon(4, i)
                                    If IsNothing(AddressP) = False Then
                                        If AddressP.Trim.Length > 0 Then
                                            xlsheet.Cell(countRows, 5).Range.Text = AddressP 'адрес
                                        Else
                                            xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                        End If
                                    Else
                                        xlsheet.Cell(countRows, 5).Range.Text = "-" 'адрес
                                    End If
                                    boolServitut = False
                                    Continue For
                                End If
                                countRows += 1
                            Next
                            xlBook = Nothing
                            xlApp.Quit()
                            xlApp = Nothing
                        End If
                    Catch
                        MsgBox("Не удалось записать данные в выбранную таблицу!!!")
                        xlBook = Nothing
                        xlApp = Nothing
                        Exit Sub
                    End Try
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    '=========================================================================================================
    'Экспликация земельных участков (отчет по всем образуемым участкам, материалы по обоснованию проекта межевания)
    <CommandMethod("ListLandAllParcels")>
    Public Sub ListLandAllParcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '==================================================================================================
        'фильтр таблиц PS
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        ReDim ArrayTablesXML(1, 1)
        Dim NameProjectOD As String = "PROJECT_PARCEL"
        Dim NameTemporaryOD As String = "TEMPORARY_PARCEL"
        ArrayTablesXML(0, 0) = NameProjectOD
        ArrayTablesXML(1, 0) = "Земельные участки образуемые"
        ArrayTablesXML(0, 1) = NameTemporaryOD
        ArrayTablesXML(1, 1) = "Участки временного отвода"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'автосохранение
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
        '========================================================================================
        'загружаем пути к сборкам
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'путь к шаблону для формирования отчета
        Dim PutcTableDoc As String = parentPatch2 & "\TemplateTables\Territory_planning\"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'забираем все мультивыноски
        Dim ArrayMLeader As String(,) = Nothing
        Dim NameStylesMLeader As String = "CADASTRAL_OBJECT"
        Dim ArrayParcelStylesMLeader As String() = {NameStylesMLeader}
        Dim boolFindMleader As Boolean = FuncCadastr.FuncFindAllMLeader(ArrayMLeader, ArrayParcelStylesMLeader)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'массивы с участками
            Dim CountArrayCadastralParcels As Integer = 0
            Dim ArrayCadastralParcels As String(,) = Nothing
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим ось существующей трассы
            Dim AlignProject As Alignment = Nothing
            Dim IDPlineAlign As ObjectId = Nothing
            Dim AlignPline As Autodesk.AutoCAD.DatabaseServices.Polyline = New Polyline()
            Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите ось трассы или полилинию для сортировки участков в таблице (не обязательно): ")
            If ppr1.Status = PromptStatus.OK Then
                Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                    AlignProject = ent1
                    IDPlineAlign = AlignProject.GetPolyline
                    AlignPline = IDPlineAlign.GetObject(OpenMode.ForRead)
                ElseIf TypeOf ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                    IDPlineAlign = ent1.Id
                    AlignPline = IDPlineAlign.GetObject(OpenMode.ForRead)
                ElseIf TypeOf ent1 Is Autodesk.AutoCAD.DatabaseServices.Polyline2d Then
                    IDPlineAlign = ent1.Id
                    Dim Pline2d As Polyline2d = IDPlineAlign.GetObject(OpenMode.ForWrite)
                    AlignPline.ConvertFrom(Pline2d, True)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем все кадастровые участки на чертеже
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                        Dim nameOD As String = ""
                        Dim boolFindTablePSParcel As Boolean = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_PARCEL")
                        If boolFindTablePSParcel = True Then
                            nameOD = "LAND_CADASTRAL_PARCEL"
                        End If
                        If boolFindTablePSParcel = False Then
                            boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_DIVISION")
                            If boolFindTablePSParcel = True Then
                                nameOD = "LAND_CADASTRAL_DIVISION"
                            End If
                        End If
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayCadastralParcels(6, CountArrayCadastralParcels)
                            ArrayCadastralParcels(0, CountArrayCadastralParcels) = ent.Handle.ToString
                            ArrayCadastralParcels(1, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.X
                            ArrayCadastralParcels(2, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.Y
                            ArrayCadastralParcels(3, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.X
                            ArrayCadastralParcels(4, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.Y
                            ArrayCadastralParcels(5, CountArrayCadastralParcels) = "" 'сюда будут записаны все образуемые ЗУ которые входят в этот полигон
                            ArrayCadastralParcels(6, CountArrayCadastralParcels) = nameOD
                            CountArrayCadastralParcels += 1
                            Continue For
                        End If
                    End If
                End If
            Next
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '=========================================================================================================
            'выбираем проектные участки
            Dim ArrayBuildParcels As String(,) = Nothing
            Dim CountBuildParcels As Integer = 0
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim ParcelPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                                    'массив с расширенными данными
                                    Dim ArrayData As String() = Nothing
                                    Dim boolFindTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayData, ArrayTablesXML)
                                    If IsArray(ArrayData) = True Then
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'определение центра полигона 
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        Dim CoordText As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(ParcelPline)
                                        '==================================================================================================
                                        'находим участок
                                        Dim boolFindParcel As Boolean = False
                                        'смотрим находится ли проектный участок внутри кадастрового участка
                                        If IsArray(ArrayCadastralParcels) = True Then
                                            For i As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                                                Dim minX As Double = Val(ArrayCadastralParcels(1, i))
                                                Dim minY As Double = Val(ArrayCadastralParcels(2, i))
                                                Dim maxX As Double = Val(ArrayCadastralParcels(3, i))
                                                Dim maxY As Double = Val(ArrayCadastralParcels(4, i))
                                                Dim NameODTemp As String = ArrayCadastralParcels(6, i)
                                                If NameODTemp Like "LAND_CADASTRAL_PARCEL" Then
                                                    If CoordText.X >= minX And CoordText.Y >= minY And CoordText.X <= maxX And CoordText.Y <= maxY Then
                                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                                        Dim PlineEntTemp As Polyline = idObj.GetObject(OpenMode.ForRead)
                                                        Dim boolCrosPointParcel As Boolean = MathFunction.FuncCrossPointPolyline(PlineEntTemp, CoordText.Convert2d(New Plane))
                                                        If boolCrosPointParcel = True Then
                                                            ArrayCadastralParcels(5, i) = ArrayCadastralParcels(5, i) & " " & ParcelPline.Handle.ToString
                                                            boolFindParcel = True
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If
                                        'если нет, то запускаем повторную проверку, но с кадастровыми кварталами
                                        If boolFindParcel = False Then
                                            If IsArray(ArrayCadastralParcels) = True Then
                                                For i As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                                                    Dim minX As Double = Val(ArrayCadastralParcels(1, i))
                                                    Dim minY As Double = Val(ArrayCadastralParcels(2, i))
                                                    Dim maxX As Double = Val(ArrayCadastralParcels(3, i))
                                                    Dim maxY As Double = Val(ArrayCadastralParcels(4, i))
                                                    Dim NameODTemp As String = ArrayCadastralParcels(6, i)
                                                    If NameODTemp Like "LAND_CADASTRAL_DIVISION" Then
                                                        If CoordText.X >= minX And CoordText.Y >= minY And CoordText.X <= maxX And CoordText.Y <= maxY Then
                                                            Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                                            Dim PlineEntTemp As Polyline = idObj.GetObject(OpenMode.ForRead)
                                                            Dim boolCrosPointParcel As Boolean = MathFunction.FuncCrossPointPolyline(PlineEntTemp, CoordText.Convert2d(New Plane))
                                                            If boolCrosPointParcel = True Then
                                                                ArrayCadastralParcels(5, i) = ArrayCadastralParcels(5, i) & " " & ParcelPline.Handle.ToString
                                                                boolFindParcel = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                        'если все равно неудача, обрабатываем участок как самостоятельный
                                        If boolFindParcel = False Then
                                            Dim boolnameTBL As Boolean = FuncPropertySet.FuncFindPSetTable(ParcelPline, NameProjectOD)
                                            If boolnameTBL = True Then
                                                Dim posArrayCadastrParcels As Integer = ArrayCadastralParcels.GetUpperBound(1) + 1
                                                ReDim Preserve ArrayCadastralParcels(6, posArrayCadastrParcels)
                                                ArrayCadastralParcels(0, posArrayCadastrParcels) = ""
                                                ArrayCadastralParcels(1, posArrayCadastrParcels) = ParcelPline.GeometricExtents.MinPoint.X
                                                ArrayCadastralParcels(2, posArrayCadastrParcels) = ParcelPline.GeometricExtents.MinPoint.Y
                                                ArrayCadastralParcels(3, posArrayCadastrParcels) = ParcelPline.GeometricExtents.MaxPoint.X
                                                ArrayCadastralParcels(4, posArrayCadastrParcels) = ParcelPline.GeometricExtents.MaxPoint.Y
                                                ArrayCadastralParcels(5, posArrayCadastrParcels) = ParcelPline.Handle.ToString
                                                ArrayCadastralParcels(6, posArrayCadastrParcels) = NameProjectOD
                                            Else
                                                Continue For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'заносим все в результирующий массив
            Dim GlobalArrayParcels As String(,) = Nothing
            Dim CountGlobalArrayParcels As Integer = 0
            '============================================================================================================
            'перебираем массив с кадастровыми участками и отфильтровываем только те которые попались в кадастровые участки
            If IsArray(ArrayCadastralParcels) = True Then
                For i As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                    Dim StrHgProjectParcel As String = ArrayCadastralParcels(5, i) 'handle проектируемых участков
                    Dim StrHgParentParcel As String = ArrayCadastralParcels(0, i) 'handle исходного земельного участка
                    'если внутри исходного ЗУ есть проектируемые участки то обрабатываем эту ветку
                    If StrHgParentParcel.Trim.Length > 0 Then
                        If IsNothing(StrHgProjectParcel) = False Then
                            If StrHgProjectParcel.Trim.Length > 0 Then
                                If IsNothing(StrHgParentParcel) = True Then
                                    Continue For
                                End If
                                If StrHgParentParcel.Trim.Length = 0 Then
                                    Continue For
                                End If
                                'получаем доступ к исходному ЗУ
                                Dim IDParentPline As ObjectId = AcadFunction.FuncReturnObjectID(StrHgParentParcel)
                                If IsNothing(IDParentPline) = True Then
                                    Continue For
                                End If
                                If IDParentPline.IsNull = True Then
                                    Continue For
                                End If
                                Dim ParentParcel As Autodesk.AutoCAD.DatabaseServices.Entity = IDParentPline.GetObject(OpenMode.ForRead)
                                'условный номер
                                Dim Cad_numb As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Cad_number", ArrayCadastralParcels(6, i))
                                If Cad_numb Like "" Then
                                    Cad_numb = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "CadastralNumber", ArrayCadastralParcels(6, i))
                                End If
                                'Вид права
                                Dim Right As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Right", ArrayCadastralParcels(6, i))
                                'имя владельца
                                Dim Patronymic As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Patronymic", ArrayCadastralParcels(6, i))
                                'адрес
                                Dim Adress As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Adress", ArrayCadastralParcels(6, i))
                                'Категория земель
                                Dim Category As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Category", ArrayCadastralParcels(6, i))
                                'Разрешенное использование
                                Dim Utilization As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Utilization", ArrayCadastralParcels(6, i))
                                If Utilization Like "" Then
                                    Utilization = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "UtilizationByDoc", ArrayCadastralParcels(6, i))
                                End If
                                'Примечание
                                Dim note As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "NOTE", ArrayCadastralParcels(6, i))
                                'Аренда
                                Dim Encumbrance As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Encumbrance", ArrayCadastralParcels(6, i))
                                'Площадь
                                Dim AreaParentParcel As Double = 0
                                If ArrayCadastralParcels(6, i) Like NameProjectOD Then
                                    Dim AreaParentParcelStr As String = FuncPropertySet.FuncReadValueSetProperties(ParentParcel, "Area", ArrayCadastralParcels(6, i))
                                    Dim AreaValParentParcel As Double = Val(AreaParentParcelStr)
                                    If AreaValParentParcel = 0 Then
                                        If TypeOf ParentParcel Is Polyline Then
                                            Dim acPlineParcel As Polyline = ParentParcel
                                            AreaParentParcel = Math.Round(acPlineParcel.Area, 0)
                                        End If
                                    End If
                                End If
                                'находим положение участка относительно трассы
                                Dim CenterPointCreateParcel As Point3d = MathFunction.FuncFindCenterObjectByGeometricExtents(ParentParcel)
                                Dim DistPline As Double = 0
                                Try
                                    Dim PKPoint As Point3d = AlignPline.GetClosestPointTo(CenterPointCreateParcel, True)
                                    DistPline = AlignPline.GetDistAtPoint(PKPoint)
                                Catch ex As System.Exception
                                End Try
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'записываем все характеристики в результирующий массив
                                ReDim Preserve GlobalArrayParcels(12, CountGlobalArrayParcels)
                                GlobalArrayParcels(0, CountGlobalArrayParcels) = DistPline 'местоположение относительно трассы
                                GlobalArrayParcels(1, CountGlobalArrayParcels) = StrHgProjectParcel 'Hahdle проектирумых участков
                                GlobalArrayParcels(2, CountGlobalArrayParcels) = Cad_numb 'кадастрывый номер исходного зу
                                GlobalArrayParcels(3, CountGlobalArrayParcels) = Right 'вид права
                                GlobalArrayParcels(4, CountGlobalArrayParcels) = "" 'не используется
                                GlobalArrayParcels(5, CountGlobalArrayParcels) = Patronymic 'правообладатель
                                GlobalArrayParcels(6, CountGlobalArrayParcels) = Adress 'адрес
                                GlobalArrayParcels(7, CountGlobalArrayParcels) = Category 'категория исходного зу
                                GlobalArrayParcels(8, CountGlobalArrayParcels) = Utilization 'разрешенное использование исходного зу
                                GlobalArrayParcels(9, CountGlobalArrayParcels) = Encumbrance 'ограничение
                                GlobalArrayParcels(10, CountGlobalArrayParcels) = AreaParentParcel 'площадь исходного ЗУ
                                GlobalArrayParcels(11, CountGlobalArrayParcels) = note 'примечание
                                CountGlobalArrayParcels += 1
                            End If
                        End If
                    Else
                        'предположительно неразграниченная собственность
                        If IsNothing(StrHgProjectParcel) = False Then
                            If StrHgProjectParcel.Trim.Length > 0 Then
                                Dim IDPline As ObjectId = AcadFunction.FuncReturnObjectID(StrHgProjectParcel)
                                If IsNothing(IDPline) = True Then
                                    Continue For
                                End If
                                If IDPline.IsNull = True Then
                                    Continue For
                                End If
                                Dim ProjectParcel As Polyline = IDPline.GetObject(OpenMode.ForRead)
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'создаем участок
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'условный номер
                                Dim Cad_numb As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Cad_number", NameProjectOD)
                                If Cad_numb Like "" Then
                                    Cad_numb = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "CadastralNumber", NameProjectOD)
                                End If
                                'Вид права
                                Dim Right As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Right", NameProjectOD)
                                'имя владельца
                                Dim Patronymic As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Patronymic", NameProjectOD)
                                'адрес
                                Dim Adress As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Adress", NameProjectOD)
                                'Категория земель
                                Dim Category As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Category", NameProjectOD)
                                'Разрешенное использование
                                Dim Utilization As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Utilization", NameProjectOD)
                                If Utilization Like "" Then
                                    Utilization = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "UtilizationByDoc", NameProjectOD)
                                End If
                                If Utilization Like "" Then
                                    Utilization = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "IntendentUse", NameProjectOD)
                                End If
                                'Примечание
                                Dim note As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "NOTE", NameProjectOD)
                                'Аренда
                                Dim Encumbrance As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Encumbrance", NameProjectOD)
                                'Площадь
                                Dim AreaParentParcel As String = FuncPropertySet.FuncReadValueSetProperties(ProjectParcel, "Area", NameProjectOD)
                                Dim AreaValParentParcel As Double = Val(AreaParentParcel)
                                If AreaValParentParcel = 0 Then
                                    AreaValParentParcel = Math.Round(ProjectParcel.Area, 0)
                                End If
                                'находим положение участка относительно трассы
                                Dim CenterPointCreateParcel As Point3d = MathFunction.FuncFindCenterObjectByGeometricExtents(ProjectParcel)
                                Dim DistPline As Double = 0
                                Try
                                    Dim PKPoint As Point3d = AlignPline.GetClosestPointTo(CenterPointCreateParcel, True)
                                    DistPline = AlignPline.GetDistAtPoint(PKPoint)
                                Catch ex As System.Exception
                                End Try
                                ReDim Preserve GlobalArrayParcels(12, CountGlobalArrayParcels)
                                GlobalArrayParcels(0, CountGlobalArrayParcels) = DistPline 'местоположение относительно трассы
                                GlobalArrayParcels(1, CountGlobalArrayParcels) = StrHgProjectParcel 'Hahdle проектирумых участков
                                GlobalArrayParcels(2, CountGlobalArrayParcels) = Cad_numb 'кадастрывый номер исходного зу
                                GlobalArrayParcels(3, CountGlobalArrayParcels) = Right 'вид права
                                GlobalArrayParcels(4, CountGlobalArrayParcels) = "" 'не используется
                                GlobalArrayParcels(5, CountGlobalArrayParcels) = Patronymic 'правообладатель
                                GlobalArrayParcels(6, CountGlobalArrayParcels) = Adress 'адрес
                                GlobalArrayParcels(7, CountGlobalArrayParcels) = Category 'категория исходного зу
                                GlobalArrayParcels(8, CountGlobalArrayParcels) = Utilization 'разрешенное использование исходного зу
                                GlobalArrayParcels(9, CountGlobalArrayParcels) = Encumbrance 'ограничение
                                GlobalArrayParcels(10, CountGlobalArrayParcels) = AreaParentParcel 'площадь исходного ЗУ
                                GlobalArrayParcels(11, CountGlobalArrayParcels) = note 'примечание
                                CountGlobalArrayParcels += 1
                            End If
                        End If
                    End If
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сортируем массив
            If IsArray(GlobalArrayParcels) = True Then
                If GlobalArrayParcels.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(GlobalArrayParcels, 0)
                End If
            Else
                MsgBox("Участки не найдены!!! Макро превано.")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выводим отчет
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'формируем участки
            If IsArray(GlobalArrayParcels) = True Then
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                Try
                    Dim PutchDoc As String = PutcTableDoc & "ПМТ_ПЗ-Экспликация земельных участков.docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "ПМТ_ПЗ-Экспликация земельных участков_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    WApp.Visible = True
                    Dim tabl As Word.Table = Nothing
                    For Each tbl As Word.Table In WDoc.Tables
                        tabl = tbl
                        If tabl.Rows.Count >= 15 Then
                            Exit For
                        End If
                    Next
                    If IsNothing(tabl) = True Then
                        MsgBox("Не удалось найти таблицу с количеством стобцов равным 15. Макро прервано!!!")
                        WDoc.Close()
                        WApp.Quit()
                        WApp = Nothing
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'заполняем таблицу
                    Dim countTableRow As Integer = 3 'число строк
                    Dim countNumbers As Integer = 1 'порядковый номер
                    For i As Integer = 0 To GlobalArrayParcels.GetUpperBound(1)
                        Dim StrHgParcel As String = GlobalArrayParcels(1, i) 'handle объекта
                        If IsNothing(StrHgParcel) = True Then
                            Continue For
                        End If
                        If StrHgParcel.Trim.Length = 0 Then
                            Continue For
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'считываем установки с родительского участка (кадастрового участка или квартала)
                        Dim Cad_numb As String = GlobalArrayParcels(2, i) 'кадастровый номер исходного зу
                        If IsNothing(Cad_numb) = True Then
                            Cad_numb = ""
                        ElseIf Cad_numb.Trim.Length = 0 Then
                            Cad_numb = ""
                        End If
                        Dim Right As String = GlobalArrayParcels(3, i)  'вид права исходного зу
                        If IsNothing(Right) = True Then
                            Right = ""
                        ElseIf Right.Trim.Length = 0 Then
                            Right = ""
                        End If
                        Dim Patronymic As String = GlobalArrayParcels(5, i)  'правообладатель исходного зу
                        If IsNothing(Patronymic) = True Then
                            Patronymic = ""
                        ElseIf Patronymic.Trim.Length = 0 Then
                            Patronymic = ""
                        End If
                        Dim Adress As String = GlobalArrayParcels(6, i) '  адрес исходного зу
                        If IsNothing(Adress) = True Then
                            Adress = ""
                        ElseIf Adress.Trim.Length = 0 Then
                            Adress = ""
                        End If
                        Dim Category As String = GlobalArrayParcels(7, i)  'категория исходного зу
                        If IsNothing(Category) = True Then
                            Category = ""
                        ElseIf Category.Trim.Length = 0 Then
                            Category = ""
                        End If
                        Dim Utilization As String = GlobalArrayParcels(8, i)   'разрешенное использование исходного зу
                        If IsNothing(Utilization) = True Then
                            Utilization = ""
                        ElseIf Utilization.Trim.Length = 0 Then
                            Utilization = ""
                        End If
                        Dim Encumbrance As String = GlobalArrayParcels(9, i)   'ограничение
                        If IsNothing(Encumbrance) = True Then
                            Encumbrance = ""
                        ElseIf Encumbrance.Trim.Length = 0 Then
                            Encumbrance = ""
                        End If
                        Dim AreaParentParcel As Double = Val(GlobalArrayParcels(10, i))   'площадь исходного ЗУ
                        Dim note As String = GlobalArrayParcels(11, i)  'примечание на исходный ЗУ
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '======================================================================================================
                        tabl.Rows.Add()
                        tabl.Cell(countTableRow, 1).Range.Text = countNumbers 'порядковый номер 
                        countNumbers += 1
                        'заполняем таблицу с характеристиками исходного ЗУ
                        '=====================================================================================================
                        'кадастровый номер и адрес исходного участка
                        Dim str As String = Cad_numb
                        If Adress.Trim.Length > 0 And str.Trim.Length > 0 Then
                            str = Cad_numb & vbLf & Adress
                        ElseIf Adress.Trim.Length > 0 And str.Trim.Length = 0 Then
                            str = Adress
                        End If
                        tabl.Cell(countTableRow, 2).Range.Text = str
                        '=====================================================================================================
                        'площадь исходного ЗУ
                        Dim AreaParentZu As Double = Math.Round(AreaParentParcel, 0)
                        If AreaParentZu > 0 Then
                            tabl.Cell(countTableRow, 3).Range.Text = AreaParentZu
                        Else
                            tabl.Cell(countTableRow, 3).Range.Text = "-"
                        End If
                        '========================================================================================================
                        'категория земель исходного участка
                        If Category.Trim.Length = 0 Then
                            Category = "Категория не установлена"
                        End If
                        tabl.Cell(countTableRow, 6).Range.Text = Category
                        '=======================================================================================================
                        'разрешенное использование исходного участка
                        If Utilization.Trim.Length = 0 Then
                            Utilization = "-"
                        End If
                        tabl.Cell(countTableRow, 7).Range.Text = Utilization
                        '====================================================================================================
                        'вид права, правообладатель, обременнеие
                        str = Right
                        If Patronymic.Trim.Length > 0 And str.Trim.Length > 0 Then
                            str = str & ", " & Patronymic
                        ElseIf Patronymic.Trim.Length > 0 And str.Trim.Length = 0 Then
                            str = Patronymic
                        End If
                        If Encumbrance.Trim.Length > 0 And str.Trim.Length > 0 Then
                            str = str & ", " & Encumbrance
                        ElseIf Encumbrance.Trim.Length > 0 And str.Trim.Length = 0 Then
                            str = Encumbrance
                        End If
                        tabl.Cell(countTableRow, 12).Range.Text = str
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '=========================================================================================================
                        'заполняем строки с образуемыми ЗУ
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim ArrayProjectParcels As String() = StrHgParcel.Trim.Split 'получаем handle всех участков которые входят в исходный ЗУ
                        Dim CountRows As Integer = 1
                        If IsArray(ArrayProjectParcels) = True Then
                            CountRows = ArrayProjectParcels.Length
                            Dim c9 As Word.Cell = tabl.Cell(countTableRow, 9)
                            Dim c10 As Word.Cell = tabl.Cell(countTableRow, 10)
                            Dim c11 As Word.Cell = tabl.Cell(countTableRow, 11)
                            Dim c12 As Word.Cell = tabl.Cell(countTableRow, 12)
                            Dim c13 As Word.Cell = tabl.Cell(countTableRow, 13)
                            Dim c14 As Word.Cell = tabl.Cell(countTableRow, 14)
                            Dim c15 As Word.Cell = tabl.Cell(countTableRow, 15)
                            If CountRows > 1 Then
                                c9.Split(CountRows, 1)
                                c10.Split(CountRows, 1)
                                c11.Split(CountRows, 1)
                                c12.Split(CountRows, 1)
                                c13.Split(CountRows, 1)
                                c14.Split(CountRows, 1)
                                c15.Split(CountRows, 1)
                            End If
                            '================================================================================================================
                            'заполняем строки
                            Dim ArrayProjectPacelsData As String(,) = Nothing
                            Dim CountArrayProjectParcelData As Integer = 0
                            Dim RezultAreaProjectParcels As Double = 0
                            For j As Integer = 0 To ArrayProjectParcels.Length - 1
                                Dim StrHgCounterParcel As String = ArrayProjectParcels(j)
                                Dim TypeProjectParcel As String = ""
                                Dim conditionalNumber As String = ""
                                Dim ProjectUtilization As String = ""
                                Dim ProjectCategory As String = "Земли промышленности, транспорта и т.п."
                                Dim ProjectRight As String = ""
                                Dim ProjectPatronymic As String = ""
                                Dim ProjectEncumbrance As String = ""
                                Dim ProjectAreaParcel As Double = 0
                                Dim ProjectConditionalNumber As String = ""
                                Dim ProjectMetodCreated As String = ""
                                Dim ProjectConfiscation_parcel As String = ""
                                Dim ProjectRationale_size As String = ""
                                Dim LocalID As String = ""
                                Dim IDCounterParcel As ObjectId = AcadFunction.FuncReturnObjectID(StrHgCounterParcel)
                                If IsNothing(IDCounterParcel) = False Then
                                    If IDCounterParcel.IsNull = False Then
                                        Dim PlineProjectParcel As Polyline = IDCounterParcel.GetObject(OpenMode.ForRead)
                                        Dim ArrayDataObj As String(,) = Nothing
                                        Dim boolFindTablePS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(PlineProjectParcel, ArrayDataObj)
                                        If IsArray(ArrayDataObj) = True Then
                                            For k As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                                Dim NameField As String = ArrayDataObj(1, k)
                                                Dim valdata As String = ArrayDataObj(2, k)
                                                If NameField Like "ConditionalNumber" Then
                                                    ProjectConditionalNumber = valdata
                                                ElseIf NameField Like "LABEL" Then
                                                    TypeProjectParcel = valdata
                                                ElseIf NameField Like "Right" Then
                                                    If ProjectRight.Trim.Length = 0 Then
                                                        ProjectRight = valdata
                                                    End If
                                                ElseIf NameField Like "Patronymic" Then
                                                    If ProjectPatronymic.Trim.Length = 0 Then
                                                        ProjectPatronymic = valdata
                                                    End If
                                                ElseIf NameField Like "IntendentUse" Then
                                                    If ProjectUtilization.Trim.Length = 0 Then
                                                        ProjectUtilization = valdata
                                                    End If
                                                ElseIf NameField Like "Utilization" Then
                                                    If ProjectUtilization.Trim.Length = 0 Then
                                                        ProjectUtilization = valdata
                                                    End If
                                                ElseIf NameField Like "Encumbrance" Then
                                                    ProjectEncumbrance = valdata
                                                ElseIf NameField Like "Category" Then
                                                    ProjectCategory = valdata
                                                ElseIf NameField Like "Method_create" Then
                                                    ProjectMetodCreated = valdata
                                                ElseIf NameField Like "Confiscation_parcel" Then
                                                    ProjectConfiscation_parcel = valdata
                                                ElseIf NameField Like "Rationale_size" Then
                                                    ProjectRationale_size = valdata
                                                ElseIf NameField Like "LOCALID" Then
                                                    LocalID = valdata
                                                End If
                                            Next k
                                        End If
                                        '=========================================================================================
                                        'площадь участка
                                        ProjectAreaParcel = PlineProjectParcel.Area
                                        '===========================================================================================
                                        'ищем мультивыноску с обозначением
                                        '1. Ищем мультивыноску с меткой обозначения
                                        If LocalID.Trim.Length > 0 Then
                                            'ищем островки у данного контура
                                            Dim ArrayOstrov As String(,) = Nothing
                                            Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov)
                                            If IsArray(ArrayOstrov) = True Then
                                                For l As Integer = 0 To ArrayOstrov.Length - 1
                                                    Dim hgOstrov As String = ArrayOstrov(1, l)
                                                    Dim IDOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                                    If IsNothing(IDOstrov) = False Then
                                                        If IDOstrov.IsNull = False Then
                                                            Dim acPline As Polyline = IDOstrov.GetObject(OpenMode.ForRead)
                                                            ProjectAreaParcel = ProjectAreaParcel - acPline.Area
                                                        End If
                                                    End If
                                                Next l
                                            End If
                                        End If
                                    End If
                                End If
                                '=============================================================================================================
                                'вычленяем короткий номер для сортировки
                                Dim NumbCounter As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ProjectConditionalNumber)
                                If TypeProjectParcel.IndexOf("ервитут") > -1 Then
                                    NumbCounter += 1000000 'для того чтобы временные участки следовали за участками постоянного отвода
                                Else
                                    If AreaParentParcel >= ProjectAreaParcel Then
                                        AreaParentParcel = AreaParentParcel - ProjectAreaParcel
                                    End If
                                    RezultAreaProjectParcels += ProjectAreaParcel
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'заполняем массив
                                ReDim Preserve ArrayProjectPacelsData(10, CountArrayProjectParcelData)
                                ArrayProjectPacelsData(0, CountArrayProjectParcelData) = NumbCounter
                                ArrayProjectPacelsData(1, CountArrayProjectParcelData) = ProjectConditionalNumber
                                ArrayProjectPacelsData(2, CountArrayProjectParcelData) = ProjectAreaParcel
                                ArrayProjectPacelsData(3, CountArrayProjectParcelData) = ProjectCategory
                                ArrayProjectPacelsData(4, CountArrayProjectParcelData) = ProjectUtilization
                                ArrayProjectPacelsData(5, CountArrayProjectParcelData) = ProjectRight
                                ArrayProjectPacelsData(6, CountArrayProjectParcelData) = ProjectPatronymic
                                ArrayProjectPacelsData(7, CountArrayProjectParcelData) = ProjectEncumbrance
                                ArrayProjectPacelsData(8, CountArrayProjectParcelData) = ProjectMetodCreated
                                ArrayProjectPacelsData(9, CountArrayProjectParcelData) = ProjectConfiscation_parcel
                                ArrayProjectPacelsData(10, CountArrayProjectParcelData) = ProjectRationale_size
                                CountArrayProjectParcelData += 1
                            Next j
                            '=================================================================================================================
                            'записываем данные из массива
                            If IsArray(ArrayProjectPacelsData) = True Then
                                If ArrayProjectPacelsData.GetUpperBound(1) > 0 Then
                                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayProjectPacelsData, 0)
                                End If
                                '==============================================================================================
                                'записываем данные в ячейки
                                For k1 As Integer = 0 To ArrayProjectPacelsData.GetUpperBound(1)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'заполняем таблицу
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'условный номер участка
                                    tabl.Cell(countTableRow + k1, 9).Range.Text = ArrayProjectPacelsData(1, k1)
                                    '=========================================================================================================
                                    'площадь проектного участка
                                    tabl.Cell(countTableRow + k1, 10).Range.Text = Math.Round(Val(ArrayProjectPacelsData(2, k1)), 0)
                                    '========================================================================================================
                                    'категория земель
                                    If ArrayProjectPacelsData(3, k1).Trim.Length = 0 Then
                                        tabl.Cell(countTableRow + k1, 11).Range.Text = "Категория не установлена"
                                    Else
                                        tabl.Cell(countTableRow + k1, 11).Range.Text = ArrayProjectPacelsData(3, k1)
                                    End If
                                    '=========================================================================================================
                                    'разрешенное использование
                                    If ArrayProjectPacelsData(4, k1).Trim.Length = 0 Then
                                        tabl.Cell(countTableRow + k1, 12).Range.Text = "-"
                                    End If
                                    tabl.Cell(countTableRow + k1, 12).Range.Text = ArrayProjectPacelsData(4, k1) '
                                    '====================================================================================================
                                    'вид права, правообладатель, обременнеие
                                    str = ArrayProjectPacelsData(5, k1)
                                    If ArrayProjectPacelsData(6, k1).Trim.Length > 0 And str.Trim.Length > 0 Then
                                        str = str & ", " & ArrayProjectPacelsData(6, k1)
                                    ElseIf ArrayProjectPacelsData(6, k1).Trim.Length > 0 And str.Trim.Length = 0 Then
                                        str = ArrayProjectPacelsData(6, k1)
                                    End If
                                    If ArrayProjectPacelsData(7, k1).Trim.Length > 0 And str.Trim.Length > 0 Then
                                        str = str & ", " & ArrayProjectPacelsData(7, k1)
                                    ElseIf ArrayProjectPacelsData(7, k1).Trim.Length > 0 And str.Trim.Length = 0 Then
                                        str = ArrayProjectPacelsData(7, k1)
                                    End If
                                    tabl.Cell(countTableRow + k1, 13).Range.Text = str
                                    '==========================================================================================================
                                    'вид кадастровых работ
                                    If ArrayProjectPacelsData(8, k1).Trim.Length = 0 Then
                                        tabl.Cell(countTableRow + k1, 14).Range.Text = "-"
                                    Else
                                        tabl.Cell(countTableRow + k1, 14).Range.Text = ArrayProjectPacelsData(8, k1)
                                    End If
                                    '=============================================================================================================
                                    'обоснование изъятия
                                    str = ArrayProjectPacelsData(9, k1)
                                    If ArrayProjectPacelsData(9, k1).Trim.Length > 0 And str.Trim.Length > 0 Then
                                        str = str & ", " & ArrayProjectPacelsData(9, k1)
                                    ElseIf ArrayProjectPacelsData(10, k1).Trim.Length > 0 And str.Trim.Length = 0 Then
                                        str = ArrayProjectPacelsData(10, k1)
                                    End If
                                    If str.Trim.Length = 0 Then
                                        str = "-"
                                    End If
                                    tabl.Cell(countTableRow + k1, 15).Range.Text = str
                                Next k1
                                '===========================================================================================================================
                                'записываем остаточную площадь и процент изъятия
                                If AreaParentParcel > 0 Then
                                    tabl.Cell(countTableRow, 5).Range.Text = Math.Round(AreaParentParcel, 0) 'остаточная площадь
                                    Dim Pr_iz As Double = (RezultAreaProjectParcels / AreaParentParcel) * 100
                                    If Pr_iz > 1 Then
                                        tabl.Cell(countTableRow, 4).Range.Text = Math.Round(Pr_iz, 0) 'процент изъятия
                                    Else
                                        tabl.Cell(countTableRow, 4).Range.Text = Math.Round(Pr_iz, 2) 'процент изъятия
                                    End If
                                Else
                                    tabl.Cell(countTableRow, 5).Range.Text = "-" 'остаточная площадь
                                    tabl.Cell(countTableRow, 4).Range.Text = "-" 'процент изъятия
                                End If
                            End If
                            countTableRow += CountRows
                        End If
                    Next i
                Catch ex As System.Exception
                End Try
            End If
            tr.Commit()
        End Using
    End Sub
    '=========================================================================================================
    'Ведомость публичных сервитутов
    <CommandMethod("ListTemporaryParcels")>
    Public Sub ListTemporaryParcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        Dim numberServitut As String = 0
        ArrayTablesXML(0, 0) = "TEMPORARY_PARCEL"
        ArrayTablesXML(1, 0) = "Участок временного отвода"
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к блокам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim minX As Double = 9999999999
                Dim minY As Double = 9999999999
                Dim maxX As Double = 0
                Dim maxY As Double = 0
                '=========================================================================================================
                'слой для вспомогательных построений
                Dim nameLayer As String = "TempCounters"
                Dim idcrLayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer)
                '=========================================================================================================
                'выбираем сервитуты
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        Dim arrayData As String(,) = {}
                                        Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, arrayData, ArrayTablesXML)
                                        If arrayData.GetUpperBound(1) > 0 Then
                                            Dim numberServ As Integer = CInt(MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2))
                                            Dim purpServ As String = MathFunction.FuncFindValueToArray2d(arrayData, "Purpose", 1, 2)
                                            Dim localID As String = MathFunction.FuncFindValueToArray2d(arrayData, "LOCALID", 1, 2)
                                            ReDim Preserve ArrayAllPoligon(3, countArrayAllPoligon)
                                            ArrayAllPoligon(0, countArrayAllPoligon) = numberServ
                                            ArrayAllPoligon(1, countArrayAllPoligon) = acEnt.Handle.ToString
                                            ArrayAllPoligon(2, countArrayAllPoligon) = purpServ
                                            ArrayAllPoligon(3, countArrayAllPoligon) = localID
                                            countArrayAllPoligon += 1
                                        End If
                                        Dim bound As Extents3d = acEnt.Bounds
                                        If bound.MinPoint.X < minX Then
                                            minX = bound.MinPoint.X
                                        End If
                                        If bound.MinPoint.Y < minY Then
                                            minY = bound.MinPoint.Y
                                        End If
                                        If bound.MaxPoint.X > maxX Then
                                            maxX = bound.MaxPoint.X
                                        End If
                                        If bound.MaxPoint.Y > maxY Then
                                            maxY = bound.MaxPoint.Y
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось найти публичные сервитуты!!!")
                    Exit Sub
                End If
                'сортируем сервитуты по номеру
                If ArrayAllPoligon.GetUpperBound(1) > 1 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
                End If

                'массив для кадастровых участком
                Dim CountArrayCadastralParcels As Integer = 0
                Dim ArrayCadastralParcels As String(,) = {}
                'массив для зданий
                Dim CountArrayCadastralBuild As Integer = 0
                Dim ArrayCadastralBuild As String(,) = {}
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем все кадастровые участки на чертеже
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim arrayTbl As String(,) = Nothing
                ReDim Preserve arrayTbl(1, 0)
                arrayTbl(0, 0) = "LAND_CADASTRAL_PARCEL"
                Dim listCadastralPolygon As Dictionary(Of String, Polygon) = New Dictionary(Of String, Polygon)
                Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
                For Each id As ObjectId In btr
                    If id.IsNull = False Then
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                        If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                            Dim arrayTablesPS As String() = {}
                            Dim findTBL As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, arrayTablesPS, arrayTbl)
                            If arrayTablesPS.Length > 0 Then
                                Dim poly As Polyline = ent
                                Dim parcelPolygon As Polygon = netTopo.CreatePolygon(poly)
                                If IsNothing(parcelPolygon) = False Then
                                    listCadastralPolygon.Add(poly.Handle.ToString, parcelPolygon)
                                End If
                            End If
                        End If
                    End If
                Next
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тело программы
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                Try
                    Dim PutchDoc As String = PutcTableDoc & "ПМТ-Ведомость публичных сервитутов.docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "ПМТ-Ведомость публичных сервитутов_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End If
                    WApp.Visible = True
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                        Dim CondParcel As String = ArrayAllPoligon(0, i) 'номер сервитута
                        Dim hgParcels As String = ArrayAllPoligon(1, i) 'hg объекта
                        Dim idParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgParcels) 'id объекта
                        Dim userParcels As Polyline = tr.GetObject(idParcel, OpenMode.ForWrite)
                        Dim userPolugon As Polygon = netTopo.CreatePolygon(userParcels)
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'Условный номер участка
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            If IsNothing(CondParcel) = False Then
                                If CondParcel.Trim.Length > 0 Then
                                    tabl.Cell(countTableword, 1).Range.Text = CondParcel 'Условный номер
                                Else
                                    tabl.Cell(countTableword, 1).Range.Text = "-" 'Условный номер
                                End If
                            Else
                                tabl.Cell(countTableword, 1).Range.Text = "-" 'Условный номер
                            End If
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'площадь
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim areaServitut As Double = MathFunction.FuncSQRTPolyline2(userParcels, 2, 0)
                            Dim localiD As String = ArrayAllPoligon(3, i)
                            Dim arrayOstrov As String(,) = {}
                            Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(localiD, arrayOstrov)
                            If IsArray(arrayOstrov) = True Then
                                If arrayOstrov.Length > 0 Then
                                    For j As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                        Dim hgOstrov As String = arrayOstrov(1, j)
                                        Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                        Dim userOstrov As Polyline = tr.GetObject(idOstrov, OpenMode.ForWrite)
                                        Dim areaOstrov As Double = MathFunction.FuncSQRTPolyline2(userOstrov, 2, 0)
                                        areaServitut -= areaOstrov
                                    Next j
                                End If
                            End If
                            tabl.Cell(countTableword, 2).Range.Text = areaServitut 'площадь
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim str As String = ""
                            If listCadastralPolygon.Count > 0 Then
                                For j As Integer = 0 To listCadastralPolygon.Count - 1
                                    Dim keyHandle As String = listCadastralPolygon.ElementAt(j).Key
                                    Dim cadParcelPolygon As Polygon = listCadastralPolygon.ElementAt(j).Value
                                    Dim boolIntersect As Boolean = netTopo.DoPolygonsIntersect(userPolugon, cadParcelPolygon)
                                    If boolIntersect = True Then
                                        Dim geo As Geometry = netTopo.GetIntersection(userPolugon, cadParcelPolygon)
                                        If geo.Area > 1 Then
                                            Dim id As ObjectId = AcadFunction.FuncReturnObjectID(keyHandle)
                                            Dim cadParcel As Polyline = id.GetObject(OpenMode.ForRead)
                                            Dim cad_number As String = FuncPropertySet.FuncReadValueSetProperties(cadParcel, "Cad_number")
                                            Dim address As String = FuncPropertySet.FuncReadValueSetProperties(cadParcel, "Adress")
                                            str = str & cad_number & vbLf & address & vbLf
                                        End If
                                    End If
                                Next j
                            End If
                            tabl.Cell(countTableword, 3).Range.Text = str
                            Exit For
                        Next
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                Catch ex As System.Exception
                    MsgBox("Не удалось записать данные в выбранную таблицу!!!" & "Номер сервитута: " & numberServitut & vbLf & ex.Message)
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub
    '=========================================================================================================
    'простой отчет по объектам
    <CommandMethod("ListUserReports")>
    Public Sub ListUserReports()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к блокам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
        Dim docloc As DocumentLock = acDoc.LockDocument()
        Using docloc
            Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
            ReDim ArrayTablesXML(1, 1)
            Dim arrayTables As String() = {}
            Dim boolTablePS As Boolean = FuncPropertySet.FuncReturnArrayNamesDictGWG(arrayTables)
            If arrayTables.Length > 0 Then
                Dim formSelectObject As SelectTemplateFiles = New SelectTemplateFiles
                formSelectObject.ListBox1.DataSource = arrayTables
                formSelectObject.ShowDialog()
                If formSelectObject.boolShow = False Then
                    Exit Sub
                End If
                ArrayTablesXML(0, 0) = formSelectObject.ListBox1.SelectedValue
                ArrayTablesXML(1, 0) = ""
            End If
            ArrayTablesXML(0, 1) = "CADASTRAL_OSTROV"
            ArrayTablesXML(1, 1) = "Островки"

            Dim arrayFields As String(,) = {}
            Dim boolf = FuncPropertySet.FuncReadAllValueFieldSetTables(ArrayTablesXML(0, 0), arrayFields)
            Dim acDBOstrovCollection As DBObjectCollection = New DBObjectCollection
            Dim acDBParcelCollection As DBObjectCollection = New DBObjectCollection
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                            If ArrayNameTableOD(i).Trim Like "CADASTRAL_OSTROV" Then
                                                acDBOstrovCollection.Add(acEnt)
                                            Else
                                                acDBParcelCollection.Add(acEnt)
                                            End If
                                        Next i
                                    End If
                                End If

                            End If
                        End If
                    Next
                End If
                'неудача - выходим из программы
                If acDBParcelCollection.Count = 0 Then
                    MsgBox("Не удалось найти объекты-участки!!!")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim parcelDictionary As Dictionary(Of Integer, String(,)) = New Dictionary(Of Integer, String(,))
                For Each acParcelsPline As Polyline In acDBParcelCollection
                    Dim arrayData As String(,) = {}
                    Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acParcelsPline, arrayData, ArrayTablesXML)
                    If IsArray(arrayData) = True Then
                        Dim ConditionalParcel As String = MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2)
                        Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionalParcel)
                        If parcelDictionary.Count > 0 Then
                            If parcelDictionary.ContainsKey(numbParcel) = True Then
                                Dim tempArray As String(,) = parcelDictionary.Item(numbParcel)
                                Dim count As Integer = tempArray.GetUpperBound(1)
                                ReDim Preserve tempArray(1, count)
                                Dim numCounter As Integer = FuncCadastr.FuncFindNumberCounterCadastralObject(ConditionalParcel)
                                tempArray(0, count) = numCounter
                                tempArray(1, count) = acParcelsPline.Handle.ToString
                                parcelDictionary.Item(numbParcel) = tempArray
                            Else
                                Dim tempArray As String(,) = {}
                                ReDim tempArray(1, 0)
                                tempArray(0, 0) = "0"
                                tempArray(1, 0) = acParcelsPline.Handle.ToString
                                parcelDictionary.Add(numbParcel, tempArray)
                            End If
                        Else
                            Dim tempArray As String(,) = {}
                            ReDim tempArray(1, 0)
                            tempArray(0, 0) = "0"
                            tempArray(1, 0) = acParcelsPline.Handle.ToString
                            parcelDictionary.Add(numbParcel, tempArray)
                        End If
                    End If
                Next
                ' Сортируем словарь по ключу с использованием LINQ
                If parcelDictionary.Count > 0 Then
                    If parcelDictionary.Count > 1 Then
                        parcelDictionary = parcelDictionary.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'выводим отчет
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Add()
                    Dim cell1 As Excel.Range 'первая ячейка в строке
                    Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                    Dim CellRng As Excel.Range 'массив ячеек
                    xlSheets = xlBook.ActiveSheet
                    Dim countRows As Integer = 2
                    Try
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'начинаем перебирать участки
                        xlApp.Visible = True
                        For i As Integer = 0 To parcelDictionary.Count - 1
                            cell1 = xlSheets.Cells(countRows, 1) 'ссылка на первую ячейку
                            cell1.Value = i + 1 'порядковый номер

                            Dim keyVal As KeyValuePair(Of Integer, String(,)) = parcelDictionary.ElementAt(i)
                            Dim arrayParcel As String(,) = keyVal.Value
                            If arrayParcel.GetUpperBound(1) > 0 Then
                                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayParcel, 0)
                            End If
                            For j As Integer = 0 To arrayParcel.GetUpperBound(1)
                                Dim hg As String = arrayParcel(1, j)
                                Dim idpLine As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                If IsNothing(idpLine) = False Then
                                    If idpLine.IsNull = False Then
                                        Dim acPline As Polyline = tr.GetObject(idpLine, OpenMode.ForRead)
                                        Dim areaParcel As Double = MathFunction.FuncSQRTPolyline2(acPline, 2, 0)
                                        Dim arrayData As String(,) = {}
                                        Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acPline, arrayData, ArrayTablesXML)
                                        If IsArray(arrayData) = True Then
                                            If i = 0 Then
                                                cell1 = xlSheets.Cells(countRows - 1, 1) 'ссылка на первую ячейку
                                                cell1.Value = "Номер п/п"
                                                For k As Integer = 0 To arrayData.GetUpperBound(1)
                                                    Dim nameField As String = arrayData(1, k)
                                                    If arrayFields.Length > 0 Then
                                                        For k2 As Integer = 0 To arrayFields.GetUpperBound(1)
                                                            If nameField Like arrayFields(1, k2) Then
                                                                nameField = arrayFields(3, k2)
                                                                Exit For
                                                            End If
                                                        Next k2
                                                    End If
                                                    cell1 = xlSheets.Cells(countRows - 1, k + 2) 'ссылка на первую ячейку
                                                    cell1.Value = nameField
                                                Next k
                                            End If
                                            Dim localID As String = MathFunction.FuncFindValueToArray2d(arrayData, "LOCALID", 1, 2)
                                            Dim arrayOstrov As String(,) = {}
                                            Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrovByDbCollection(acDBOstrovCollection, localID, arrayOstrov)
                                            If arrayOstrov.Length > 0 Then
                                                For k As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                                    Dim hgOstrov As String = arrayOstrov(1, k)
                                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hgOstrov)
                                                    If IsNothing(idOstrov) = False Then
                                                        If idOstrov.IsNull = False Then
                                                            Dim PlineOstrov As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                                            Dim areaOstrov As Double = MathFunction.FuncSQRTPolyline2(PlineOstrov, 2, 0)
                                                            areaParcel -= areaOstrov
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                            For k As Integer = 0 To arrayData.GetUpperBound(1)
                                                Dim strField As String = arrayData(1, k)
                                                If strField Like "Area" Then
                                                    cell1 = xlSheets.Cells(countRows, k + 2) 'ссылка на первую ячейку
                                                    cell1.Value = areaParcel
                                                Else
                                                    cell1 = xlSheets.Cells(countRows, k + 2) 'ссылка на первую ячейку
                                                    cell1.Value = arrayData(2, k)
                                                End If
                                            Next k
                                        End If
                                    End If
                                End If
                                countRows += 1
                            Next j
                        Next i
                    Catch ex As SystemException
                    End Try
                    xlBook = Nothing
                    xlApp.Quit()
                    xlApp = Nothing
                End If
            End Using
            docloc.Dispose()
        End Using
    End Sub
    '==========================================================================================================
    'отчет по изменяемым ЗУ
    <CommandMethod("ListIzmParcels")>
    Public Sub ListIzmParcels()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        Dim numberServitut As String = 0
        ArrayTablesXML(0, 0) = "LAND_CADASTRAL_PARCEL"
        ArrayTablesXML(1, 0) = ""
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к блокам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                '=========================================================================================================
                'выбираем проектные участки
                Dim listProjectPolygon As List(Of Polygon) = New List(Of Polygon)
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim polyProjectParcel As Polyline = acEnt
                                    Dim boolFindTablePS As Boolean = FuncPropertySet.FuncFindPSetTable(acEnt, "PROJECT_PARCEL")
                                    If boolFindTablePS = True Then
                                        Dim geoPolygon As Polygon = netTopo.CreatePolygon(polyProjectParcel)
                                        If IsNothing(geoPolygon) = False Then
                                            If geoPolygon.Area > 0 Then
                                                listProjectPolygon.Add(geoPolygon)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем все кадастровые участки на чертеже
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim arrayTbl As String(,) = Nothing
                ReDim Preserve arrayTbl(1, 0)
                arrayTbl(0, 0) = "LAND_CADASTRAL_PARCEL"
                Dim listCadastralPolygon As Dictionary(Of String, Polygon) = New Dictionary(Of String, Polygon)
                For Each id As ObjectId In btr
                    If id.IsNull = False Then
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                        If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                            Dim arrayTablesPS As String() = {}
                            Dim findTBL As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, arrayTablesPS, arrayTbl)
                            If arrayTablesPS.Length > 0 Then
                                Dim poly As Polyline = ent
                                Dim parcelPolygon As Polygon = netTopo.CreatePolygon(poly)
                                If IsNothing(parcelPolygon) = False Then
                                    listCadastralPolygon.Add(poly.Handle.ToString, parcelPolygon)
                                End If
                            End If
                        End If
                    End If
                Next
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If listCadastralPolygon.Count > 0 And listProjectPolygon.Count > 0 Then
                    'временный каталог для автосохранения таблицы
                    Dim TempDir As String = ""
                    Try
                        TempDir = System.IO.Path.GetTempPath()
                    Catch ex As System.Security.SecurityException
                        TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Word документ
                    Dim WApp As Word.Application = Nothing
                    Dim WDoc As Word.Document = Nothing
                    Dim boolWDVisibl As Boolean = False
                    Try
                        Dim PutchDoc As String = PutcTableDoc & "ПМТ-Ведомость изменяемых ЗУ.docx"
                        Dim rng As Random = New Random
                        Dim Numbrng As Integer = rng.Next
                        Dim DestDoc As String = TempDir & "ПМТ-Ведомость изменяемых ЗУ_" & Numbrng & ".docx"
                        Try
                            If File.Exists(PutchDoc) = True Then
                                File.Copy(PutchDoc, DestDoc)
                            End If
                        Catch ex As UnauthorizedAccessException
                        Catch ex As ArgumentException
                        Catch ex As PathTooLongException
                        Catch ex As DirectoryNotFoundException
                        Catch ex As FileNotFoundException
                        Catch ex As IOException
                        Catch ex As NotSupportedException
                        End Try
                        If File.Exists(DestDoc) = True Then
                            WApp = CreateObject("Word.Application")
                            WDoc = WApp.Documents.Open(DestDoc)
                        Else
                            MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                            WApp.Quit()
                            WApp = Nothing
                            Exit Sub
                        End If
                        WApp.Visible = True
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'начинаем перебирать участки
                        For i As Integer = 0 To listCadastralPolygon.Count - 1
                            Dim hgStr As String = listCadastralPolygon.ElementAt(i).Key
                            Dim cadastrPolugon As Polygon = listCadastralPolygon.ElementAt(i).Value
                            'проверяем пересечения
                            Dim totalAreaIntersect As Double = 0
                            For j As Integer = 0 To listProjectPolygon.Count - 1
                                Dim projectPolygon As Polygon = listProjectPolygon(j)
                                Dim boolIntersect As Boolean = netTopo.DoPolygonsIntersect(cadastrPolugon, projectPolygon)
                                If boolIntersect = True Then
                                    Dim geoIntersect As Geometry = netTopo.GetIntersection(cadastrPolugon, projectPolygon)
                                    If geoIntersect.Area > 0 Then
                                        totalAreaIntersect += geoIntersect.Area
                                    End If
                                End If
                            Next j
                            If totalAreaIntersect >= 1 Then
                                Dim idParcel As ObjectId = AcadFunction.FuncReturnObjectID(hgStr) 'id объекта
                                Dim cadastralParcels As Polyline = tr.GetObject(idParcel, OpenMode.ForRead)
                                Dim arrayData As String(,) = {}
                                Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(cadastralParcels, arrayData, ArrayTablesXML)
                                If arrayData.GetUpperBound(1) > 0 Then
                                    Dim cadNumber As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cad_number", 1, 2)
                                    Dim areaParcel As Double = CDbl(MathFunction.FuncFindValueToArray2d(arrayData, "Area", 1, 2))
                                    Dim Adress As String = MathFunction.FuncFindValueToArray2d(arrayData, "Adress", 1, 2)
                                    Dim Category As String = MathFunction.FuncFindValueToArray2d(arrayData, "Category", 1, 2)
                                    Dim UtilizationByDoc As String = MathFunction.FuncFindValueToArray2d(arrayData, "UtilizationByDoc", 1, 2)
                                    Dim Utilization As String = MathFunction.FuncFindValueToArray2d(arrayData, "Utilization", 1, 2)
                                    Dim deltaArea As Double = Math.Round(areaParcel - totalAreaIntersect, 0)
                                    For Each tabl As Word.Table In WDoc.Tables
                                        Dim countTableword As Integer = tabl.Rows.Count
                                        tabl.Rows.Add()
                                        tabl.Cell(countTableword, 1).Range.Text = cadNumber
                                        tabl.Cell(countTableword, 2).Range.Text = Adress
                                        If Utilization.Trim.Length > 0 Then
                                            tabl.Cell(countTableword, 3).Range.Text = Utilization
                                        Else
                                            tabl.Cell(countTableword, 3).Range.Text = UtilizationByDoc
                                        End If
                                        tabl.Cell(countTableword, 4).Range.Text = Category
                                        tabl.Cell(countTableword, 5).Range.Text = areaParcel
                                        tabl.Cell(countTableword, 6).Range.Text = deltaArea
                                        Exit For
                                    Next
                                End If
                            End If
                        Next i
                        WDoc = Nothing
                        WApp = Nothing
                    Catch ex As System.Exception
                        MsgBox("Не удалось записать данные в выбранную таблицу!!!" & "Номер сервитута: " & numberServitut & vbLf & ex.Message)
                        WDoc = Nothing
                        WApp = Nothing
                        Exit Sub
                    End Try
                End If
            End Using
        Catch
        End Try
    End Sub
    '===========================================================================================================
    'отчет по реконструируемым коммуникациям (ведомость инженерных сетей подлежащих реконструкции)
    <CommandMethod("ListProjectEngСommunications")>
    Public Sub ListProjectEngСommunications()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        Dim numberServitut As String = 0
        ArrayTablesXML(0, 0) = "PROJECT_LINE"
        ArrayTablesXML(1, 0) = ""
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                '=========================================================================================================
                'выбираем полилинии
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim polyPipe As Polyline = acEnt
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        Dim arrayData As String(,) = {}
                                        Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, arrayData, ArrayTablesXML)
                                        If arrayData.GetUpperBound(1) > 0 Then
                                            Dim numberLine As String = MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2)
                                            If numberLine.Trim.Length = 0 Then
                                                numberLine = 1
                                            End If
                                            Dim name As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 1, 2)
                                            Dim KeyParameters As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 1, 2)
                                            Dim TypeWork As String = MathFunction.FuncFindValueToArray2d(arrayData, "TypeWork", 1, 2)
                                            Dim Patronymic As String = MathFunction.FuncFindValueToArray2d(arrayData, "Patronymic", 1, 2)
                                            If TypeWork Like "Реконструкция" Then
                                                ReDim Preserve ArrayAllPoligon(4, countArrayAllPoligon)
                                                ArrayAllPoligon(0, countArrayAllPoligon) = numberLine
                                                ArrayAllPoligon(1, countArrayAllPoligon) = name
                                                ArrayAllPoligon(2, countArrayAllPoligon) = KeyParameters
                                                ArrayAllPoligon(3, countArrayAllPoligon) = Patronymic
                                                ArrayAllPoligon(4, countArrayAllPoligon) = Math.Round(polyPipe.Length, 2)
                                                countArrayAllPoligon += 1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось выбрать объекты подлежащие реконструкции. (Набор данных - PROJECT_LINE, значение поля TypeWork - Реконструкция!!!")
                    Exit Sub
                End If
                'сортируем сервитуты по номеру
                If ArrayAllPoligon.GetUpperBound(1) > 1 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
                End If
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тело программы
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                Try
                    Dim PutchDoc As String = PutcTableDoc & "ППТ-Ведомость инженерных коммуникаций (реконструкция).docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "ППТ-Ведомость инженерных коммуникаций (реконструкция)_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End If
                    WApp.Visible = True
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                        Dim CondLine As String = ArrayAllPoligon(0, i) 'номер объекта
                        Dim Name As String = ArrayAllPoligon(1, i) 'имя объекта
                        Dim KeyParam As String = ArrayAllPoligon(2, i) 'характеристика
                        Dim lenght As String = ArrayAllPoligon(4, i) 'длина переустройства
                        Dim Patron As String = ArrayAllPoligon(3, i) 'правообладатель
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            tabl.Cell(countTableword, 1).Range.Text = CondLine
                            tabl.Cell(countTableword, 2).Range.Text = Name
                            tabl.Cell(countTableword, 3).Range.Text = KeyParam
                            tabl.Cell(countTableword, 4).Range.Text = lenght
                            tabl.Cell(countTableword, 5).Range.Text = Patron
                            Exit For
                        Next
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                Catch ex As System.Exception
                    MsgBox("Не удалось записать данные в выбранную таблицу!!!" & vbLf & ex.Message)
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub
    '===========================================================================================================
    'отчет по демонтируемым коммуникациям 
    <CommandMethod("ListDistmatingСommunications")>
    Public Sub ListDistmatingСommunications()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        Dim numberServitut As String = 0
        ArrayTablesXML(0, 0) = "PROJECT_DISMANTLING"
        ArrayTablesXML(1, 0) = ""
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                '=========================================================================================================
                'выбираем полилинии
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim polyPipe As Polyline = acEnt
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        Dim arrayData As String(,) = {}
                                        Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, arrayData, ArrayTablesXML)
                                        If arrayData.GetUpperBound(1) > 0 Then
                                            Dim name As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 1, 2)
                                            Dim cad_number As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cad_number", 1, 2)
                                            Dim KeyParameters As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 1, 2)
                                            Dim TypeWork As String = MathFunction.FuncFindValueToArray2d(arrayData, "TypesWork", 1, 2)
                                            Dim Patronymic As String = MathFunction.FuncFindValueToArray2d(arrayData, "Patronymic", 1, 2)
                                            ReDim Preserve ArrayAllPoligon(5, countArrayAllPoligon)
                                            ArrayAllPoligon(0, countArrayAllPoligon) = name
                                            ArrayAllPoligon(1, countArrayAllPoligon) = cad_number
                                            ArrayAllPoligon(2, countArrayAllPoligon) = KeyParameters
                                            ArrayAllPoligon(3, countArrayAllPoligon) = Patronymic
                                            ArrayAllPoligon(4, countArrayAllPoligon) = TypeWork
                                            ArrayAllPoligon(5, countArrayAllPoligon) = Math.Round(polyPipe.Length, 2)
                                            countArrayAllPoligon += 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось выбрать объекты подлежащие реконструкции. (Набор данных - PROJECT_LINE, значение поля TypeWork - Реконструкция!!!")
                    Exit Sub
                End If
                'сортируем сервитуты по номеру
                If ArrayAllPoligon.GetUpperBound(1) > 1 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
                End If
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тело программы
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                Try
                    Dim PutchDoc As String = PutcTableDoc & "ППТ-Ведомость демонтажа инженерных коммуникаций.docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "ППТ-Ведомость демонтажа инженерных коммуникаций_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End If
                    WApp.Visible = True
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                        Dim Name As String = ArrayAllPoligon(0, i) 'имя объекта
                        Dim cad_number As String = ArrayAllPoligon(1, i) 'имя объекта
                        Dim KeyParam As String = ArrayAllPoligon(2, i) 'характеристика
                        Dim Patron As String = ArrayAllPoligon(3, i) 'правообладатель
                        Dim TypeWork As String = ArrayAllPoligon(4, i) 'правообладатель
                        Dim lenght As String = ArrayAllPoligon(5, i) 'длина переустройства
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            tabl.Cell(countTableword, 1).Range.Text = Name
                            tabl.Cell(countTableword, 2).Range.Text = KeyParam
                            tabl.Cell(countTableword, 3).Range.Text = cad_number
                            tabl.Cell(countTableword, 4).Range.Text = lenght
                            tabl.Cell(countTableword, 5).Range.Text = Patron
                            tabl.Cell(countTableword, 6).Range.Text = TypeWork
                            Exit For
                        Next
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                Catch ex As System.Exception
                    MsgBox("Не удалось записать данные в выбранную таблицу!!!" & vbLf & ex.Message)
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub
    '===========================================================================================================
    'отчет по демонтируемым сооружениям
    <CommandMethod("ListDistmatingBuidings")>
    Public Sub ListDistmatingBuidings()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        Dim numberServitut As String = 0
        ArrayTablesXML(0, 0) = "PROJECT_DISMANTLING"
        ArrayTablesXML(1, 0) = ""
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам
        Dim PutcTableDoc As String = "d:\Civil3DToolsUtility\TemplateTables\Territory_planning\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Territory_planning\") Then
            PutcTableDoc = parentPatch2 & "\TemplateTables\Territory_planning\"
        Else
            If IO.Directory.Exists(PutcTableDoc) = False Then
                MsgBox("Шаблон таблицы не найден!!!")
                Exit Sub
            End If
        End If
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                '=========================================================================================================
                'выбираем полилинии
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                If TypeOf acEnt Is Autodesk.AutoCAD.DatabaseServices.Polyline Then
                                    Dim polyPipe As Polyline = acEnt
                                    Dim ArrayNameTableOD As String() = Nothing
                                    Dim boolTblPS As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        Dim arrayData As String(,) = {}
                                        Dim boolFindData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, arrayData, ArrayTablesXML)
                                        If arrayData.GetUpperBound(1) > 0 Then
                                            Dim CondNumber As String = MathFunction.FuncFindValueToArray2d(arrayData, "ConditionalNumber", 1, 2)
                                            Dim name As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 1, 2)
                                            Dim cad_number As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cad_number", 1, 2)
                                            Dim address As String = MathFunction.FuncFindValueToArray2d(arrayData, "Address", 1, 2)
                                            Dim KeyParameters As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 1, 2)
                                            Dim Patronymic As String = MathFunction.FuncFindValueToArray2d(arrayData, "Patronymic", 1, 2)
                                            ReDim Preserve ArrayAllPoligon(5, countArrayAllPoligon)
                                            ArrayAllPoligon(0, countArrayAllPoligon) = CondNumber
                                            ArrayAllPoligon(1, countArrayAllPoligon) = name
                                            ArrayAllPoligon(2, countArrayAllPoligon) = cad_number
                                            ArrayAllPoligon(3, countArrayAllPoligon) = KeyParameters
                                            ArrayAllPoligon(4, countArrayAllPoligon) = address
                                            ArrayAllPoligon(5, countArrayAllPoligon) = Patronymic
                                            countArrayAllPoligon += 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось выбрать объекты подлежащие демонтажу. (Набор данных - PROJECT_LINE)!!!")
                    Exit Sub
                End If
                'сортируем сервитуты по номеру
                If ArrayAllPoligon.GetUpperBound(1) > 1 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayAllPoligon, 0)
                End If
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'тело программы
                'временный каталог для автосохранения таблицы
                Dim TempDir As String = ""
                Try
                    TempDir = System.IO.Path.GetTempPath()
                Catch ex As System.Security.SecurityException
                    TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim boolWDVisibl As Boolean = False
                Try
                    Dim PutchDoc As String = PutcTableDoc & "ППТ-Ведомость демонтажа зданий.docx"
                    Dim rng As Random = New Random
                    Dim Numbrng As Integer = rng.Next
                    Dim DestDoc As String = TempDir & "ППТ-Ведомость демонтажа зданий_" & Numbrng & ".docx"
                    Try
                        If File.Exists(PutchDoc) = True Then
                            File.Copy(PutchDoc, DestDoc)
                        End If
                    Catch ex As UnauthorizedAccessException
                    Catch ex As ArgumentException
                    Catch ex As PathTooLongException
                    Catch ex As DirectoryNotFoundException
                    Catch ex As FileNotFoundException
                    Catch ex As IOException
                    Catch ex As NotSupportedException
                    End Try
                    If File.Exists(DestDoc) = True Then
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(DestDoc)
                    Else
                        MsgBox("Не удалось открыть шаблон таблицы. Макро прервано!!!")
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End If
                    WApp.Visible = True
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                        Dim CondNumber As String = ArrayAllPoligon(0, i) 'номер на плане
                        Dim Name As String = ArrayAllPoligon(1, i) 'имя объекта
                        Dim cad_number As String = ArrayAllPoligon(2, i) 'кад.номер
                        Dim KeyParam As String = ArrayAllPoligon(3, i) 'характеристика
                        Dim Patron As String = ArrayAllPoligon(5, i) 'правообладатель
                        Dim Adress As String = ArrayAllPoligon(4, i) 'адрес
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            tabl.Cell(countTableword, 1).Range.Text = CondNumber
                            tabl.Cell(countTableword, 2).Range.Text = Name
                            tabl.Cell(countTableword, 3).Range.Text = KeyParam
                            tabl.Cell(countTableword, 4).Range.Text = cad_number
                            tabl.Cell(countTableword, 5).Range.Text = Adress
                            tabl.Cell(countTableword, 6).Range.Text = Patron
                            Exit For
                        Next
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                Catch ex As System.Exception
                    MsgBox("Не удалось записать данные в выбранную таблицу!!!" & vbLf & ex.Message)
                    WDoc = Nothing
                    WApp = Nothing
                    Exit Sub
                End Try
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub
End Class
