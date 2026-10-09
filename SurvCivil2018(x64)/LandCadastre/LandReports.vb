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
Imports Microsoft.Office.Interop.Excel
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class LandReports
    '=========================================================================================================
    'Экспликация земельных участков (для отчета по инвентаризации земель)
    <CommandMethod("ListLandParcelsInventary")>
    Public Sub ListLandParcelsInventary()
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
        Dim PutcTableDoc As String = parentPatch2 & "\TemplateTables\Land_Cadastral\"
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
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем все кадастровые участки на чертеже
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'массивы с участками
            Dim CountArrayCadastralParcels As Integer = 0
            Dim ArrayCadastralParcels As String(,) = Nothing
            'массивы с кварталами
            Dim CountArrayCadastralDivision As Integer = 0
            Dim ArrayCadastralDivision As String(,) = Nothing
            'массивы со зданиями
            Dim CountArrayCadastralBuilds As Integer = 0
            Dim ArrayArrayCadastralBuilds As String(,) = Nothing
            'проектные участки постоянного отвода
            Dim CountArrayProjectParcel As Integer = 0
            Dim ArrayProjectParcel As String(,) = Nothing
            'проектные участки временного отвода
            Dim CountArrayTemporaryParcel As Integer = 0
            Dim ArrayTemporaryParcel As String(,) = Nothing
            'массив с островками
            Dim CountArrayProjectParcelOstrov As Integer = 0
            Dim ArrayArrayProjectParcelOstrov As String(,) = Nothing
            'забираем в массив весь кадастр
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    Dim idLayer As ObjectId = ent.LayerId
                    Dim acLyrTblRec As LayerTableRecord = tr.GetObject(idLayer, OpenMode.ForRead)
                    If acLyrTblRec.IsOff = True Then
                        Continue For
                    End If
                    If acLyrTblRec.IsFrozen = True Then
                        Continue For
                    End If
                    If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                        Dim nameOD As String = ""
                        Dim boolFindTablePSParcel As Boolean = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_PARCEL")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayCadastralParcels(6, CountArrayCadastralParcels)
                            ArrayCadastralParcels(0, CountArrayCadastralParcels) = ent.Handle.ToString
                            ArrayCadastralParcels(1, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.X
                            ArrayCadastralParcels(2, CountArrayCadastralParcels) = ent.GeometricExtents.MinPoint.Y
                            ArrayCadastralParcels(3, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.X
                            ArrayCadastralParcels(4, CountArrayCadastralParcels) = ent.GeometricExtents.MaxPoint.Y
                            ArrayCadastralParcels(5, CountArrayCadastralParcels) = FuncPropertySet.FuncReadValueSetProperties(ent, "Cad_number", "LAND_CADASTRAL_PARCEL").Trim
                            ArrayCadastralParcels(6, CountArrayCadastralParcels) = "" 'сюда будут записаны все сооружения находящиеся на данном участке
                            CountArrayCadastralParcels += 1
                            GoTo Line3
                        End If
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_DIVISION")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayCadastralDivision(6, CountArrayCadastralDivision)
                            ArrayCadastralDivision(0, CountArrayCadastralDivision) = ent.Handle.ToString
                            ArrayCadastralDivision(1, CountArrayCadastralDivision) = ent.GeometricExtents.MinPoint.X
                            ArrayCadastralDivision(2, CountArrayCadastralDivision) = ent.GeometricExtents.MinPoint.Y
                            ArrayCadastralDivision(3, CountArrayCadastralDivision) = ent.GeometricExtents.MaxPoint.X
                            ArrayCadastralDivision(4, CountArrayCadastralDivision) = ent.GeometricExtents.MaxPoint.Y
                            ArrayCadastralDivision(5, CountArrayCadastralDivision) = FuncPropertySet.FuncReadValueSetProperties(ent, "Cadastral_number", "LAND_CADASTRAL_DIVISION")
                            ArrayCadastralDivision(6, CountArrayCadastralDivision) = "" 'сюда будут записаны все сооружения находящиеся на данном участке
                            CountArrayCadastralDivision += 1
                            GoTo Line3
                        End If
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_BUILDING")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayArrayCadastralBuilds(5, CountArrayCadastralBuilds)
                            ArrayArrayCadastralBuilds(0, CountArrayCadastralBuilds) = ent.Handle.ToString
                            ArrayArrayCadastralBuilds(1, CountArrayCadastralBuilds) = ent.GeometricExtents.MinPoint.X
                            ArrayArrayCadastralBuilds(2, CountArrayCadastralBuilds) = ent.GeometricExtents.MinPoint.Y
                            ArrayArrayCadastralBuilds(3, CountArrayCadastralBuilds) = ent.GeometricExtents.MaxPoint.X
                            ArrayArrayCadastralBuilds(4, CountArrayCadastralBuilds) = ent.GeometricExtents.MaxPoint.Y
                            ArrayArrayCadastralBuilds(5, CountArrayCadastralBuilds) = ""
                            CountArrayCadastralBuilds += 1
                            GoTo Line3
                        End If
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "LAND_CADASTRAL_CONSTRUCTION")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayArrayCadastralBuilds(5, CountArrayCadastralBuilds)
                            ArrayArrayCadastralBuilds(0, CountArrayCadastralBuilds) = ent.Handle.ToString
                            ArrayArrayCadastralBuilds(1, CountArrayCadastralBuilds) = ent.GeometricExtents.MinPoint.X
                            ArrayArrayCadastralBuilds(2, CountArrayCadastralBuilds) = ent.GeometricExtents.MinPoint.Y
                            ArrayArrayCadastralBuilds(3, CountArrayCadastralBuilds) = ent.GeometricExtents.MaxPoint.X
                            ArrayArrayCadastralBuilds(4, CountArrayCadastralBuilds) = ent.GeometricExtents.MaxPoint.Y
                            ArrayArrayCadastralBuilds(5, CountArrayCadastralBuilds) = ""
                            CountArrayCadastralBuilds += 1
                            GoTo Line3
                        End If
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "PROJECT_PARCEL")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayProjectParcel(3, CountArrayProjectParcel)
                            ArrayProjectParcel(0, CountArrayProjectParcel) = ent.Handle.ToString
                            ArrayProjectParcel(1, CountArrayProjectParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "GLOBALID", NameProjectOD)
                            ArrayProjectParcel(2, CountArrayProjectParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "LOCALID", NameProjectOD)
                            Dim TempPline As Polyline = ent
                            ArrayProjectParcel(3, CountArrayProjectParcel) = Math.Round(TempPline.Area, 2)
                            CountArrayProjectParcel += 1
                            GoTo Line3
                        End If
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "TEMPORARY_PARCEL")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayTemporaryParcel(3, CountArrayTemporaryParcel)
                            ArrayTemporaryParcel(0, CountArrayTemporaryParcel) = ent.Handle.ToString
                            ArrayTemporaryParcel(1, CountArrayTemporaryParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "GLOBALID", NameTemporaryOD)
                            ArrayTemporaryParcel(2, CountArrayTemporaryParcel) = FuncPropertySet.FuncReadValueSetProperties(ent, "LOCALID", NameTemporaryOD)
                            Dim TempPline As Polyline = ent
                            ArrayTemporaryParcel(3, CountArrayTemporaryParcel) = Math.Round(TempPline.Area, 2)
                            CountArrayTemporaryParcel += 1
                            GoTo Line3
                        End If
Line3:
                        boolFindTablePSParcel = False
                        boolFindTablePSParcel = FuncPropertySet.FuncFindPSetTable(ent, "CADASTRAL_OSTROV")
                        If boolFindTablePSParcel = True Then
                            ReDim Preserve ArrayArrayProjectParcelOstrov(3, CountArrayProjectParcelOstrov)
                            ArrayArrayProjectParcelOstrov(0, CountArrayProjectParcelOstrov) = ent.Handle.ToString
                            ArrayArrayProjectParcelOstrov(1, CountArrayProjectParcelOstrov) = FuncPropertySet.FuncReadValueSetProperties(ent, "IDKOUNTER", "CADASTRAL_OSTROV")
                            ArrayArrayProjectParcelOstrov(2, CountArrayProjectParcelOstrov) = FuncPropertySet.FuncReadValueSetProperties(ent, "ConditionalNumber", "CADASTRAL_OSTROV")
                            Dim TempPline As Polyline = ent
                            ArrayArrayProjectParcelOstrov(3, CountArrayProjectParcelOstrov) = Math.Round(TempPline.Area, 2)
                            CountArrayProjectParcelOstrov += 1
                            Continue For
                        End If
                    End If
                End If
            Next
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'фильтруем точки
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
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '=========================================================================================================
            'выбираем участки постоянного отвода
            '0-ПК или номера,1-Handle ЗУ,2-Номер ЗУ на карте,3-KN,4-Площадь,5-Адрес,6-Категория,7-Разрешенное использование,8-Вид права,9-Правообладатель,10-Проектная площадь
            Dim ArrayBuildParcels As String(,) = Nothing
            Dim CountBuildParcels As Integer = 0
            If IsArray(ArrayProjectParcel) = True Then
                For i1 As Integer = 0 To ArrayProjectParcel.GetUpperBound(1)
                    Try
                        Dim hgCounter As String = ArrayProjectParcel(0, i1)
                        Dim idLoc As String = ArrayProjectParcel(2, i1)
                        Dim IDCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = IDCounter.GetObject(OpenMode.ForRead)
                        Dim ProjectParcelPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt

                        Dim CouditionalNumber As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "ConditionalNumber", NameProjectOD)
                        Dim CadNumberProject As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number", NameProjectOD)
                        Dim CadastralParcel As Polyline = Nothing
                        If CadNumberProject.Length > 2 Then
                            If IsArray(ArrayCadastralParcels) = True Then
                                For i As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                                    If ArrayCadastralParcels(5, i) Like CadNumberProject Then
                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                        CadastralParcel = idObj.GetObject(OpenMode.ForRead)
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        If IsNothing(CadastralParcel) = True Then
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'определение центра полигона 
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim PointCenterProjectParcels As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(ProjectParcelPline)
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
                                    If PointCenterProjectParcels.X >= minX And PointCenterProjectParcels.Y >= minY And PointCenterProjectParcels.X <= maxX And PointCenterProjectParcels.Y <= maxY Then
                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                        Dim TempCadastralParcel As Polyline = idObj.GetObject(OpenMode.ForRead)
                                        Dim boolCrosPointParcel As Boolean = MathFunction.FuncCrossPointPolyline(TempCadastralParcel, PointCenterProjectParcels.Convert2d(New Plane))
                                        If boolCrosPointParcel = True Then
                                            CadastralParcel = TempCadastralParcel
                                            Exit For
                                        End If
                                    End If
                                Next i
                            End If
                        End If
                        '==============================================================================================================================================
                        'смотрим нумерацию
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim StrNumberPoint As String = ""
                        Dim TempNumber As String = FuncCadastr.FuncReturnGeoPointInPolyline(GeoPointCoolection, ProjectParcelPline, 0.05)
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
                                Dim HandlePLine As String = ArrayOstrov1(1, k)
                                Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                If IsNothing(idOstrov) = True Then
                                    Continue For
                                End If
                                If idOstrov.IsNull = True Then
                                    Continue For
                                End If
                                Dim PlineObj As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
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
                        ReDim Preserve ArrayBuildParcels(12, CountBuildParcels)
                        ArrayBuildParcels(0, CountBuildParcels) = FuncCadastr.FuncFindShortNumberCadastralObject(CouditionalNumber) 'короткий номер участка
                        ArrayBuildParcels(1, CountBuildParcels) = acEnt.Handle.ToString 'Handle проектных участков
                        ArrayBuildParcels(2, CountBuildParcels) = FuncCadastr.FuncFindNumberCadastralObject(CouditionalNumber) 'обозначения на карте
                        If IsNothing(CadastralParcel) = False Then
                            ArrayBuildParcels(3, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Cad_number", "LAND_CADASTRAL_PARCEL")
                            ArrayBuildParcels(4, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Area", "LAND_CADASTRAL_PARCEL")
                            ArrayBuildParcels(5, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Adress", "LAND_CADASTRAL_PARCEL")
                            ArrayBuildParcels(6, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Category", "LAND_CADASTRAL_PARCEL")
                            Dim Utilization As String = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Utilization", "LAND_CADASTRAL_PARCEL")
                            If Utilization Like "" Then
                                Utilization = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "UtilizationByDoc", "LAND_CADASTRAL_PARCEL")
                            End If
                            ArrayBuildParcels(7, CountBuildParcels) = Utilization
                            ArrayBuildParcels(8, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Right_type", "LAND_CADASTRAL_PARCEL")
                            ArrayBuildParcels(9, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Patronymic", "LAND_CADASTRAL_PARCEL")
                            ArrayBuildParcels(10, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Encumbrance", "LAND_CADASTRAL_PARCEL")
                            'ArrayBuildParcels(11, CountBuildParcels) = GlobalID
                            ArrayBuildParcels(12, CountBuildParcels) = TempNumber
                            CountBuildParcels += 1
                        Else
                            ReDim Preserve ArrayBuildParcels(12, CountBuildParcels)
                            ArrayBuildParcels(0, CountBuildParcels) = FuncCadastr.FuncFindShortNumberCadastralObject(CouditionalNumber) 'короткий номер участка
                            ArrayBuildParcels(1, CountBuildParcels) = acEnt.Handle.ToString 'Handle проектных участков
                            ArrayBuildParcels(2, CountBuildParcels) = FuncCadastr.FuncFindNumberCadastralObject(CouditionalNumber) 'обозначения на карте
                            ArrayBuildParcels(3, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number", "PROJECT_PARCEL")
                            ArrayBuildParcels(4, CountBuildParcels) = "-"
                            ArrayBuildParcels(5, CountBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Adress", "PROJECT_PARCEL")
                            ArrayBuildParcels(6, CountBuildParcels) = "Категория земель не установлена"
                            ArrayBuildParcels(7, CountBuildParcels) = "-"
                            ArrayBuildParcels(8, CountBuildParcels) = "Неразграниченная государственная собственность"
                            ArrayBuildParcels(9, CountBuildParcels) = "-"
                            ArrayBuildParcels(10, CountBuildParcels) = "-"
                            'ArrayBuildParcels(11, CountBuildParcels) = GlobalID
                            ArrayBuildParcels(12, CountBuildParcels) = TempNumber
                            CountBuildParcels += 1
                        End If
                    Catch ex As System.Exception
                    End Try
Line1:
                Next i1
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сортируем массив
            If IsArray(ArrayBuildParcels) = True Then
                If ArrayBuildParcels.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayBuildParcels, 0)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '=========================================================================================================
            'выбираем участки временного отвода
            '0-ПК или номера,1-Handle ЗУ,2-Номер ЗУ на карте,3-KN,4-Площадь,5-Адрес,6-Категория,7-Разрешенное использование,8-Вид права,9-Правообладатель,10-Проектная площадь
            Dim ArrayTempBuildParcels As String(,) = Nothing
            Dim CountTempBuildParcels As Integer = 0
            If IsArray(ArrayTemporaryParcel) = True Then
                For i1 As Integer = 0 To ArrayTemporaryParcel.GetUpperBound(1)
                    Try
                        Dim hgCounter As String = ArrayTemporaryParcel(0, i1)
                        Dim IDCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = IDCounter.GetObject(OpenMode.ForRead)
                        Dim GlobalID As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "GLOBALID", NameTemporaryOD).Trim
                        If GlobalID.Length > 1 Then
                            If IsArray(ArrayTempBuildParcels) = True Then
                                For i2 As Integer = 0 To ArrayTempBuildParcels.GetUpperBound(1)
                                    If ArrayTempBuildParcels(11, i2) Like GlobalID Then
                                        ArrayTempBuildParcels(1, i2) = ArrayTempBuildParcels(1, i2) & " " & acEnt.Handle.ToString 'Handle проектных участков
                                        GoTo Line2
                                    End If
                                Next i2
                            End If
                        End If
                        Dim CouditionalNumber As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "ConditionalNumber", NameTemporaryOD)
                        Dim CadNumberProject As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number", NameTemporaryOD)
                        Dim CadastralParcel As Polyline = Nothing
                        If CadNumberProject.Length > 2 Then
                            If IsArray(ArrayCadastralParcels) = True Then
                                For i As Integer = 0 To ArrayCadastralParcels.GetUpperBound(1)
                                    If ArrayCadastralParcels(5, i) Like CadNumberProject Then
                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                        CadastralParcel = idObj.GetObject(OpenMode.ForRead)
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        If IsNothing(CadastralParcel) = True Then
                            Dim ProjectParcelPline As Autodesk.AutoCAD.DatabaseServices.Polyline = acEnt
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'определение центра полигона 
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim PointCenterProjectParcels As Point3d = MathFunction.FuncFindCrossCenterPointPolyline(ProjectParcelPline)
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
                                    If PointCenterProjectParcels.X >= minX And PointCenterProjectParcels.Y >= minY And PointCenterProjectParcels.X <= maxX And PointCenterProjectParcels.Y <= maxY Then
                                        Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(ArrayCadastralParcels(0, i))
                                        Dim TempCadastralParcel As Polyline = idObj.GetObject(OpenMode.ForRead)
                                        Dim boolCrosPointParcel As Boolean = MathFunction.FuncCrossPointPolyline(TempCadastralParcel, PointCenterProjectParcels.Convert2d(New Plane))
                                        If boolCrosPointParcel = True Then
                                            CadastralParcel = TempCadastralParcel
                                            Exit For
                                        End If
                                    End If
                                Next i
                            End If
                        End If
                        ReDim Preserve ArrayTempBuildParcels(11, CountTempBuildParcels)
                        ArrayTempBuildParcels(0, CountTempBuildParcels) = FuncCadastr.FuncFindShortNumberCadastralObject(CouditionalNumber) 'короткий номер участка
                        ArrayTempBuildParcels(1, CountTempBuildParcels) = acEnt.Handle.ToString 'Handle проектных участков
                        ArrayTempBuildParcels(2, CountTempBuildParcels) = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(CouditionalNumber) 'обозначения на карте
                        If IsNothing(CadastralParcel) = False Then
                            ArrayTempBuildParcels(3, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Cad_number", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(4, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Area", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(5, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Adress", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(6, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Category", "LAND_CADASTRAL_PARCEL")
                            Dim Utilization As String = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Utilization", "LAND_CADASTRAL_PARCEL")
                            If Utilization Like "" Then
                                Utilization = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "UtilizationByDoc", "LAND_CADASTRAL_PARCEL")
                            End If
                            ArrayTempBuildParcels(7, CountTempBuildParcels) = Utilization
                            ArrayTempBuildParcels(8, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Right_type", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(9, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Patronymic", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(10, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(CadastralParcel, "Encumbrance", "LAND_CADASTRAL_PARCEL")
                            ArrayTempBuildParcels(11, CountTempBuildParcels) = GlobalID
                            CountTempBuildParcels += 1
                        Else
                            ReDim Preserve ArrayBuildParcels(11, CountTempBuildParcels)
                            ArrayTempBuildParcels(0, CountTempBuildParcels) = FuncCadastr.FuncFindShortNumberCadastralObject(CouditionalNumber) 'короткий номер участка
                            ArrayTempBuildParcels(1, CountTempBuildParcels) = acEnt.Handle.ToString 'Handle проектных участков
                            ArrayTempBuildParcels(2, CountTempBuildParcels) = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(CouditionalNumber) 'обозначения на карте
                            ArrayTempBuildParcels(3, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Cad_number", "PROJECT_PARCEL")
                            ArrayTempBuildParcels(4, CountTempBuildParcels) = "-"
                            ArrayTempBuildParcels(5, CountTempBuildParcels) = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Adress", "PROJECT_PARCEL")
                            ArrayTempBuildParcels(6, CountTempBuildParcels) = "Категория земель не установлена"
                            ArrayTempBuildParcels(7, CountTempBuildParcels) = "-"
                            ArrayTempBuildParcels(8, CountTempBuildParcels) = "Неразграниченная государственная собственность"
                            ArrayTempBuildParcels(9, CountTempBuildParcels) = "-"
                            ArrayTempBuildParcels(10, CountTempBuildParcels) = "-"
                            ArrayTempBuildParcels(11, CountTempBuildParcels) = GlobalID

                            CountTempBuildParcels += 1
                        End If
                    Catch ex As System.Exception
                    End Try
Line2:
                Next i1
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сортируем массив
            If IsArray(ArrayTempBuildParcels) = True Then
                If ArrayTempBuildParcels.GetUpperBound(1) > 0 Then
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayTempBuildParcels, 0)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выводим отчет
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'формируем участки
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Word документ
            Dim WApp As Word.Application = Nothing
            Dim WDoc As Word.Document = Nothing
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
            End Try
            Dim boolWDVisibl As Boolean = False
            Dim PutchDoc As String = PutcTableDoc & "Ведомость имущественно-правовой инвентаризации.docx"
            Dim rng As Random = New Random
            Dim Numbrng As Integer = rng.Next
            Dim DestDoc As String = TempDir & "Ведомость имущественно-правовой инвентаризации_" & Numbrng & ".docx"
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
            'WApp.Visible = True
            Dim tabl As Word.Table = Nothing
            For Each tbl As Word.Table In WDoc.Tables
                tabl = tbl
                If tabl.Rows.Count > 13 Then
                    Exit For
                End If
            Next
            If IsNothing(tabl) = True Then
                MsgBox("Не удалось найти таблицу с количеством стобцов равным 15. Макро прервано!!!")
                WDoc.Close()
                WApp.Quit()
                WApp = Nothing
            End If
            Dim countTableRow As Integer = 4 'число строк
            Dim countNumbers As Integer = 1 'порядковый номер
            If IsArray(ArrayBuildParcels) = True Then
                Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'заполняем таблицу
                    '0-ПК или номера,1-Handle ЗУ,2-Номер ЗУ на карте,3-KN,4-Площадь,5-Адрес,6-Категория,7-Разрешенное использование,8-Вид права,9-Правообладатель,10-Проектная площадь,11-Handle сооружения
                    For i As Integer = 0 To ArrayBuildParcels.GetUpperBound(1)
                        Dim ArrayHandleBuild As String() = Nothing
                        Dim CountArrayHandleBuild As Integer = 0
                        Dim AreaParcel As Double = 0
                        Dim StrHandle As String() = ArrayBuildParcels(1, i).Split
                        If IsArray(StrHandle) = True Then
                            For j As Integer = 0 To StrHandle.Length - 1
                                Dim hgCounter As String = StrHandle(j)
                                Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                                Dim PlineCounter As Polyline = tr.GetObject(idCounter, OpenMode.ForRead)
                                AreaParcel = AreaParcel + PlineCounter.Area
                                Dim localId As String = FuncPropertySet.FuncReadValueSetProperties(PlineCounter, "LOCALID", "PROJECT_PARCEL")
                                If IsArray(ArrayArrayProjectParcelOstrov) = True Then
                                    For k As Integer = 0 To ArrayArrayProjectParcelOstrov.GetUpperBound(1)
                                        If ArrayArrayProjectParcelOstrov(1, k) Like localId Then
                                            AreaParcel = AreaParcel - Val(ArrayArrayProjectParcelOstrov(3, k))
                                        End If
                                    Next k
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'смотрим пересечения сооружений
                                If IsArray(ArrayArrayCadastralBuilds) = True Then
                                    Dim Pt3dColl As Point3dCollection = New Point3dCollection
                                    For k As Integer = 0 To PlineCounter.NumberOfVertices - 1
                                        Pt3dColl.Add(PlineCounter.GetPoint3dAt(k))
                                    Next k
                                    If Pt3dColl.Count > 2 Then
                                        Dim acSSBuild As PromptSelectionResult = acDoc.Editor.SelectCrossingPolygon(Pt3dColl)
                                        If acSSBuild.Status = PromptStatus.OK Then
                                            Dim ssetEnt As SelectionSet = acSSBuild.Value
                                            If ssetEnt.Count > 0 Then
                                                For Each acSSObj As SelectedObject In ssetEnt
                                                    If Not IsDBNull(ssetEnt) Then
                                                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                                        For k As Integer = 0 To ArrayArrayCadastralBuilds.GetUpperBound(1)
                                                            Dim hgBuild As String = ArrayArrayCadastralBuilds(0, k)
                                                            If hgBuild Like acEnt.Handle.ToString Then
                                                                ReDim Preserve ArrayHandleBuild(CountArrayHandleBuild)
                                                                ArrayHandleBuild(CountArrayHandleBuild) = acEnt.Handle.ToString
                                                                CountArrayHandleBuild += 1
                                                                Exit For
                                                            End If
                                                        Next k
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '======================================================================================================
                        tabl.Rows.Add()
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'считываем установки с родительского участка (кадастрового участка или квартала)
                        Dim ConditionalParcel As String = ArrayBuildParcels(2, i) 'номер проектного участка
                        tabl.Cell(countTableRow, 1).Range.Text = ConditionalParcel
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Cad_numb As String = ArrayBuildParcels(3, i) 'кадастровый номер исходного зу
                        If IsNothing(Cad_numb) = True Then
                            Cad_numb = ""
                        ElseIf Cad_numb.Trim.Length = 0 Then
                            Cad_numb = "-"
                        End If
                        tabl.Cell(countTableRow, 2).Range.Text = Cad_numb
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim StrNumberPoint As String = ArrayBuildParcels(12, i)
                        If IsNothing(StrNumberPoint) = False Then
                            If StrNumberPoint.Trim.Length > 0 Then
                                StrNumberPoint = FuncFormatZN.FuncConvertStrNumberPoint(StrNumberPoint)
                                tabl.Cell(countTableRow, 3).Range.Text = StrNumberPoint
                            End If
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Area As String = ArrayBuildParcels(4, i)  'площадь кадастрового участка
                        If IsNothing(Area) = True Then
                            Area = "-"
                        ElseIf Area.Trim.Length = 0 Then
                            Area = "-"
                        End If
                        tabl.Cell(countTableRow, 4).Range.Text = Area & vbLf
                        tabl.Cell(countTableRow, 5).Range.Text = Math.Round(AreaParcel, 0)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Adress As String = ArrayBuildParcels(5, i) '  адрес исходного зу
                        If IsNothing(Adress) = True Then
                            Adress = "-"
                        ElseIf Adress.Trim.Length = 0 Then
                            Adress = "-"
                        End If
                        tabl.Cell(countTableRow, 6).Range.Text = Adress
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Category As String = ArrayBuildParcels(6, i)  'категория исходного зу
                        If IsNothing(Category) = True Then
                            Category = "-"
                        ElseIf Category.Trim.Length = 0 Then
                            Category = "-"
                        End If
                        tabl.Cell(countTableRow, 7).Range.Text = Category
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Utilization As String = ArrayBuildParcels(7, i)   'разрешенное использование исходного зу
                        If IsNothing(Utilization) = True Then
                            Utilization = "-"
                        ElseIf Utilization.Trim.Length = 0 Then
                            Utilization = "-"
                        End If
                        tabl.Cell(countTableRow, 8).Range.Text = Utilization
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim RightPatron As String = "-"
                        Dim Right As String = ArrayBuildParcels(8, i) 'вид права исходного зу
                        If IsNothing(Right) = True Then
                            Right = "-"
                        ElseIf Right.Trim.Length = 0 Then
                            Right = "-"
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Patronymic As String = ArrayBuildParcels(9, i)  'правообладатель исходного зу
                        If IsNothing(Patronymic) = True Then
                            Patronymic = "-"
                        ElseIf Patronymic.Trim.Length = 0 Then
                            Patronymic = "-"
                        End If
                        If Right.Trim.Length > 2 Then
                            RightPatron = Right
                        End If
                        If Patronymic.Length > 2 Then
                            If RightPatron.Length > 2 Then
                                RightPatron = RightPatron & ", " & Patronymic
                            Else
                                RightPatron = Patronymic
                            End If
                        End If
                        tabl.Cell(countTableRow, 9).Range.Text = RightPatron
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Encumbrance As String = ArrayBuildParcels(10, i)   'ограничение
                        If IsNothing(Encumbrance) = True Then
                            Encumbrance = "-"
                        ElseIf Encumbrance.Trim.Length = 0 Then
                            Encumbrance = "-"
                        End If
                        tabl.Cell(countTableRow, 10).Range.Text = Encumbrance
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If IsArray(ArrayHandleBuild) = True Then
                            Dim CountRows As Integer = ArrayHandleBuild.Length - 1
                            Dim c10 As Word.Cell = tabl.Cell(countTableRow, 11)
                            Dim c11 As Word.Cell = tabl.Cell(countTableRow, 12)
                            Dim c12 As Word.Cell = tabl.Cell(countTableRow, 13)
                            If CountRows > 0 Then
                                c10.Split(CountRows + 1, 1)
                                c11.Split(CountRows + 1, 1)
                                c12.Split(CountRows + 1, 1)
                            Else
                                CountRows = 0
                            End If
                            For k1 As Integer = 0 To ArrayHandleBuild.Length - 1
                                Dim hgBuild As String = ArrayHandleBuild(k1)
                                Dim IDBuild As ObjectId = AcadFunction.FuncReturnObjectID(hgBuild)
                                Dim EntBuild As Autodesk.AutoCAD.DatabaseServices.Entity = IDBuild.GetObject(OpenMode.ForRead)
                                Dim Cad_number As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Cad_number")
                                Dim Address As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Adress")
                                Dim KeyParam As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "KeyParameter")
                                Dim AreaBuild As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Area")
                                Dim Assign As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Assignation")
                                Dim Charact As String = ""
                                If AreaBuild.Trim.Length > 0 Then
                                    Charact = "Площадь-" & AreaBuild & "кв.м."
                                End If
                                If KeyParam.Trim.Length > 0 Then
                                    Charact = Charact & " " & KeyParam
                                End If

                                tabl.Cell(countTableRow + k1, 11).Range.Text = Cad_number
                                tabl.Cell(countTableRow + k1, 12).Range.Text = Assign & vbLf & Charact & vbLf & KeyParam.Trim
                                tabl.Cell(countTableRow + k1, 13).Range.Text = Address
                            Next k1
                            countTableRow += CountRows
                        Else
                            tabl.Cell(countTableRow, 11).Range.Text = "-"
                            tabl.Cell(countTableRow, 12).Range.Text = "-"
                            tabl.Cell(countTableRow, 13).Range.Text = "-"
                        End If
                        countTableRow += 1
                    Next i
                Catch ex As System.Exception
                End Try
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'записываем временный отвод
            If IsArray(ArrayTempBuildParcels) = True Then
                Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'заполняем таблицу
                    '0-ПК или номера,1-Handle ЗУ,2-Номер ЗУ на карте,3-KN,4-Площадь,5-Адрес,6-Категория,7-Разрешенное использование,8-Вид права,9-Правообладатель,10-Проектная площадь,11-Handle сооружения
                    For i As Integer = 0 To ArrayTempBuildParcels.GetUpperBound(1)
                        Dim ArrayHandleBuild As String() = Nothing
                        Dim CountArrayHandleBuild As Integer = 0
                        Dim AreaParcel As Double = 0
                        Dim StrHandle As String() = ArrayTempBuildParcels(1, i).Split
                        If IsArray(StrHandle) = True Then
                            For j As Integer = 0 To StrHandle.Length - 1
                                Dim hgCounter As String = StrHandle(j)
                                Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(hgCounter)
                                Dim PlineCounter As Polyline = tr.GetObject(idCounter, OpenMode.ForRead)
                                AreaParcel = AreaParcel + PlineCounter.Area
                                Dim localId As String = FuncPropertySet.FuncReadValueSetProperties(PlineCounter, "LOCALID", NameProjectOD)
                                If IsArray(ArrayArrayProjectParcelOstrov) = True Then
                                    For k As Integer = 0 To ArrayArrayProjectParcelOstrov.GetUpperBound(1)
                                        If ArrayArrayProjectParcelOstrov(1, k) Like localId Then
                                            AreaParcel = AreaParcel - Val(ArrayArrayProjectParcelOstrov(3, k))
                                        End If
                                    Next k
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'смотрим пересечения сооружений
                                If IsArray(ArrayArrayCadastralBuilds) = True Then
                                    Dim Pt3dColl As Point3dCollection = New Point3dCollection
                                    For k As Integer = 0 To PlineCounter.NumberOfVertices - 1
                                        Pt3dColl.Add(PlineCounter.GetPoint3dAt(k))
                                    Next k
                                    If Pt3dColl.Count > 2 Then
                                        Dim acSSBuild As PromptSelectionResult = acDoc.Editor.SelectCrossingPolygon(Pt3dColl)
                                        If acSSBuild.Status = PromptStatus.OK Then
                                            Dim ssetEnt As SelectionSet = acSSBuild.Value
                                            If ssetEnt.Count > 0 Then
                                                For Each acSSObj As SelectedObject In ssetEnt
                                                    If Not IsDBNull(ssetEnt) Then
                                                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                                        For k As Integer = 0 To ArrayArrayCadastralBuilds.GetUpperBound(1)
                                                            Dim hgBuild As String = ArrayArrayCadastralBuilds(0, k)
                                                            If hgBuild Like acEnt.Handle.ToString Then
                                                                ReDim Preserve ArrayHandleBuild(CountArrayHandleBuild)
                                                                ArrayHandleBuild(CountArrayHandleBuild) = acEnt.Handle.ToString
                                                                CountArrayHandleBuild += 1
                                                                Exit For
                                                            End If
                                                        Next k
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '======================================================================================================
                        tabl.Rows.Add()
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'считываем установки с родительского участка (кадастрового участка или квартала)
                        Dim ConditionalParcel As String = ArrayTempBuildParcels(2, i) 'номер проектного участка
                        tabl.Cell(countTableRow, 1).Range.Text = ConditionalParcel
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Cad_numb As String = ArrayTempBuildParcels(3, i) 'кадастровый номер исходного зу
                        If IsNothing(Cad_numb) = True Then
                            Cad_numb = ""
                        ElseIf Cad_numb.Trim.Length = 0 Then
                            Cad_numb = "-"
                        End If
                        tabl.Cell(countTableRow, 2).Range.Text = Cad_numb
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Area As String = ArrayTempBuildParcels(4, i)  'площадь кадастрового участка
                        If IsNothing(Area) = True Then
                            Area = "-"
                        ElseIf Area.Trim.Length = 0 Then
                            Area = "-"
                        End If
                        tabl.Cell(countTableRow, 3).Range.Text = Area & vbLf
                        tabl.Cell(countTableRow, 4).Range.Text = Math.Round(AreaParcel, 0)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Adress As String = ArrayTempBuildParcels(5, i) '  адрес исходного зу
                        If IsNothing(Adress) = True Then
                            Adress = "-"
                        ElseIf Adress.Trim.Length = 0 Then
                            Adress = "-"
                        End If
                        tabl.Cell(countTableRow, 5).Range.Text = Adress
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Category As String = ArrayTempBuildParcels(6, i)  'категория исходного зу
                        If IsNothing(Category) = True Then
                            Category = "-"
                        ElseIf Category.Trim.Length = 0 Then
                            Category = "-"
                        End If
                        tabl.Cell(countTableRow, 6).Range.Text = Category
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Utilization As String = ArrayTempBuildParcels(7, i)   'разрешенное использование исходного зу
                        If IsNothing(Utilization) = True Then
                            Utilization = "-"
                        ElseIf Utilization.Trim.Length = 0 Then
                            Utilization = "-"
                        End If
                        tabl.Cell(countTableRow, 7).Range.Text = Utilization
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim RightPatron As String = "-"
                        Dim Right As String = ArrayTempBuildParcels(8, i) 'вид права исходного зу
                        If IsNothing(Right) = True Then
                            Right = "-"
                        ElseIf Right.Trim.Length = 0 Then
                            Right = "-"
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Patronymic As String = ArrayTempBuildParcels(9, i)  'правообладатель исходного зу
                        If IsNothing(Patronymic) = True Then
                            Patronymic = "-"
                        ElseIf Patronymic.Trim.Length = 0 Then
                            Patronymic = "-"
                        End If
                        If Right.Trim.Length > 2 Then
                            RightPatron = Right
                        End If
                        If Patronymic.Length > 2 Then
                            If RightPatron.Length > 2 Then
                                RightPatron = RightPatron & ", " & Patronymic
                            Else
                                RightPatron = Patronymic
                            End If
                        End If
                        tabl.Cell(countTableRow, 8).Range.Text = RightPatron
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim Encumbrance As String = ArrayTempBuildParcels(10, i)   'ограничение
                        If IsNothing(Encumbrance) = True Then
                            Encumbrance = "-"
                        ElseIf Encumbrance.Trim.Length = 0 Then
                            Encumbrance = "-"
                        End If
                        tabl.Cell(countTableRow, 9).Range.Text = Encumbrance
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        If IsArray(ArrayHandleBuild) = True Then
                            Dim CountRows As Integer = ArrayHandleBuild.Length - 1
                            Dim c10 As Word.Cell = tabl.Cell(countTableRow, 10)
                            Dim c11 As Word.Cell = tabl.Cell(countTableRow, 11)
                            Dim c12 As Word.Cell = tabl.Cell(countTableRow, 12)
                            Dim c13 As Word.Cell = tabl.Cell(countTableRow, 13)
                            If CountRows > 0 Then
                                c10.Split(CountRows + 1, 1)
                                c11.Split(CountRows + 1, 1)
                                c12.Split(CountRows + 1, 1)
                                c13.Split(CountRows + 1, 1)
                            Else
                                CountRows = 0
                            End If
                            For k1 As Integer = 0 To ArrayHandleBuild.Length - 1
                                Dim hgBuild As String = ArrayHandleBuild(k1)
                                Dim IDBuild As ObjectId = AcadFunction.FuncReturnObjectID(hgBuild)
                                Dim EntBuild As Autodesk.AutoCAD.DatabaseServices.Entity = IDBuild.GetObject(OpenMode.ForRead)
                                Dim Cad_number As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Cad_number")
                                Dim Address As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Adress")
                                Dim KeyParam As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "KeyParameter")
                                Dim AreaBuild As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Area")
                                Dim Assign As String = FuncPropertySet.FuncReadValueSetProperties(EntBuild, "Assignation")
                                tabl.Cell(countTableRow + k1, 10).Range.Text = Cad_number
                                tabl.Cell(countTableRow + k1, 11).Range.Text = Assign
                                tabl.Cell(countTableRow + k1, 12).Range.Text = Address
                                Dim Charact As String = ""
                                If AreaBuild.Trim.Length > 0 Then
                                    Charact = "Площадь-" & AreaBuild & "кв.м."
                                End If
                                If KeyParam.Trim.Length > 0 Then
                                    Charact = Charact & " " & KeyParam
                                End If
                                tabl.Cell(countTableRow + k1, 13).Range.Text = KeyParam.Trim
                            Next k1
                            countTableRow += CountRows
                        Else
                            tabl.Cell(countTableRow, 10).Range.Text = "-"
                            tabl.Cell(countTableRow, 11).Range.Text = "-"
                            tabl.Cell(countTableRow, 12).Range.Text = "-"
                            tabl.Cell(countTableRow, 13).Range.Text = "-"
                        End If
                        countTableRow += 1
                    Next i

                Catch ex As System.Exception
                End Try
            End If
            tr.Commit()
        End Using
    End Sub
    '===========================================================================================================
    'таблица участков в екселе для схемы использования территории
    <CommandMethod("ListLandParcelsInventaryPPT")>
    Public Sub ListLandParcelsInventaryPPT()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        Dim PutcTableDoc As String = parentPatch2 & "\TemplateTables\Land_Cadastral\"

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        ReDim ArrayTablesXML(1, 0)
        ArrayTablesXML(0, 0) = "LAND_CADASTRAL_PARCEL"
        ArrayTablesXML(1, 0) = "Кадастровые участки"
        Dim arrayParcels As String(,) = Nothing
        Dim countArrayParcels As Integer = 0
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim boundPline As Polyline = Nothing
            'выбираем полосу отвода
            Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите полилинию для поиска участков:")
            If ptResult.Status = PromptStatus.OK Then
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                If TypeOf ent Is Polyline Then
                    boundPline = ent
                Else
                    Dim msRez As MsgBoxResult = MsgBox("Зона не выбрана. Продолжить?", Buttons:=MsgBoxStyle.OkCancel)
                    If msRez = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If
            End If
            '=========================================================================================================
            'выбираем участки
            Dim ptOpts As PromptSelectionOptions
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Polyline Then
                                Dim pLineParcel As Polyline = acEnt
                                Dim arraydata As String(,) = Nothing
                                Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(pLineParcel, arraydata, ArrayTablesXML)
                                If IsArray(arraydata) = True Then
                                    Dim tempCollection As DBObjectCollection = New DBObjectCollection()
                                    tempCollection.Add(boundPline)
                                    Dim parentRegionCollection As DBObjectCollection = New DBObjectCollection()
                                    parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                    Dim parentRegion As Region = New Region()
                                    If parentRegionCollection.Count > 0 Then
                                        parentRegion = parentRegionCollection.Item(0)
                                    Else
                                        Continue For
                                    End If

                                    Dim childRegion As Region = New Region()
                                    If pLineParcel.Area > 0 Then
                                        Dim tempCollection2 As DBObjectCollection = New DBObjectCollection()
                                        tempCollection2.Add(pLineParcel)
                                        Dim shildRegionCollection As DBObjectCollection = New DBObjectCollection()
                                        shildRegionCollection = Region.CreateFromCurves(tempCollection2)
                                        If shildRegionCollection.Count > 0 Then
                                            childRegion = shildRegionCollection.Item(0)
                                        Else
                                            Continue For
                                        End If
                                    End If

                                    childRegion.BooleanOperation(BooleanOperationType.BoolIntersect, parentRegion)
                                    Dim newArea As Double = childRegion.Area
                                    If newArea > 0 Then
                                        ReDim Preserve arrayParcels(7, countArrayParcels)
                                        For i As Integer = 0 To arraydata.GetUpperBound(1)
                                            Dim NField As String = arraydata(1, i)
                                            Dim Zn As String = arraydata(2, i)
                                            If Zn.Trim.Length = 0 Then Zn = "Нет данных"
                                            If NField Like "Cad_number" Then
                                                arrayParcels(0, countArrayParcels) = Zn
                                            ElseIf NField Like "Adress" Then
                                                arrayParcels(1, countArrayParcels) = Zn
                                            ElseIf NField Like "Area" Then
                                                arrayParcels(2, countArrayParcels) = Zn
                                            ElseIf NField Like "Category" Then
                                                arrayParcels(3, countArrayParcels) = Zn
                                            ElseIf NField Like "UtilizationByDoc" Then
                                                arrayParcels(4, countArrayParcels) = Zn
                                            ElseIf NField Like "Right" Then
                                                arrayParcels(5, countArrayParcels) = Zn
                                            ElseIf NField Like "Patronymic" Then
                                                arrayParcels(6, countArrayParcels) = Zn
                                            ElseIf NField Like "Encumbrance" Then
                                                arrayParcels(7, countArrayParcels) = Zn
                                            End If
                                        Next i
                                        countArrayParcels += 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End Using
        If IsArray(arrayParcels) = True Then
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Try
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.ActiveSheet
                Dim cell1 As Excel.Range 'первая ячейка в строке
                cell1 = xlSheets.Cells(1, 1)
                cell1.Value = "Кадастровый номер"

                cell1 = xlSheets.Cells(1, 2)
                cell1.Value = "Адрес"

                cell1 = xlSheets.Cells(1, 3)
                cell1.Value = "Площадь, кв.м."

                cell1 = xlSheets.Cells(1, 4)
                cell1.Value = "Категория земель"

                cell1 = xlSheets.Cells(1, 5) 'ссылка на первую ячейку
                cell1.Value = "Вид разрешенного использования" 'условный номер участка

                cell1 = xlSheets.Cells(1, 6)
                cell1.Value = "Вид права"

                cell1 = xlSheets.Cells(1, 7)
                cell1.Value = "Правообладатель" 'условный номер участка

                cell1 = xlSheets.Cells(1, 8) 'ссылка на первую ячейку
                cell1.Value = "Обременения" 'условный номер участка

                Dim countRows As Integer = 2
                For i As Integer = 0 To arrayParcels.GetUpperBound(1)
                    For j As Integer = 0 To arrayParcels.GetUpperBound(0)
                        Dim valStr As String = arrayParcels(j, i)
                        If IsNothing(valStr) = False Then
                            cell1 = xlSheets.Cells(countRows, j + 1) 'ссылка на первую ячейку
                            cell1.Value = valStr
                        End If
                    Next j
                    countRows += 1
                Next
                xlApp.Visible = True
            Catch ex As System.Exception
            End Try
            xlBook = Nothing
            xlApp = Nothing
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'новый отчет по инвентаризации земель
    <CommandMethod("ListLandParcelsInventaryPPT2")>
    Public Sub ListLandParcelsInventaryPPT2()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        Dim PutcTableDoc As String = parentPatch2 & "\TemplateTables\Land_Cadastral\"

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тело программы
        Dim ArrayTablesXML As String(,) = Nothing 'массив с таблицами для фильтрации
        ReDim ArrayTablesXML(1, 0)
        ArrayTablesXML(0, 0) = "LAND_CADASTRAL_PARCEL"
        ArrayTablesXML(1, 0) = "Кадастровые участки"
        Dim arrayParcels As String(,) = Nothing
        Dim countArrayParcels As Integer = 0
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim arrayZoneInventary As String(,) = Nothing
            Dim countArrayZoneInventary As Integer = 0
            '=========================================================================================================
            'выбираем зоны
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        Dim typeLine As String = FuncPropertySet.FuncReadValueSetProperties(acEnt, "Type_line")
                        If typeLine.Trim.Length > 0 Then
                            If IsArray(arrayZoneInventary) = False Then
                                ReDim Preserve arrayZoneInventary(1, countArrayZoneInventary)
                                arrayZoneInventary(0, countArrayZoneInventary) = typeLine
                                arrayZoneInventary(1, countArrayZoneInventary) = acEnt.Handle.ToString
                                countArrayZoneInventary += 1
                            Else
                                Dim boolfind As Boolean = False
                                For i As Integer = 0 To arrayZoneInventary.GetUpperBound(1)
                                    Dim tempnamezone As String = arrayZoneInventary(0, i)
                                    If tempnamezone.Trim Like typeLine Then
                                        arrayZoneInventary(1, i) = arrayZoneInventary(1, i) & ";" & acEnt.Handle.ToString
                                        boolfind = True
                                        Exit For
                                    End If
                                Next
                                If boolfind = False Then
                                    ReDim Preserve arrayZoneInventary(1, countArrayZoneInventary)
                                    arrayZoneInventary(0, countArrayZoneInventary) = typeLine
                                    arrayZoneInventary(1, countArrayZoneInventary) = acEnt.Handle.ToString
                                    countArrayZoneInventary += 1
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            If IsArray(arrayZoneInventary) = False Then
                Exit Sub
            End If
            If arrayZoneInventary.GetUpperBound(1) > 1 Then
                Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayZoneInventary, 0)
            End If
            '=========================================================================================================
            'выбираем участки
            Dim acSSPrompt2 As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt2.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt2.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Polyline Then
                                Dim acPline As Polyline = acEnt
                                Dim arraydata As String(,) = Nothing
                                Dim boolData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acEnt, arraydata, ArrayTablesXML)
                                If IsArray(arraydata) = True Then
                                    ReDim Preserve arrayParcels(8, countArrayParcels)
                                    arrayParcels(8, countArrayParcels) = acPline.Handle.ToString
                                    For i As Integer = 0 To arraydata.GetUpperBound(1)
                                        Dim NField As String = arraydata(1, i)
                                        Dim Zn As String = arraydata(2, i)
                                        If Zn.Trim.Length > 0 Then
                                            If NField Like "Cad_number" Then
                                                arrayParcels(0, countArrayParcels) = Zn
                                            ElseIf NField Like "Adress" Then
                                                arrayParcels(1, countArrayParcels) = Zn
                                            ElseIf NField Like "UtilizationByDoc" Then
                                                arrayParcels(2, countArrayParcels) = Zn
                                            ElseIf NField Like "Category" Then
                                                arrayParcels(3, countArrayParcels) = Zn
                                            ElseIf NField Like "Area" Then
                                                arrayParcels(4, countArrayParcels) = Zn
                                            ElseIf NField Like "Patronymic" Then
                                                arrayParcels(5, countArrayParcels) = Zn
                                            ElseIf NField Like "Encumbrance" Then
                                                arrayParcels(6, countArrayParcels) = Zn
                                            ElseIf NField Like "CadastralCost" Then
                                                arrayParcels(7, countArrayParcels) = Zn
                                            End If
                                        End If
                                    Next i
                                    countArrayParcels += 1
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Word документ
            Dim WApp As Word.Application = Nothing
            Dim WDoc As Word.Document = Nothing
            'временный каталог для автосохранения таблицы
            Dim TempDir As String = ""
            Try
                TempDir = System.IO.Path.GetTempPath()
            Catch ex As System.Security.SecurityException
                TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Временный каталог")
            End Try
            Dim boolWDVisibl As Boolean = False
            Dim PutchDoc As String = PutcTableDoc & "Ведомость имущественно-правовой инвентаризации_.docx"
            Dim rng As Random = New Random
            Dim Numbrng As Integer = rng.Next
            Dim DestDoc As String = TempDir & "Ведомость имущественно-правовой инвентаризации" & Numbrng & ".docx"
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
            'WApp.Visible = True
            Dim tabl As Word.Table = Nothing
            For Each tbl As Word.Table In WDoc.Tables
                tabl = tbl
                If tabl.Rows.Count > 9 Then
                    Exit For
                End If
            Next
            If IsNothing(tabl) = True Then
                MsgBox("Не удалось найти таблицу с количеством стобцов равным 10. Макро прервано!!!")
                WDoc.Close()
                WApp.Quit()
                WApp = Nothing
            End If

            Dim oldNameZone As String = ""
            Dim countRow As Integer = 4
            For i As Integer = 0 To arrayZoneInventary.GetUpperBound(1)
                Dim arrayParcelInventary As String(,) = Nothing
                Dim countArrayParcelInventary As Integer = 0
                Dim nameZone As String = arrayZoneInventary(0, i)
                Dim arrayHandleStr As String() = arrayZoneInventary(1, i).Split(";")
                If IsArray(arrayHandleStr) = True Then
                    For j As Integer = 0 To arrayHandleStr.Length - 1
                        Dim handleStr As String = arrayHandleStr(j)
                        Dim idZone As ObjectId = AcadFunction.FuncReturnObjectID(handleStr)
                        If idZone.IsNull = False Then
                            Dim acLineZone As Polyline = idZone.GetObject(OpenMode.ForRead)
                            If acLineZone.Area > 0 Then
                                Dim tempCollection As DBObjectCollection = New DBObjectCollection()
                                tempCollection.Add(acLineZone)
                                Dim parentRegionCollection As DBObjectCollection = New DBObjectCollection()
                                parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                Dim parentRegion As Region = New Region()
                                If parentRegionCollection.Count > 0 Then
                                    parentRegion = parentRegionCollection.Item(0)
                                Else
                                    Continue For
                                End If
                                '=============================================================================================
                                If IsArray(arrayParcels) = True Then
                                    For k As Integer = 0 To arrayParcels.GetUpperBound(1)
                                        Dim strParcelHandle As String = arrayParcels(8, k)
                                        Dim id As ObjectId = AcadFunction.FuncReturnObjectID(strParcelHandle)
                                        If id.IsNull = False Then
                                            Dim pLineParcel As Polyline = id.GetObject(OpenMode.ForRead)
                                            Dim childRegion As Region = New Region()
                                            If pLineParcel.Area > 0 Then
                                                Dim tempCollection2 As DBObjectCollection = New DBObjectCollection()
                                                tempCollection2.Add(pLineParcel)
                                                Dim shildRegionCollection As DBObjectCollection = New DBObjectCollection()
                                                shildRegionCollection = Region.CreateFromCurves(tempCollection2)
                                                If shildRegionCollection.Count > 0 Then
                                                    childRegion = shildRegionCollection.Item(0)
                                                Else
                                                    Continue For
                                                End If
                                            End If

                                            If parentRegion.Area = 0 Then
                                                tempCollection = New DBObjectCollection()
                                                tempCollection.Add(acLineZone)
                                                parentRegionCollection = New DBObjectCollection()
                                                parentRegionCollection = Region.CreateFromCurves(tempCollection)
                                                If parentRegionCollection.Count > 0 Then
                                                    parentRegion = parentRegionCollection.Item(0)
                                                End If
                                            End If

                                            childRegion.BooleanOperation(BooleanOperationType.BoolIntersect, parentRegion)
                                            Dim newArea As Double = childRegion.Area
                                            If newArea > 0 Then
                                                Dim cadNumber As String = arrayParcels(0, k)
                                                If IsNothing(cadNumber) = True Then
                                                    cadNumber = "-"
                                                ElseIf cadNumber.Trim.Length = 0 Then
                                                    cadNumber = "-"
                                                End If
                                                'вычисляем кадастровую стоимость
                                                Dim sqrParcel As Double = Val(arrayParcels(4, k))
                                                Dim coat As String = arrayParcels(7, k)
                                                Dim valCoat As Double = 0
                                                If IsNothing(coat) = True Then
                                                    valCoat = 0
                                                ElseIf coat.Trim.Length = 0 Then
                                                    valCoat = 0
                                                Else
                                                    valCoat = Val(coat)
                                                End If
                                                Dim tempCoat As Double = 0
                                                If sqrParcel > 0 Then
                                                    tempCoat = (valCoat / sqrParcel) * newArea
                                                End If
                                                'ищем участок в массиве
                                                Dim boolFindParcel As Boolean = False
                                                If Not (cadNumber Like "-") Then
                                                    If IsArray(arrayParcelInventary) = True Then
                                                        For k1 As Integer = 0 To arrayParcelInventary.GetUpperBound(1)
                                                            Dim tempCadNumber As String = arrayParcelInventary(0, k1)
                                                            If tempCadNumber.Trim Like cadNumber Then
                                                                arrayParcelInventary(2, k1) = Val(arrayParcelInventary(2, k1)) + newArea
                                                                arrayParcelInventary(7, k1) = Val(arrayParcelInventary(7, k1)) + tempCoat
                                                                boolFindParcel = True
                                                            End If
                                                        Next k1
                                                    End If
                                                End If
                                                If boolFindParcel = False Then
                                                    ReDim Preserve arrayParcelInventary(8, countArrayParcelInventary)
                                                    arrayParcelInventary(0, countArrayParcelInventary) = cadNumber
                                                    arrayParcelInventary(1, countArrayParcelInventary) = arrayParcels(4, k)
                                                    arrayParcelInventary(2, countArrayParcelInventary) = Math.Round(newArea)
                                                    'адрес
                                                    Dim adr As String = arrayParcels(1, k)
                                                    If IsNothing(adr) = True Then
                                                        adr = "-"
                                                    ElseIf adr.Trim.Length = 0 Then
                                                        adr = "-"
                                                    End If
                                                    arrayParcelInventary(3, countArrayParcelInventary) = adr

                                                    'категория земель
                                                    Dim category As String = arrayParcels(3, k)
                                                    If IsNothing(category) = True Then
                                                        category = "-"
                                                    ElseIf category.Trim.Length = 0 Then
                                                        category = "-"
                                                    End If
                                                    arrayParcelInventary(4, countArrayParcelInventary) = category

                                                    'разрешенное использование
                                                    Dim utiliz As String = arrayParcels(2, k)
                                                    If IsNothing(utiliz) = True Then
                                                        utiliz = "-"
                                                    ElseIf utiliz.Trim.Length = 0 Then
                                                        utiliz = "-"
                                                    End If
                                                    arrayParcelInventary(5, countArrayParcelInventary) = utiliz

                                                    'правообладатель
                                                    Dim patr As String = arrayParcels(5, k)
                                                    If IsNothing(patr) = True Then
                                                        patr = "Нет данных"
                                                    ElseIf patr.Trim.Length = 0 Then
                                                        patr = "Нет данных"
                                                    End If
                                                    arrayParcelInventary(6, countArrayParcelInventary) = utiliz

                                                    'кадастровая стоимость
                                                    arrayParcelInventary(7, countArrayParcelInventary) = valCoat

                                                    countArrayParcelInventary += 1
                                                End If
                                            End If
                                        End If
                                    Next k
                                End If
                            End If
                        End If
                    Next j
                End If
                If IsArray(arrayParcelInventary) = True Then
                    For j As Integer = 0 To arrayParcelInventary.GetUpperBound(1)
                        If j = 0 Then
                            tabl.Rows.Add()
                            tabl.Cell(countRow, 1).Merge(MergeTo:=tabl.Cell(countRow, 10))
                            tabl.Cell(countRow, 1).Range.Text = nameZone
                            countRow += 1
                        End If
                        tabl.Rows.Add()
                        tabl.Cell(countRow, 1).Range.Text = arrayParcelInventary(0, j)
                        'площадь исходного участка
                        tabl.Cell(countRow, 2).Range.Text = arrayParcelInventary(1, j)
                        'площадь образуемого участка
                        tabl.Cell(countRow, 3).Range.Text = arrayParcelInventary(2, j)
                        'адрес
                        tabl.Cell(countRow, 4).Range.Text = arrayParcelInventary(3, j)
                        'категория земель
                        tabl.Cell(countRow, 5).Range.Text = arrayParcelInventary(4, j)
                        'разрешенное использование
                        tabl.Cell(countRow, 6).Range.Text = arrayParcelInventary(5, j)
                        'правообладатель
                        tabl.Cell(countRow, 7).Range.Text = arrayParcelInventary(6, j)
                        'кадастровая стоимость
                        tabl.Cell(countRow, 8).Range.Text = arrayParcelInventary(7, j)
                        countRow += 1
                    Next j
                End If
            Next
            WApp.Quit()
            WApp = Nothing
        End Using
    End Sub
End Class