Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD
Imports System.Globalization
Imports Autodesk.AutoCAD.EditorInput
Imports System.Xml
Imports NetTopologySuite.Geometries
'Участки постоянного отвода
Public Class PROJECT_PARCEL
    Public LABEL As String 'наименование объекта
    Public ConditionalNumber As String 'условный номер (обозначение в плане)
    Public Cad_number As String 'кадастровый номер
    Public Area As Double 'площадь
    Public ProjectCategory As String 'планируемая категория земель
    Public ProjectUtilization As String 'планируемое разрешенное использоапние
    Public Adress As String 'адрес
    Public Purpose As String 'цель предоставления
    Public Right As String 'вид права
    Public Patronymic As String 'собственник
    Public Encumbrance As String 'обременение
    Public Method_create As String 'Способ образования земельного участка (образование, раздел, раздел с исходным зу, перераспределение
    Public Confiscation As String 'изъятие/резервирование
    Public Rationale_size As String 'Обоснование размера земельного участка (ссылка на закон)
    Public GLOBALID As String 'Глобальный идентификатор
    Public LOCALID As String 'Локальный идентификатор
    Public NOTE As String 'Примечание
End Class
Public Class TEMPORARY_PARCEL
    Public LABEL As String 'наименование объекта
    Public ConditionalNumber As String 'условный номер (обозначение в плане)
    Public Cad_number As String 'кадастровый номер
    Public Area As Double 'площадь
    Public Right As String 'вид права (ct
    Public Patronymic As String 'собственник
    Public Encumbrance As String 'обременение
    Public Purpose As String 'цель предоставления
    Public Rationale_size As String 'Обоснование размера земельного участка (ссылка на закон)
    Public ThLayer As Double 'Толщина растительного слоя пригодная для рекультивации
    Public BadThLayer As Double 'Толщина растительного слоя не пригодная для рекультивации
    Public GLOBALID As String 'Глобальный идентификатор
    Public LOCALID As String 'Локальный идентификатор
    Public NOTE As String 'Примечание
    Public Handle As String
End Class
Public Class CADASTRAL_OSTROV
    Public LABEL As String 'наименование объекта
    Public ConditionalNumber As String 'условный номер
    Public IDKOUNTER As String 'идентификатор контура
    Public LOCALID As String 'Локальный идентификатор
    Public NOTE As String 'Примечание
    Public Handle As String
End Class

Public Class FuncCadastr
    '==============================================================================================
    'функция возвращает категорию земель по его коду
    Public Shared Function FuncReturnDictionaryParcel(ByVal userDBObjectColl As DBObjectCollection, ByVal userDBObjectCollOstrov As DBObjectCollection, ByVal nameTablePS As String) As Dictionary(Of Integer, String(,))
        Dim tempDictionary As Dictionary(Of Integer, String(,)) = New Dictionary(Of Integer, String(,))
        If userDBObjectColl.Count > 0 Then
            Dim arrayParcels As String(,) = Nothing
            Dim countArrayParcels As Integer = 0
            For Each acEnt As Autodesk.AutoCAD.DatabaseServices.Entity In userDBObjectColl
                Dim json_str As String = ""
                Dim strOstrov As String = ""
                Dim arrayCounter As String() = Nothing
                ReDim arrayCounter(4)
                Dim arrayDataPS As String(,) = Nothing
                Dim boolFindDataPs As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(acEnt, arrayDataPS, nameTablePS)
                If IsArray(arrayDataPS) = True Then
                    Dim conditionalNumber As String = MathFunction.FuncFindValueToArray2d(arrayDataPS, "ConditionalNumber", 1, 2)
                    If conditionalNumber.Trim.Length > 0 Then
                        'отделяем номер контура от участка
                        Dim fullNumberParcel As String = FuncFindFullConditionalNumberCadastralObject(conditionalNumber)
                        'получаем короткий номер участка
                        Dim numberParcel As Integer = FuncFindShortNumberCadastralObject(fullNumberParcel)
                        'получаем номер контура
                        Dim numberCounter As Integer = FuncFindNumberCounterCadastralObject(conditionalNumber)
                        If nameTablePS Like "PROJECT_PARCEL" Then
                            Dim parcels As PROJECT_PARCEL = New PROJECT_PARCEL
                            For i As Integer = 0 To arrayDataPS.GetUpperBound(1)
                                Dim ind As Integer = 4 + i
                                If arrayDataPS(1, i) Like "LABEL" Then
                                    parcels.LABEL = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "ConditionalNumber" Then
                                    parcels.ConditionalNumber = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Cad_number" Then
                                    parcels.Cad_number = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Area" Then
                                    If TypeOf acEnt Is Polyline Then
                                        Dim poly As Polyline = acEnt
                                        parcels.Area = MathFunction.FuncSQRTPolyline2(poly)
                                    End If
                                ElseIf arrayDataPS(1, i) Like "Adress" Then
                                    parcels.Adress = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Category" Then
                                    parcels.ProjectCategory = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Utilization" Then
                                    parcels.ProjectUtilization = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Purpose" Then
                                    parcels.Purpose = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Right" Then
                                    parcels.Right = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Patronymic" Then
                                    parcels.Patronymic = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Encumbrance" Then
                                    parcels.Encumbrance = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Method_create" Then
                                    parcels.Method_create = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Confiscation" Then
                                    parcels.Confiscation = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Rationale_size" Then
                                    parcels.Rationale_size = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "GLOBALID" Then
                                    parcels.GLOBALID = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "LOCALID" Then
                                    parcels.LOCALID = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "NOTE" Then
                                    parcels.NOTE = arrayDataPS(2, i)
                                End If
                            Next i
                            json_str = Newtonsoft.Json.JsonConvert.SerializeObject(parcels)

                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'находим островки
                            Dim arrayOstrov As String(,) = Nothing
                            Dim boolFindOstrov As Boolean = FuncFindOstrovByDbCollection(userDBObjectCollOstrov, parcels.LOCALID, arrayOstrov)
                            If IsArray(arrayOstrov) = True Then
                                For i As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                    strOstrov = strOstrov & " " & arrayOstrov(1, i)
                                Next i
                            End If
                        ElseIf nameTablePS Like "TEMPORARY_PARCEL" Then
                            Dim parcels As TEMPORARY_PARCEL = New TEMPORARY_PARCEL
                            For i As Integer = 0 To arrayDataPS.GetUpperBound(1)
                                Dim ind As Integer = 4 + i
                                If arrayDataPS(1, i) Like "LABEL" Then
                                    parcels.LABEL = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "ConditionalNumber" Then
                                    parcels.ConditionalNumber = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Cad_number" Then
                                    parcels.Cad_number = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Area" Then
                                    If TypeOf acEnt Is Polyline Then
                                        Dim poly As Polyline = acEnt
                                        parcels.Area = MathFunction.FuncSQRTPolyline2(poly)
                                    End If
                                ElseIf arrayDataPS(1, i) Like "Purpose" Then
                                    parcels.Purpose = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Right" Then
                                    parcels.Right = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Patronymic" Then
                                    parcels.Patronymic = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "Encumbrance" Then
                                    parcels.Encumbrance = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "ThLayer" Then
                                    parcels.ThLayer = Val(arrayDataPS(2, i))
                                ElseIf arrayDataPS(1, i) Like "BadThLayer" Then
                                    parcels.BadThLayer = Val(arrayDataPS(2, i))
                                ElseIf arrayDataPS(1, i) Like "GLOBALID" Then
                                    parcels.GLOBALID = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "LOCALID" Then
                                    parcels.LOCALID = arrayDataPS(2, i)
                                ElseIf arrayDataPS(1, i) Like "NOTE" Then
                                    parcels.NOTE = arrayDataPS(2, i)
                                End If
                            Next i
                            json_str = Newtonsoft.Json.JsonConvert.SerializeObject(parcels)
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'находим островки
                            Dim arrayOstrov As String(,) = Nothing
                            Dim boolFindOstrov As Boolean = FuncFindOstrovByDbCollection(userDBObjectCollOstrov, parcels.LOCALID, arrayOstrov)
                            If IsArray(arrayOstrov) = True Then
                                For i As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                    strOstrov = strOstrov & " " & arrayOstrov(1, i)
                                Next i
                            End If
                        End If
                        If tempDictionary.Count = 0 Then
                            ReDim Preserve arrayParcels(3, 0)
                            arrayParcels(0, 0) = numberCounter
                            arrayParcels(1, 0) = acEnt.Handle.ToString 'handle контура
                            arrayParcels(2, 0) = json_str
                            arrayParcels(3, 0) = strOstrov
                            tempDictionary.Add(numberParcel, arrayParcels)
                        Else
                            If tempDictionary.ContainsKey(numberParcel) = False Then
                                ReDim Preserve arrayParcels(3, 0)
                                arrayParcels(0, 0) = numberCounter
                                arrayParcels(1, 0) = acEnt.Handle.ToString 'handle контура
                                arrayParcels(2, 0) = json_str
                                arrayParcels(3, 0) = strOstrov
                                tempDictionary.Add(numberParcel, arrayParcels)
                            Else
                                arrayParcels = tempDictionary.Item(numberParcel)
                                Dim ind As Integer = arrayParcels.GetUpperBound(1) + 1
                                ReDim Preserve arrayParcels(3, ind)
                                arrayParcels(0, ind) = numberCounter
                                arrayParcels(1, ind) = acEnt.Handle.ToString 'handle контура
                                arrayParcels(2, ind) = json_str
                                arrayParcels(3, ind) = strOstrov
                                tempDictionary.Item(numberParcel) = arrayParcels
                            End If
                        End If
                    End If
                End If
            Next
        End If
        Return tempDictionary
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '==============================================================================================
    'функция возвращает категорию земель по его коду
    Public Shared Function FuncFindLandCategoryByCode(ByVal kodCategory As String) As String
        FuncFindLandCategoryByCode = kodCategory
        Try
            If kodCategory Like "003001000000" Then
                FuncFindLandCategoryByCode = "Земли сельскохозяйственного назначения"
            ElseIf kodCategory Like "003002000000" Then
                FuncFindLandCategoryByCode = "Земли населенных пунктов"
            ElseIf kodCategory Like "003003000000" Then
                FuncFindLandCategoryByCode = "Земли промышленности, транспорта и т.п."
            ElseIf kodCategory Like "003004000000" Then
                FuncFindLandCategoryByCode = "Земли особо охраняемых территорий и объектов"
            ElseIf kodCategory Like "003005000000" Then
                FuncFindLandCategoryByCode = "Земли лесного фонда"
            ElseIf kodCategory Like "003006000000" Then
                FuncFindLandCategoryByCode = "Земли водного фонда"
            ElseIf kodCategory Like "003007000000" Then
                FuncFindLandCategoryByCode = "Земли запаса"
            ElseIf kodCategory Like "003008000000" Then
                FuncFindLandCategoryByCode = "Категория не установлена"
            End If
        Catch
        End Try
    End Function
    '============================================================================================
    'функция возвращает разрешенное использование по коду
    Public Shared Function FuncFindLandUtilizationByCod(ByVal kodUtilization As String) As String
        FuncFindLandUtilizationByCod = kodUtilization
        Try
            If kodUtilization Like "141000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов сельскохозяйственного назначения и сельскохозяйственных угодий"
            ElseIf kodUtilization Like "141001000000" Then
                FuncFindLandUtilizationByCod = "Для сельскохозяйственного производства"
            ElseIf kodUtilization Like "141001010000" Then
                FuncFindLandUtilizationByCod = "Для использования в качестве сельскохозяйственных угодий"
            ElseIf kodUtilization Like "141001020000" Then
                FuncFindLandUtilizationByCod = "Для размещения зданий, строений, сооружений, используемых для производства, хранения и первичной переработки сельскохозяйственной продукции"
            ElseIf kodUtilization Like "141001030000" Then
                FuncFindLandUtilizationByCod = "Для размещения внутрихозяйственных дорог и коммуникаций"
            ElseIf kodUtilization Like "141001040000" Then
                FuncFindLandUtilizationByCod = "Для размещения водных объектов"
            ElseIf kodUtilization Like "141002000000" Then
                FuncFindLandUtilizationByCod = "Для ведения крестьянского (фермерского) хозяйства"
            ElseIf kodUtilization Like "141003000000" Then
                FuncFindLandUtilizationByCod = "Для ведения личного подсобного хозяйства"
            ElseIf kodUtilization Like "141004000000" Then
                FuncFindLandUtilizationByCod = "Для ведения гражданами садоводства и огородничества"
            ElseIf kodUtilization Like "141005000000" Then
                FuncFindLandUtilizationByCod = "Для ведения гражданами животноводства"
            ElseIf kodUtilization Like "141006000000" Then
                FuncFindLandUtilizationByCod = "Для дачного строительства"
            ElseIf kodUtilization Like "141007000000" Then
                FuncFindLandUtilizationByCod = "Для размещения древесно-кустарниковой растительности, предназначенной для защиты земель от воздействия негативных (вредных) природных, антропогенных и техногенных явлений"
            ElseIf kodUtilization Like "141008000000" Then
                FuncFindLandUtilizationByCod = "Для научно-исследовательских целей"
            ElseIf kodUtilization Like "141009000000" Then
                FuncFindLandUtilizationByCod = "Для учебных целей"
            ElseIf kodUtilization Like "141010000000" Then
                FuncFindLandUtilizationByCod = "Для сенокошения и выпаса скота гражданами"
            ElseIf kodUtilization Like "141011000000" Then
                FuncFindLandUtilizationByCod = "Фонд перераспределения"
            ElseIf kodUtilization Like "141012000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов охотничьего хозяйства"
            ElseIf kodUtilization Like "141013000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов рыбного хозяйства"
            ElseIf kodUtilization Like "141014000000" Then
                FuncFindLandUtilizationByCod = "Для иных видов сельскохозяйственного использования"
            ElseIf kodUtilization Like "142000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов, характерных для населённых пунктов"
            ElseIf kodUtilization Like "142001000000" Then
                FuncFindLandUtilizationByCod = "Для объектов жилой застройки"
            ElseIf kodUtilization Like "142001010000" Then
                FuncFindLandUtilizationByCod = "Для индивидуальной жилой застройки"
            ElseIf kodUtilization Like "142001020000" Then
                FuncFindLandUtilizationByCod = "Для многоквартирной застройки"
            ElseIf kodUtilization Like "142001020100" Then
                FuncFindLandUtilizationByCod = "Для малоэтажной застройки"
            ElseIf kodUtilization Like "142001020200" Then
                FuncFindLandUtilizationByCod = "Для среднеэтажной застройки"
            ElseIf kodUtilization Like "142001020300" Then
                FuncFindLandUtilizationByCod = "Для многоэтажной застройки"
            ElseIf kodUtilization Like "142001020400" Then
                FuncFindLandUtilizationByCod = "Для иных видов жилой застройки"
            ElseIf kodUtilization Like "142001030000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов дошкольного, начального, общего и среднего (полного) общего образования"
            ElseIf kodUtilization Like "142001040000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов, допустимых в жилых зонах и не перечисленных в классификаторе"
            ElseIf kodUtilization Like "142002000000" Then
                FuncFindLandUtilizationByCod = "Для объектов общественно-делового значения"
            ElseIf kodUtilization Like "142002010000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов социального и коммунально-бытового назначения"
            ElseIf kodUtilization Like "142002020000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов здравоохранения"
            ElseIf kodUtilization Like "142002030000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов культуры"
            ElseIf kodUtilization Like "142002040000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов торговли"
            ElseIf kodUtilization Like "142002040100" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов розничной торговли"
            ElseIf kodUtilization Like "142002040200" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов оптовой торговли"
            ElseIf kodUtilization Like "142002050000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов общественного питания"
            ElseIf kodUtilization Like "142002060000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов предпринимательской деятельности"
            ElseIf kodUtilization Like "142002070000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов среднего профессионального и высшего профессионального образования"
            ElseIf kodUtilization Like "142002080000" Then
                FuncFindLandUtilizationByCod = "Для размещения административных зданий"
            ElseIf kodUtilization Like "142002090000" Then
                FuncFindLandUtilizationByCod = "Для размещения научно-исследовательских учреждений"
            ElseIf kodUtilization Like "142002100000" Then
                FuncFindLandUtilizationByCod = "Для размещения культовых зданий"
            ElseIf kodUtilization Like "142002110000" Then
                FuncFindLandUtilizationByCod = "Для стоянок автомобильного транспорта"
            ElseIf kodUtilization Like "142002120000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов делового назначения, в том числе офисных центров"
            ElseIf kodUtilization Like "142002130000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов финансового назначения"
            ElseIf kodUtilization Like "142002140000" Then
                FuncFindLandUtilizationByCod = "Для размещения гостиниц"
            ElseIf kodUtilization Like "142002150000" Then
                FuncFindLandUtilizationByCod = "Для размещения подземных или многоэтажных гаражей"
            ElseIf kodUtilization Like "142002160000" Then
                FuncFindLandUtilizationByCod = "Для размещения индивидуальных гаражей"
            ElseIf kodUtilization Like "142002170000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов общественно-делового значения, обеспечивающих жизнь граждан"
            ElseIf kodUtilization Like "142003000000" Then
                FuncFindLandUtilizationByCod = "Для общего пользования (уличная сеть)"
            ElseIf kodUtilization Like "142004000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов специального назначения"
            ElseIf kodUtilization Like "142004010000" Then
                FuncFindLandUtilizationByCod = "Для размещения кладбищ"
            ElseIf kodUtilization Like "142004020000" Then
                FuncFindLandUtilizationByCod = "Для размещения крематориев"
            ElseIf kodUtilization Like "142004030000" Then
                FuncFindLandUtilizationByCod = "Для размещения скотомогильников"
            ElseIf kodUtilization Like "142004040000" Then
                FuncFindLandUtilizationByCod = "Под объектами размещения отходов потребления"
            ElseIf kodUtilization Like "142004050000" Then
                FuncFindLandUtilizationByCod = "Под иными объектами специального назначения"
            ElseIf kodUtilization Like "142005000000" Then
                FuncFindLandUtilizationByCod = "Для размещения коммунальных, складских объектов"
            ElseIf kodUtilization Like "142006000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов жилищно-коммунального хозяйства"
            ElseIf kodUtilization Like "142007000000" Then
                FuncFindLandUtilizationByCod = "Для иных видов использования, характерных для населённых пунктов"
            ElseIf kodUtilization Like "143000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики, обеспечения космической деятельности, обороны, безопасности и иного специального назначения"
            ElseIf kodUtilization Like "143001000000" Then
                FuncFindLandUtilizationByCod = "Для размещения промышленных объектов"
            ElseIf kodUtilization Like "143001010000" Then
                FuncFindLandUtilizationByCod = "Для размещения производственных и административных зданий, строений, сооружений и обслуживающих их объектов"
            ElseIf kodUtilization Like "143001010100" Then
                FuncFindLandUtilizationByCod = "Для размещения производственных зданий"
            ElseIf kodUtilization Like "143001010200" Then
                FuncFindLandUtilizationByCod = "Для размещения коммуникаций"
            ElseIf kodUtilization Like "143001010300" Then
                FuncFindLandUtilizationByCod = "Для размещения подъездных путей"
            ElseIf kodUtilization Like "143001010400" Then
                FuncFindLandUtilizationByCod = "Для размещения складских помещений"
            ElseIf kodUtilization Like "143001010500" Then
                FuncFindLandUtilizationByCod = "Для размещения административных зданий"
            ElseIf kodUtilization Like "143001010600" Then
                FuncFindLandUtilizationByCod = "Для размещения культурно-бытовых зданий"
            ElseIf kodUtilization Like "143001010700" Then
                FuncFindLandUtilizationByCod = "Для размещения иных сооружений промышленности"
            ElseIf kodUtilization Like "143001020000" Then
                FuncFindLandUtilizationByCod = "Для добычи и разработки полезных ископаемых"
            ElseIf kodUtilization Like "143001030000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов промышленности"
            ElseIf kodUtilization Like "143002000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов энергетики"
            ElseIf kodUtilization Like "143002010000" Then
                FuncFindLandUtilizationByCod = "Для размещения электростанций и обслуживающих сооружений и объектов"
            ElseIf kodUtilization Like "143002010100" Then
                FuncFindLandUtilizationByCod = "Для размещения гидроэлектростанций"
            ElseIf kodUtilization Like "143002010200" Then
                FuncFindLandUtilizationByCod = "Для размещения атомных станций"
            ElseIf kodUtilization Like "143002010300" Then
                FuncFindLandUtilizationByCod = "Для размещения ядерных установок"
            ElseIf kodUtilization Like "143002010400" Then
                FuncFindLandUtilizationByCod = "Для размещения пунктов хранения ядерных материалов и радиоактивных веществ энергетики"
            ElseIf kodUtilization Like "143002010500" Then
                FuncFindLandUtilizationByCod = "Для размещения хранилищ радиоактивных отходов"
            ElseIf kodUtilization Like "143002010600" Then
                FuncFindLandUtilizationByCod = "Для размещения тепловых станций"
            ElseIf kodUtilization Like "143002010700" Then
                FuncFindLandUtilizationByCod = "Для размещения иных типов электростанций"
            ElseIf kodUtilization Like "143002010800" Then
                FuncFindLandUtilizationByCod = "Для размещения иных обслуживающих сооружений и объектов"
            ElseIf kodUtilization Like "143002020000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов электросетевого хозяйства"
            ElseIf kodUtilization Like "143002020100" Then
                FuncFindLandUtilizationByCod = "Для размещения воздушных линий электропередачи"
            ElseIf kodUtilization Like "143002020200" Then
                FuncFindLandUtilizationByCod = "Для размещения наземных сооружений кабельных линий электропередачи"
            ElseIf kodUtilization Like "143002020300" Then
                FuncFindLandUtilizationByCod = "Для размещения подстанций"
            ElseIf kodUtilization Like "143002020400" Then
                FuncFindLandUtilizationByCod = "Для размещения распределительных пунктов"
            ElseIf kodUtilization Like "143002020500" Then
                FuncFindLandUtilizationByCod = "Для размещения других сооружений и объектов электросетевого хозяйства"
            ElseIf kodUtilization Like "143002030000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов энергетики"
            ElseIf kodUtilization Like "143003000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов транспорта"
            ElseIf kodUtilization Like "143003010000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации объектов железнодорожного транспорта"
            ElseIf kodUtilization Like "143003010100" Then
                FuncFindLandUtilizationByCod = "Для размещения железнодорожных путей и их конструктивных элементов"
            ElseIf kodUtilization Like "143003010200" Then
                FuncFindLandUtilizationByCod = "Для размещения полос отвода железнодорожных путей"
            ElseIf kodUtilization Like "143003010300" Then
                FuncFindLandUtilizationByCod = "Для размещения, эксплуатации, расширения и реконструкции строений, зданий, сооружений, в том числе железнодорожных вокзалов, железнодорожных станций, а также устройств и других объектов, необходимых для эксплуатации, содержания, строительства, реконструкции, ремонта, развития наземных и подземных зданий, строений, сооружений, устройств и других объектов железнодорожного транспорта"
            ElseIf kodUtilization Like "143003010301" Then
                FuncFindLandUtilizationByCod = "Для размещения железнодорожных вокзалов"
            ElseIf kodUtilization Like "143003010302" Then
                FuncFindLandUtilizationByCod = "Для размещения железнодорожных станций"
            ElseIf kodUtilization Like "143003010303" Then
                FuncFindLandUtilizationByCod = "Для размещения устройств и других объектов, необходимых для эксплуатации, содержания, строительства, реконструкции, ремонта, развития наземных и подземных зданий, строений, сооружений, устройств и других объектов железнодорожного транспорта"
            ElseIf kodUtilization Like "143003020000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации объектов автомобильного транспорта и объектов дорожного хозяйства"
            ElseIf kodUtilization Like "143003020100" Then
                FuncFindLandUtilizationByCod = "Для размещения автомобильных дорог и их конструктивных элементов"
            ElseIf kodUtilization Like "143003020200" Then
                FuncFindLandUtilizationByCod = "Для размещения полос отвода"
            ElseIf kodUtilization Like "143003020300" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов дорожного сервиса в полосах отвода автомобильных дорог"
            ElseIf kodUtilization Like "143003020400" Then
                FuncFindLandUtilizationByCod = "Для размещения дорожных сооружений"
            ElseIf kodUtilization Like "143003020500" Then
                FuncFindLandUtilizationByCod = "Для размещения автовокзалов и автостанций"
            ElseIf kodUtilization Like "143003020600" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов автомобильного транспорта и дорожного хозяйства"
            ElseIf kodUtilization Like "143003030000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации объектов морского, внутреннего водного транспорта"
            ElseIf kodUtilization Like "143003030100" Then
                FuncFindLandUtilizationByCod = "Для размещения искусственно созданных внутренних водных путей"
            ElseIf kodUtilization Like "143003030200" Then
                FuncFindLandUtilizationByCod = "Для размещения морских и речных портов, причалов, пристаней"
            ElseIf kodUtilization Like "143003030300" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов морского, внутреннего водного транспорта"
            ElseIf kodUtilization Like "143003030400" Then
                FuncFindLandUtilizationByCod = "Для выделения береговой полосы"
            ElseIf kodUtilization Like "143003040000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации объектов воздушного транспорта"
            ElseIf kodUtilization Like "143003040100" Then
                FuncFindLandUtilizationByCod = "Для размещения аэропортов и аэродромов"
            ElseIf kodUtilization Like "143003040200" Then
                FuncFindLandUtilizationByCod = "Для размещения аэровокзалов"
            ElseIf kodUtilization Like "143003040300" Then
                FuncFindLandUtilizationByCod = "Для размещения взлетно-посадочных полос"
            ElseIf kodUtilization Like "143003040400" Then
                FuncFindLandUtilizationByCod = "Для размещения иных наземных объектов воздушного транспорта"
            ElseIf kodUtilization Like "143003050000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации объектов трубопроводного транспорта"
            ElseIf kodUtilization Like "143003050100" Then
                FuncFindLandUtilizationByCod = "Для размещения нефтепроводов"
            ElseIf kodUtilization Like "143003050200" Then
                FuncFindLandUtilizationByCod = "Для размещения газопроводов"
            ElseIf kodUtilization Like "143003050300" Then
                FuncFindLandUtilizationByCod = "Для размещения иных трубопроводов"
            ElseIf kodUtilization Like "143003050400" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов трубопроводного транспорта"
            ElseIf kodUtilization Like "143003060000" Then
                FuncFindLandUtilizationByCod = "Для размещения и эксплуатации иных объектов транспорта"
            ElseIf kodUtilization Like "143004000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов связи, радиовещания, телевидения, информатики"
            ElseIf kodUtilization Like "143004010000" Then
                FuncFindLandUtilizationByCod = "Для размещения эксплуатационных предприятий связи и обслуживания линий связи"
            ElseIf kodUtilization Like "143004020000" Then
                FuncFindLandUtilizationByCod = "Для размещения кабельных, радиорелейных и воздушных линий связи и линий радиофикации на трассах кабельных и воздушных линий связи и радиофикации и их охранные зоны"
            ElseIf kodUtilization Like "143004030000" Then
                FuncFindLandUtilizationByCod = "Для размещения подземных кабельных и воздушных линий связи и радиофикации и их охранные зоны"
            ElseIf kodUtilization Like "143004040000" Then
                FuncFindLandUtilizationByCod = "Для размещения наземных и подземных необслуживаемых усилительных пунктов на кабельных линиях связи и их охранные зоны"
            ElseIf kodUtilization Like "143004050000" Then
                FuncFindLandUtilizationByCod = "Для размещения наземных сооружений и инфраструктур спутниковой связи"
            ElseIf kodUtilization Like "143004060000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов связи, радиовещания, телевидения, информатики"
            ElseIf kodUtilization Like "143005000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов, предназначенных для обеспечения космической деятельности"
            ElseIf kodUtilization Like "143005010000" Then
                FuncFindLandUtilizationByCod = "Для размещения космодромов, стартовых комплексов и пусковых установок"
            ElseIf kodUtilization Like "143005020000" Then
                FuncFindLandUtilizationByCod = "Для размещения командно-измерительных комплексов, центров и пунктов управления полетами космических объектов, приема, хранения и переработки информации"
            ElseIf kodUtilization Like "143005030000" Then
                FuncFindLandUtilizationByCod = "Для размещения баз хранения космической техники"
            ElseIf kodUtilization Like "143005040000" Then
                FuncFindLandUtilizationByCod = "Для размещения полигонов приземления космических объектов и взлетно-посадочных полос"
            ElseIf kodUtilization Like "143005050000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов экспериментальной базы для отработки космической техники"
            ElseIf kodUtilization Like "143005060000" Then
                FuncFindLandUtilizationByCod = "Для размещения центров и оборудования для подготовки космонавтов"
            ElseIf kodUtilization Like "143005070000" Then
                FuncFindLandUtilizationByCod = "Для размещения других наземных сооружений и техники, используемых при осуществлении космической деятельности"
            ElseIf kodUtilization Like "143006000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов, предназначенных для обеспечения обороны и безопасности"
            ElseIf kodUtilization Like "143006010000" Then
                FuncFindLandUtilizationByCod = "Для обеспечения задач обороны"
            ElseIf kodUtilization Like "143006010100" Then
                FuncFindLandUtilizationByCod = "Для размещения военных организаций, учреждений и других объектов"
            ElseIf kodUtilization Like "143006010200" Then
                FuncFindLandUtilizationByCod = "Для дислокации войск и сил флота"
            ElseIf kodUtilization Like "143006010300" Then
                FuncFindLandUtilizationByCod = "Для проведения учений и иных мероприятий"
            ElseIf kodUtilization Like "143006010400" Then
                FuncFindLandUtilizationByCod = "Для испытательных полигонов"
            ElseIf kodUtilization Like "143006010500" Then
                FuncFindLandUtilizationByCod = "Для мест уничтожения оружия и захоронения отходов"
            ElseIf kodUtilization Like "143006010600" Then
                FuncFindLandUtilizationByCod = "Для создания запасов материальных ценностей в государственном и мобилизационном резервах (хранилища, склады и другие)"
            ElseIf kodUtilization Like "143006010700" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов обороны"
            ElseIf kodUtilization Like "143006020000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов (территорий), обеспечивающих защиту и охрану Государственной границы Российской Федерации"
            ElseIf kodUtilization Like "143006020100" Then
                FuncFindLandUtilizationByCod = "Для обустройства и содержания инженерно-технических сооружений и заграждений"
            ElseIf kodUtilization Like "143006020200" Then
                FuncFindLandUtilizationByCod = "Для обустройства и содержания пограничных знаков"
            ElseIf kodUtilization Like "143006020300" Then
                FuncFindLandUtilizationByCod = "Для обустройства и содержания пограничных просек"
            ElseIf kodUtilization Like "143006020400" Then
                FuncFindLandUtilizationByCod = "Для обустройства и содержания коммуникаций"
            ElseIf kodUtilization Like "143006020500" Then
                FuncFindLandUtilizationByCod = "Для обустройства и содержания пунктов пропуска через Государственную границу Российской Федерации"
            ElseIf kodUtilization Like "143006020600" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов для защиты и охраны Государственной границы Российской Федерации"
            ElseIf kodUtilization Like "143006030000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов обороны и безопасности"
            ElseIf kodUtilization Like "143007000000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов промышленности, энергетики, транспорта, связи, радиовещания, телевидения, информатики, обеспечения космической деятельности, обороны, безопасности и иного специального назначения"
            ElseIf kodUtilization Like "144000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения особо охраняемых историко-культурных и природных объектов (территорий)"
            ElseIf kodUtilization Like "144001000000" Then
                FuncFindLandUtilizationByCod = "Для размещения особо охраняемых природных объектов (территорий)"
            ElseIf kodUtilization Like "144001010000" Then
                FuncFindLandUtilizationByCod = "Для размещения государственных природных заповедников (в том числе биосферных)"
            ElseIf kodUtilization Like "144001020000" Then
                FuncFindLandUtilizationByCod = "Для размещения государственных природных заказников"
            ElseIf kodUtilization Like "144001030000" Then
                FuncFindLandUtilizationByCod = "Для размещения национальных парков"
            ElseIf kodUtilization Like "144001040000" Then
                FuncFindLandUtilizationByCod = "Для размещения природных парков"
            ElseIf kodUtilization Like "144001050000" Then
                FuncFindLandUtilizationByCod = "Для размещения дендрологических парков"
            ElseIf kodUtilization Like "144001060000" Then
                FuncFindLandUtilizationByCod = "Для размещения ботанических садов"
            ElseIf kodUtilization Like "144001070000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов санаторного и курортного назначения"
            ElseIf kodUtilization Like "144001080000" Then
                FuncFindLandUtilizationByCod = "Территории месторождений минеральных вод, лечебных грязей, рапы лиманов и озер"
            ElseIf kodUtilization Like "144001090000" Then
                FuncFindLandUtilizationByCod = "Для традиционного природопользования"
            ElseIf kodUtilization Like "144001100000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных особо охраняемых природных территорий (объектов)"
            ElseIf kodUtilization Like "144002000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов (территорий) природоохранного назначения"
            ElseIf kodUtilization Like "144003000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов (территорий) рекреационного назначения"
            ElseIf kodUtilization Like "144003010000" Then
                FuncFindLandUtilizationByCod = "Для размещения домов отдыха, пансионатов, кемпингов"
            ElseIf kodUtilization Like "144003020000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов физической культуры и спорта"
            ElseIf kodUtilization Like "144003030000" Then
                FuncFindLandUtilizationByCod = "Для размещения туристических баз, стационарных и палаточных туристско-оздоровительных лагерей, домов рыболова и охотника, детских туристических станций"
            ElseIf kodUtilization Like "144003040000" Then
                FuncFindLandUtilizationByCod = "Для размещения туристических парков"
            ElseIf kodUtilization Like "144003050000" Then
                FuncFindLandUtilizationByCod = "Для размещения лесопарков"
            ElseIf kodUtilization Like "144003060000" Then
                FuncFindLandUtilizationByCod = "Для размещения учебно-туристических троп и трасс"
            ElseIf kodUtilization Like "144003070000" Then
                FuncFindLandUtilizationByCod = "Для размещения детских и спортивных лагерей"
            ElseIf kodUtilization Like "144003080000" Then
                FuncFindLandUtilizationByCod = "Для размещения скверов, парков, городских садов"
            ElseIf kodUtilization Like "144003090000" Then
                FuncFindLandUtilizationByCod = "Для размещения пляжей"
            ElseIf kodUtilization Like "144003100000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных объектов (территорий) рекреационного назначения"
            ElseIf kodUtilization Like "144004000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов историко-культурного назначения"
            ElseIf kodUtilization Like "144004010000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов культурного наследия народов Российской Федерации (памятников истории и культуры), в том числе объектов археологического наследия"
            ElseIf kodUtilization Like "144004020000" Then
                FuncFindLandUtilizationByCod = "Для размещения военных и гражданских захоронений"
            ElseIf kodUtilization Like "144005000000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных особо охраняемых историко-культурных и природных объектов (территорий)"
            ElseIf kodUtilization Like "145000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов лесного фонда"
            ElseIf kodUtilization Like "145001000000" Then
                FuncFindLandUtilizationByCod = "Для размещения лесной растительности"
            ElseIf kodUtilization Like "145002000000" Then
                FuncFindLandUtilizationByCod = "Для восстановления лесной растительности"
            ElseIf kodUtilization Like "145003000000" Then
                FuncFindLandUtilizationByCod = "Для прочих объектов лесного хозяйства"
            ElseIf kodUtilization Like "146000000000" Then
                FuncFindLandUtilizationByCod = "Для размещения объектов водного фонда"
            ElseIf kodUtilization Like "146001000000" Then
                FuncFindLandUtilizationByCod = "Под водными объектами"
            ElseIf kodUtilization Like "146002000000" Then
                FuncFindLandUtilizationByCod = "Для размещения гидротехнических сооружений"
            ElseIf kodUtilization Like "146003000000" Then
                FuncFindLandUtilizationByCod = "Для размещения иных сооружений, расположенных на водных объектах"
            ElseIf kodUtilization Like "147000000000" Then
                FuncFindLandUtilizationByCod = "Земли запаса (неиспользуемые)"
            End If
        Catch
        End Try
    End Function
    '=======================================================================================
    'функция возвращает вид права по его коду
    Public Shared Function FuncFindRightByCod(ByVal kodRight As String) As String
        FuncFindRightByCod = kodRight
        Try
            If kodRight Like "001001000000" Then
                FuncFindRightByCod = "Собственность"
            ElseIf kodRight Like "001002000000" Then
                FuncFindRightByCod = "Долевая собственность"
            ElseIf kodRight Like "001003000000" Then
                FuncFindRightByCod = "Совместная собственность"
            ElseIf kodRight Like "001004000000" Then
                FuncFindRightByCod = "Хозяйственное ведение"
            ElseIf kodRight Like "001005000000" Then
                FuncFindRightByCod = "Оперативное управление"
            ElseIf kodRight Like "001006000000" Then
                FuncFindRightByCod = "Пожизненное наследуемое владение"
            ElseIf kodRight Like "001007000000" Then
                FuncFindRightByCod = "Постоянное (бессрочное) пользование"
            ElseIf kodRight Like "001008000000" Then
                FuncFindRightByCod = "Сервитут (право)"
            ElseIf kodRight Like "001009000000" Then
                FuncFindRightByCod = "Иные права"
            Else
                FuncFindRightByCod = kodRight & "Право не установлено"
            End If
        Catch
        End Try
    End Function
    '=======================================================================================
    'функция возвращает назначение здания по его коду
    Public Shared Function FuncFindPurposeBuildByCod(ByVal kodBuild As String) As String
        FuncFindPurposeBuildByCod = kodBuild
        Try
            If kodBuild Like "204001000000" Then
                FuncFindPurposeBuildByCod = "Нежилое здание"
            ElseIf kodBuild Like "204002000000" Then
                FuncFindPurposeBuildByCod = "Жилой дом"
            ElseIf kodBuild Like "204003000000" Then
                FuncFindPurposeBuildByCod = "Многоквартирный дом"
            Else
                FuncFindPurposeBuildByCod = kodBuild
            End If
        Catch
        End Try
    End Function
    '=======================================================================================
    'функция возвращает вид объекта недвижимости
    Public Shared Function FuncFindTypeObjectByCod(ByVal kodObject As String) As String
        FuncFindTypeObjectByCod = kodObject
        Try
            If kodObject Like "002001000000" Then
                FuncFindTypeObjectByCod = "Объекты учёта и регистрации"
            ElseIf kodObject Like "002001001000" Then
                FuncFindTypeObjectByCod = "Земельный участок"
            ElseIf kodObject Like "002001002000" Then
                FuncFindTypeObjectByCod = "Здание"
            ElseIf kodObject Like "002001003000" Then
                FuncFindTypeObjectByCod = "Помещение"
            ElseIf kodObject Like "002001004000" Then
                FuncFindTypeObjectByCod = "Сооружение"
            ElseIf kodObject Like "002001004001" Then
                FuncFindTypeObjectByCod = "Линейное сооружение, расположенное более чем в одном кадастровом округе"
            ElseIf kodObject Like "002001004002" Then
                FuncFindTypeObjectByCod = "Условная часть линейного сооружения"
            ElseIf kodObject Like "002001005000" Then
                FuncFindTypeObjectByCod = "Объект незавершённого строительства"
            ElseIf kodObject Like "002001006000" Then
                FuncFindTypeObjectByCod = "Предприятие как имущественный комплекс (ПИК)"
            ElseIf kodObject Like "002001007000" Then
                FuncFindTypeObjectByCod = "Участок недр"
            ElseIf kodObject Like "002001008000" Then
                FuncFindTypeObjectByCod = "Единый недвижимый комплекс"
            ElseIf kodObject Like "002002000000" Then
                FuncFindTypeObjectByCod = "Границы"
            ElseIf kodObject Like "002002001000" Then
                FuncFindTypeObjectByCod = "Государственная граница Российской Федерации"
            ElseIf kodObject Like "002002002000" Then
                FuncFindTypeObjectByCod = "Граница между субъектами Российской Федерации"
            ElseIf kodObject Like "002002003000" Then
                FuncFindTypeObjectByCod = "Граница муниципального образования"
            ElseIf kodObject Like "002002004000" Then
                FuncFindTypeObjectByCod = "Граница населённого пункта"
            ElseIf kodObject Like "002003000000" Then
                FuncFindTypeObjectByCod = "Зоны"
            ElseIf kodObject Like "002003001000" Then
                FuncFindTypeObjectByCod = "Зона с особыми условиями использования территорий"
            ElseIf kodObject Like "002003002000" Then
                FuncFindTypeObjectByCod = "Территориальная зона"
            Else
                FuncFindTypeObjectByCod = kodObject
            End If
        Catch
        End Try
    End Function
    '======================================================================================
    'функция возвращает еденицу измерения по коду
    Public Shared Function FuncFindCadastrUnitByCode(ByVal Unit As Integer) As String
        FuncFindCadastrUnitByCode = Unit.ToString
        Select Case Unit
            Case 3
                FuncFindCadastrUnitByCode = "мм"
            Case 4
                FuncFindCadastrUnitByCode = "см"
            Case 5
                FuncFindCadastrUnitByCode = "дм"
            Case 6
                FuncFindCadastrUnitByCode = "м"
            Case 8
                FuncFindCadastrUnitByCode = "км"
            Case 9
                FuncFindCadastrUnitByCode = "мгм"
            Case 47
                FuncFindCadastrUnitByCode = "мор.миля"
            Case 50
                FuncFindCadastrUnitByCode = "кв.мм"
            Case 51
                FuncFindCadastrUnitByCode = "кв.см"
            Case 53
                FuncFindCadastrUnitByCode = "кв.дм"
            Case 55
                FuncFindCadastrUnitByCode = "кв.м"
            Case 58
                FuncFindCadastrUnitByCode = "тыс.кв.м"
            Case 59
                FuncFindCadastrUnitByCode = "га"
            Case 61
                FuncFindCadastrUnitByCode = "кв.км"
            Case 109
                FuncFindCadastrUnitByCode = "Ар"
            Case 359
                FuncFindCadastrUnitByCode = "сут."
            Case 360
                FuncFindCadastrUnitByCode = "нед."
            Case 361
                FuncFindCadastrUnitByCode = "дек."
            Case 362
                FuncFindCadastrUnitByCode = "мес."
            Case 364
                FuncFindCadastrUnitByCode = "кв."
            Case 365
                FuncFindCadastrUnitByCode = "полугод."
            Case 356
                FuncFindCadastrUnitByCode = "час"
            Case 355
                FuncFindCadastrUnitByCode = "мин."
            Case 366
                FuncFindCadastrUnitByCode = "год"
            Case 383
                FuncFindCadastrUnitByCode = "руб."
            Case 384
                FuncFindCadastrUnitByCode = "тыс.руб."
            Case 385
                FuncFindCadastrUnitByCode = "млн.руб."
            Case 386
                FuncFindCadastrUnitByCode = "млрд.руб."
            Case 1000
                FuncFindCadastrUnitByCode = ""
            Case 1001
                FuncFindCadastrUnitByCode = ""
            Case 1002
                FuncFindCadastrUnitByCode = "руб/кв.м"
            Case 1003
                FuncFindCadastrUnitByCode = "руб/Ар"
            Case 1004
                FuncFindCadastrUnitByCode = "руб/га"
            Case 1005
                FuncFindCadastrUnitByCode = ""
            Case 1006
                FuncFindCadastrUnitByCode = "руб/кв.км"
            Case 1007
                FuncFindCadastrUnitByCode = "раб.дн."
        End Select
    End Function
    '====================================================================================
    'функция возвращает статус участка
    Public Shared Function FuncCadastrStateParcel(ByVal StateIndex As Long) As String
        FuncCadastrStateParcel = StateIndex.ToString
        Select Case StateIndex
            Case 1
                FuncCadastrStateParcel = "Ранее учтенный"
            Case 5
                FuncCadastrStateParcel = "Временный"
            Case 6
                FuncCadastrStateParcel = "Учтенный"
            Case 7
                FuncCadastrStateParcel = "Снят с учета"
            Case 8
                FuncCadastrStateParcel = "Аннулированный"
        End Select
    End Function
    '=======================================================================================
    'функция возвращает характеристику сооружения
    Public Shared Function FuncTypeKeyParameter(ByVal TypeIndex As Long, ByVal Val As String) As String
        FuncTypeKeyParameter = Val
        Select Case TypeIndex
            Case 1
                FuncTypeKeyParameter = "Протяженность " & Val & " м."
            Case 2
                FuncTypeKeyParameter = "Глубина" & Val & " м."
            Case 3
                FuncTypeKeyParameter = "Объем" & Val & " куб.м."
            Case 4
                FuncTypeKeyParameter = "Высота" & Val & " м."
            Case 5
                FuncTypeKeyParameter = "Площадь" & Val & " кв.м."
            Case 6
                FuncTypeKeyParameter = "Площадь застройки" & Val & " кв.м."
            Case 7
                FuncTypeKeyParameter = "Глубина залегания" & Val & " м."
        End Select
    End Function
    '=======================================================================================
    'функция возвращает материал стен
    Public Shared Function FuncCadastrMaterialWall(ByVal MaterialWal As Long) As String
        FuncCadastrMaterialWall = MaterialWal
        Select Case MaterialWal
            Case 61001000000
                FuncCadastrMaterialWall = "Стены"
            Case 61001001000
                FuncCadastrMaterialWall = "Кирпичные"
            Case 61001001002
                FuncCadastrMaterialWall = "Кирпичные облегченные"
            Case 61001001003
                FuncCadastrMaterialWall = "Из природного камня"
            Case 61001002000
                FuncCadastrMaterialWall = "Деревянные"
            Case 61001002001
                FuncCadastrMaterialWall = "Рубленые"
            Case 61001002002
                FuncCadastrMaterialWall = "Каркасно-засыпные"
            Case 61001002003
                FuncCadastrMaterialWall = "Каркасно-обшивные"
            Case 61001002004
                FuncCadastrMaterialWall = "Сборно-щитовые"
            Case 61001002005
                FuncCadastrMaterialWall = "Дощатые"
            Case 61001002006
                FuncCadastrMaterialWall = "Деревянный каркас без обшивки"
            Case 61001003000
                FuncCadastrMaterialWall = "Смешанные"
            Case 61001003001
                FuncCadastrMaterialWall = "Каменные и деревянные"
            Case 61001003002
                FuncCadastrMaterialWall = "Каменные и бетонные"
            Case 61001004000
                FuncCadastrMaterialWall = "Легкие из местных материалов"
            Case 61001005000
                FuncCadastrMaterialWall = "Из прочих материалов"
            Case 61001006000
                FuncCadastrMaterialWall = "Бетонные"
            Case 61001006001
                FuncCadastrMaterialWall = "Монолитные"
            Case 61001006002
                FuncCadastrMaterialWall = "Из мелких бетонных блоков"
            Case 61001006003
                FuncCadastrMaterialWall = "Из легкобетонных панелей"
            Case 61001007000
                FuncCadastrMaterialWall = "Железобетонные"
            Case 61001007001
                FuncCadastrMaterialWall = "Крупнопанельные"
            Case 61001007002
                FuncCadastrMaterialWall = "Каркасно-панельные"
            Case 61001007003
                FuncCadastrMaterialWall = "Монолитные"
            Case 61001007004
                FuncCadastrMaterialWall = "Крупноблочные"
            Case 61001007005
                FuncCadastrMaterialWall = "Из унифицированных железобетонных элементов"
            Case 61001007006
                FuncCadastrMaterialWall = "Из железобетонных сегментов"
            Case 61001008000
                FuncCadastrMaterialWall = "Шлакобетонные"
            Case 61001009000
                FuncCadastrMaterialWall = " Металлические"
        End Select
    End Function
    '=======================================================================================
    'метод определения координат
    Public Shared Function FuncGeopointOpred(ByVal Val As String) As String
        FuncGeopointOpred = Val
        Select Case Val
            Case 692001000000
                FuncGeopointOpred = "Геодезический метод"
            Case 692002000000
                FuncGeopointOpred = "Фотограмметрический метод"
            Case 692003000000
                FuncGeopointOpred = "Картометрический метод"
            Case 692004000000
                FuncGeopointOpred = "Иное описание"
            Case 692005000000
                FuncGeopointOpred = "Метод спутниковых геодезических измерений (определений)"
            Case 692006000000
                FuncGeopointOpred = "Аналитический метод"
        End Select
    End Function
    '=======================================================================================
    'метод определения координат
    Public Shared Function FuncGeopointOpredKod(ByVal Val As String) As String
        FuncGeopointOpredKod = Val
        Select Case Val
            Case "Геодезический метод"
                FuncGeopointOpredKod = "692001000000"
            Case "Фотограмметрический метод"
                FuncGeopointOpredKod = "692002000000"
            Case "Картометрический метод"
                FuncGeopointOpredKod = "692003000000"
            Case "Иное описание"
                FuncGeopointOpredKod = "692004000000"
            Case "Метод спутниковых геодезических измерений (определений)"
                FuncGeopointOpredKod = "692005000000"
            Case "Аналитический метод"
                FuncGeopointOpredKod = "692006000000"
            Case Else
                FuncGeopointOpredKod = "692004000000"
        End Select
    End Function
    '=======================================================================================
    'погрешность определения координат точки
    Public Shared Function FuncReturnDeltaByCategory(ByVal strCategory As String) As Double
        FuncReturnDeltaByCategory = 0
        If IsNothing(strCategory) = True Then
            Return ""
        End If
        If strCategory.IndexOf("ельскохоз") > -1 Then
            FuncReturnDeltaByCategory = 2.5
        ElseIf strCategory.IndexOf("промышл") > -1 Then
            FuncReturnDeltaByCategory = 0.5
        ElseIf strCategory.IndexOf("водн") > -1 Then
            FuncReturnDeltaByCategory = 2.5
        ElseIf strCategory.IndexOf("запас") > -1 Then
            FuncReturnDeltaByCategory = 2.5
        ElseIf strCategory.IndexOf("особо") > -1 Then
            FuncReturnDeltaByCategory = 2.5
        ElseIf strCategory.IndexOf("лесн") > -1 Then
            FuncReturnDeltaByCategory = 5
        ElseIf strCategory.IndexOf("насел") > -1 Then
            FuncReturnDeltaByCategory = 0.1
        End If
    End Function
    '=======================================================================================
    'описание закрепления точки
    Public Shared Function FuncZakrepPointKod(ByVal Val As String) As String
        FuncZakrepPointKod = ""
        Select Case Val
            Case "Долговременный межевой знак"
                FuncZakrepPointKod = "626001000000"
            Case "Временный межевой знак"
                FuncZakrepPointKod = "626002000000"
            Case "Закрепление отсутствует"
                FuncZakrepPointKod = "626003000000"
        End Select
    End Function
    '=======================================================================================
    'функция возвращает тип зоны по коду
    Public Shared Function FuncCadastrTypeZone(ByVal kodZona As Long) As String
        Select Case kodZona
            Case 218010000000
                FuncCadastrTypeZone = "Территориальные зоны"
            Case 218010010000
                FuncCadastrTypeZone = "Жилая зона"
            Case 218010020000
                FuncCadastrTypeZone = "Общественно-деловая зона"
            Case 218010030000
                FuncCadastrTypeZone = "Производственная зона, зона инженерной и транспортной инфраструктур"
            Case 218010040000
                FuncCadastrTypeZone = "Зона сельскохозяйственного использования"
            Case 218010050000
                FuncCadastrTypeZone = "Зона рекреационного назначения"
            Case 218010060000
                FuncCadastrTypeZone = "Зона особо охраняемых территорий"
            Case 218010070000
                FuncCadastrTypeZone = "Зона специального назначения"
            Case 218010080000
                FuncCadastrTypeZone = "Зона размещения военных объектов"
            Case 218010090000
                FuncCadastrTypeZone = "Иная зона"
            Case 218020000000
                FuncCadastrTypeZone = "Зоны с особыми условиями использования территории"
            Case 218020010000
                FuncCadastrTypeZone = "Зоны охраны природных объектов"
            Case 218020010001
                FuncCadastrTypeZone = "Территория особо охраняемого природного объекта"
            Case 218020010002
                FuncCadastrTypeZone = "Охранная зона особо охраняемого природного объекта"
            Case 218020010003
                FuncCadastrTypeZone = "Водоохранная зона"
            Case 218020010004
                FuncCadastrTypeZone = "Прибрежная защитная полоса"
            Case 218020010005
                FuncCadastrTypeZone = "Зона санитарной охраны источников водоснабжения и водопроводов питьевого назначения"
            Case 218020010006
                FuncCadastrTypeZone = "Особо ценные земли"
            Case 218020020000
                FuncCadastrTypeZone = "Зоны охраны искусственных объектов"
            Case 218020020001
                FuncCadastrTypeZone = "Территория объекта культурного наследия"
            Case 218020020002
                FuncCadastrTypeZone = "Зона охраны объекта культурного наследия"
            Case 218020020003
                FuncCadastrTypeZone = "Охранная зона геодезического пункта"
            Case 218020020004
                FuncCadastrTypeZone = "Охранная зона стационарного пункта наблюдений за состоянием окружающей природной среды"
            Case 218020020005
                FuncCadastrTypeZone = "Охранная зона транспорта"
            Case 218020020006
                FuncCadastrTypeZone = "Охранная зона инженерных коммуникаций"
            Case 218020020007
                FuncCadastrTypeZone = "Охранная зона линий и сооружений связи и линий и сооружений радиофикации"
            Case 218020020008
                FuncCadastrTypeZone = "Придорожная полоса"
            Case 218020030000
                FuncCadastrTypeZone = "Зоны защиты населения"
            Case 218020030001
                FuncCadastrTypeZone = "Запретная зона при военном складе"
            Case 218020030002
                FuncCadastrTypeZone = "Запретный район при военном складе"
            Case 218020030003
                FuncCadastrTypeZone = "Загрязненные земли"
            Case 218020030004
                FuncCadastrTypeZone = "Охранная зона загрязненных земель"
            Case 218020030005
                FuncCadastrTypeZone = "Санитарно-защитная зона предприятий, сооружений и иных объектов"
            Case 218020030006
                FuncCadastrTypeZone = "Санитарный разрыв(санитарная полоса отчуждения)"
            Case 218020030007
                FuncCadastrTypeZone = "Зона ограничения от передающего радиотехнического объекта"
            Case 218020030008
                FuncCadastrTypeZone = "Район падения отделяющихся частей ракет"
            Case 218020040000
                FuncCadastrTypeZone = "Прочие зоны с особыми условиями использования территории"
            Case 218020040001
                FuncCadastrTypeZone = "Зона публичного сервитута"
            Case 218020040002
                FuncCadastrTypeZone = "Зона резервирования земель"
            Case 218020040003
                FuncCadastrTypeZone = "Горный отвод"
            Case 218020040004
                FuncCadastrTypeZone = "Пригородная зона"
            Case 218020040005
                FuncCadastrTypeZone = "Зеленая зона"
            Case 218020040006
                FuncCadastrTypeZone = "Особая экономическая зона"
            Case 218020040007
                FuncCadastrTypeZone = "Пограничная зона"
            Case 218020040008
                FuncCadastrTypeZone = "Территория традиционного природопользования"
            Case 218020050000
                FuncCadastrTypeZone = "Иные зоны с особыми условиями использования территории"
            Case 218020050001
                FuncCadastrTypeZone = "Иная зона с особыми условиями использования территории"
        End Select
    End Function
    '===============================================================================================
    'функция рисует полилини, по координатам из массва
    Public Shared Function FuncDrawPolyline(ByVal ArrayVertex As String(,)) As ObjectId
        FuncDrawPolyline = Nothing
        If IsArray(ArrayVertex) = False Then Exit Function
        If ArrayVertex.GetUpperBound(1) < 1 Then Exit Function
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acPoly As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline()
                For i As Integer = 0 To ArrayVertex.GetUpperBound(1)
                    Try
                        Dim x As Double = CDbl(ArrayVertex(3, i))
                        Dim y As Double = CDbl(ArrayVertex(2, i))
                        Dim Pt1 As Point2d = New Point2d(x, y)
                        acPoly.AddVertexAt(i, Pt1, 0, 0, 0)
                    Catch
                    End Try
                Next i
                If acPoly.StartPoint.GetVectorTo(acPoly.EndPoint).Length = 0 Then
                    acPoly.Closed = True
                End If
                acPoly.SetDatabaseDefaults()
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)
                acTrans.Commit()
                FuncDrawPolyline = acPoly.ObjectId
            End Using
        Catch ex As System.Exception
        End Try
    End Function
    '-===================================================================================
    'вставка в Чертеж блока с атрибутами точки
    Public Shared Function FuncInsertCadastrPoint(ByVal insertPosition As Point3d, ByRef AttrPrompt As Double(), ByVal NumbPoint As String,
ByVal DeltaPoint As String, ByVal ZakrPoint As String, ByVal HText1 As Double, ByVal layerGeoPoint As String,
ByVal PutshBlock As String, ByVal Mash1 As Scale3d, Optional ByVal RotS As Double = 0, Optional ByVal idTextStyle As ObjectId = Nothing, Optional ByVal metodCoord As String = "Геодезический метод") As BlockReference
        FuncInsertCadastrPoint = Nothing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        Dim br As BlockReference = Nothing
        Using (tr)
            Try
                Dim pointAttrXY As Point3d = New Point3d(AttrPrompt(0), AttrPrompt(1), 0)
                Dim Goriz As Integer = AttrPrompt(2)
                Dim Vert As Integer = AttrPrompt(3)
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForWrite, False), BlockTable)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'вставка блока
                Dim blockId As ObjectId = bt(PutshBlock)
                br = New BlockReference(insertPosition, blockId)
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
                br.Layer = "0"
                br.ScaleFactors = Mash1
                btr.AppendEntity(br)
                tr.AddNewlyCreatedDBObject(br, True)
                'номер точки
                Dim ad As AttributeDefinition = New AttributeDefinition
                ad.SetDatabaseDefaults()
                Dim ar As AttributeReference = New AttributeReference()
                ar.SetAttributeFromBlock(ad, br.BlockTransform)
                br.AttributeCollection.AppendAttribute(ar)
                tr.AddNewlyCreatedDBObject(ar, True)
                ar.Layer = "0"
                ar.Tag = "NUMBER"
                ar.TextString = NumbPoint
                ar.Position = pointAttrXY
                ar.Oblique = RotS
                If IsNothing(idTextStyle) = False Then
                    If idTextStyle.IsNull = False Then
                        ar.TextStyleId = idTextStyle
                    End If
                End If
                If Goriz = 1 Then
                    ar.HorizontalMode = TextHorizontalMode.TextLeft
                Else
                    ar.HorizontalMode = TextHorizontalMode.TextRight
                End If
                If Vert = 1 Then
                    ar.VerticalMode = TextVerticalMode.TextBottom
                Else
                    ar.VerticalMode = TextVerticalMode.TextTop
                End If
                ar.AlignmentPoint = pointAttrXY
                ar.Height = HText1
                'погрешность точки
                Dim ad2 As AttributeDefinition = New AttributeDefinition
                ad2.Layer = "0"
                ad2.Tag = "DELTA"
                ad2.TextString = DeltaPoint
                ad2.Position = pointAttrXY
                ad2.Invisible = True
                ad2.SetDatabaseDefaults()
                Dim ar2 As AttributeReference = New AttributeReference()
                ar2.SetAttributeFromBlock(ad2, br.BlockTransform)
                br.AttributeCollection.AppendAttribute(ar2)
                tr.AddNewlyCreatedDBObject(ar2, True)
                'описание закрепления точки
                Dim ad3 As AttributeDefinition = New AttributeDefinition
                ad3.Layer = "0"
                ad3.Tag = "FASTDESK"
                ad3.TextString = ZakrPoint
                ad3.Position = pointAttrXY
                ad3.Invisible = True
                ad3.SetDatabaseDefaults()
                Dim ar3 As AttributeReference = New AttributeReference()
                ar3.SetAttributeFromBlock(ad3, br.BlockTransform)
                br.AttributeCollection.AppendAttribute(ar3)
                tr.AddNewlyCreatedDBObject(ar3, True)
                'метод определения координат
                Dim ad4 As AttributeDefinition = New AttributeDefinition
                ad4.Layer = "0"
                ad4.Tag = "METHODCOORDINATES"
                ad4.TextString = metodCoord
                ad4.Position = pointAttrXY
                ad4.Invisible = True
                ad4.SetDatabaseDefaults()
                Dim ar4 As AttributeReference = New AttributeReference()
                ar4.SetAttributeFromBlock(ad4, br.BlockTransform)
                br.AttributeCollection.AppendAttribute(ar4)
                tr.AddNewlyCreatedDBObject(ar4, True)
                'слой для блока
                If layerGeoPoint.Trim.Length > 0 Then
                    br.Layer = layerGeoPoint
                Else
                    br.Layer = "0"
                End If
            Catch ex As System.Exception
                tr.Abort()
                FuncInsertCadastrPoint = Nothing
            End Try
            tr.Commit()
            Return br
        End Using
    End Function
    '-===================================================================================
    'Функция возвращает координаты вставки атрибута и параметры точки выравнивания(массив-X,Y,ГорВыр 1-лево 2-право,ВертВыр 1-низ 2-верх)
    Public Shared Function FuncReplaseInsertCoordAttr2(ByVal Vert1 As Point2d, ByVal Vert2 As Point2d, ByVal Vert3 As Point2d, ByRef insAttr As Double(), ByVal Dist As Double, Optional clockPline As Boolean = True) As Boolean
        FuncReplaseInsertCoordAttr2 = False
        Try
            'определяем дирекционный угол на первую линию
            Dim Angle1 As Double = Vert2.GetVectorTo(Vert1).Angle
            'определяем дирекционый угол на вторую линию
            Dim Angle2 As Double = Vert2.GetVectorTo(Vert3).Angle
            'определяем внутренний угол
            Dim Angle As Double = Angle2 - Angle1
            If Angle < 0 Then Angle = Angle + Math.PI * 2 'если угол меньше 0 тогда добавим 360
            Dim Angle2D As Double = Angle / 2 'разделим угол пополам
            Dim Napravlenie As Double = 0
            If clockPline = True Then
                Napravlenie = Angle1 + Angle2D + Math.PI  'вычисляем дир.напраление
            Else
                Napravlenie = Angle1 + Angle2D  'вычисляем дир.напраление
            End If
            If Napravlenie > Math.PI * 2 Then Napravlenie = Napravlenie - Math.PI * 2 'если угол больше 360 градусов
            'вычисляем координаты точки вставки атрибута
            Dim Dx As Double = Dist * Math.Cos(Napravlenie) 'вычисляем приращения координат
            Dim Dy As Double = Dist * Math.Sin(Napravlenie)
            ReDim insAttr(3)
            insAttr(0) = Vert2.X + Dx
            insAttr(1) = Vert2.Y + Dy
            'возврящаем точку выравнивания трибута
            Select Case Napravlenie
                Case 0 To Math.PI / 2
                    insAttr(2) = 1 'горизонтальное выравнивание(лево)
                    insAttr(3) = 1 'вертикальное выравнивание(низ)
                Case Math.PI / 2 To Math.PI
                    insAttr(2) = 2 'горизонтальное выравнивание(право)
                    insAttr(3) = 1 'вертикальное выравнивание(низ)
                Case Math.PI To Math.PI + Math.PI / 2
                    insAttr(2) = 2 'горизонтальное выравнивание(право)
                    insAttr(3) = 2 'вертикальное выравнивание(верх)
                Case Math.PI + Math.PI / 2 To 2 * Math.PI
                    insAttr(2) = 1 'горизонтальное выравнивание(лево)
                    insAttr(3) = 2 'вертикальное выравнивание(верх)
            End Select
            FuncReplaseInsertCoordAttr2 = True
        Catch ex As System.Exception
            insAttr(0) = Vert2.X
            insAttr(1) = Vert2.Y
            insAttr(2) = 1
            insAttr(3) = 1
        End Try
    End Function
    '-===================================================================================
    'Функция возвращает координаты вставки атрибута и параметры точки выравнивания(массив-X,Y,ГорВыр,ВертВыр)
    Public Shared Function FuncReplaseInsertCoordAttr(ByVal Vert1 As Point2d, ByVal Vert2 As Point2d, ByVal Vert3 As Point2d, ByRef insAttr As Double(), ByVal Dist As Double, Optional closePline As Boolean = False) As Boolean
        FuncReplaseInsertCoordAttr = False
        Try
            'приращения координат по первой линии
            Dim DX1 As Double = Vert1(0) - Vert2(0)
            Dim DY1 As Double = Vert1(1) - Vert2(1)
            'приращения координат по второй линии
            Dim DX2 As Double = Vert3(0) - Vert2(0)
            Dim DY2 As Double = Vert3(1) - Vert2(1)
            'определяем дирекционный угол на первую линию
            Dim Angle1 As Double = Math.Atan2(DY1, DX1)
            If DX1 > 0 And DY1 > 0 Then
                Angle1 = Angle1
            ElseIf DX1 < 0 And DY1 > 0 Then
                Angle1 = Angle1
            ElseIf DX1 < 0 And DY1 < 0 Then
                Angle1 = Math.PI * 2 - Math.Abs(Angle1)
            ElseIf DX1 > 0 And DY1 < 0 Then
                Angle1 = Math.PI * 2 - Math.Abs(Angle1)
            End If
            'определяем дирекционый угол на вторую линию
            Dim Angle2 As Double = Math.Atan2(DY2, DX2)
            If DX2 > 0 And DY2 > 0 Then
                Angle2 = Angle2
            ElseIf DX2 < 0 And DY2 > 0 Then
                Angle2 = Angle2
            ElseIf DX2 < 0 And DY2 < 0 Then
                Angle2 = Math.PI * 2 - Math.Abs(Angle2)
            ElseIf DX2 > 0 And DY2 < 0 Then
                Angle2 = Math.PI * 2 - Math.Abs(Angle2)
            End If
            'определяем внутренний угол
            Dim Angle As Double = Angle2 - Angle1
            If Angle < 0 Then Angle = Angle + Math.PI * 2 'если угол меньше 0 тогда добавим 360
            Dim Angle2D As Double = Angle / 2 'разделим угол пополам
            Dim Napravlenie As Double = Angle1 + Angle2D + Math.PI * 2 'если угол развернутый
            ' Dim Napravlenie As Double= 0 'направление на атрибут (начальное значение равно 0)
            ' If Angle > Math.PI Then
            'Napravlenie = Angle2 - Angle2D
            'If Napravlenie <0 Then Napravlenie= Napravlenie + Math.PI * 2 'если угол развернутый
            ' Else
            'Napravlenie = Angle2 - Angle2D + Math.PI 'иначе меняем направление на обратное
            'End If
            If Napravlenie > Math.PI * 2 Then Napravlenie = Napravlenie - Math.PI * 2 'если угол больше 360градусов
            Dim Dx As Double = Dist * Math.Cos(Napravlenie) 'вычисляем приращения координат
            Dim Dy As Double = Dist * Math.Sin(Napravlenie)
            ReDim insAttr(3)
            insAttr(0) = Vert2.X + Dx
            insAttr(1) = Vert2.Y + Dy
            Select Case Napravlenie
                Case 0 To Math.PI / 2
                    insAttr(2) = 1 'горизонтальное выравнивание(право)
                    insAttr(3) = 1 'вертикальное выравнивание(низ)
                Case Math.PI / 2 To Math.PI
                    insAttr(2) = 2 'горизонтальное выравнивание(право)
                    insAttr(3) = 1 'вертикальное выравнивание(низ)
                Case Math.PI To Math.PI + Math.PI / 2
                    insAttr(2) = 2 'горизонтальное выравнивание(право)
                    insAttr(3) = 2 'вертикальное выравнивание(низ)
                Case Math.PI + Math.PI / 2 To 2 * Math.PI
                    insAttr(2) = 1 'горизонтальное выравнивание(право)
                    insAttr(3) = 2 'вертикальное выравнивание(низ)
            End Select
            FuncReplaseInsertCoordAttr = True
        Catch
            insAttr(0) = Vert2.X
            insAttr(1) = Vert2.Y
            insAttr(2) = 1
            insAttr(3) = 1
        End Try
    End Function
    '-=====================================================================================
    '-Функция ищет центройд объекта 
    Public Shared Function FuncFindGeometricExtentsPline(ByVal Pline As Autodesk.AutoCAD.DatabaseServices.Polyline) As Point3d
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim pt3dmin As Point3d = Pline.GeometricExtents.MinPoint
                Dim pt3dmax As Point3d = Pline.GeometricExtents.MaxPoint
                Dim X As Single = (pt3dmax(0) - pt3dmin(0)) \ 2 + pt3dmin(0)
                Dim Y As Single = (pt3dmax(1) - pt3dmin(1)) \ 2 + pt3dmin(1)
                FuncFindGeometricExtentsPline = New Point3d(X, Y, 0)
                acTrans.Commit()
            Catch
                FuncFindGeometricExtentsPline = Pline.StartPoint
            End Try
        End Using
    End Function
    'Поиск объектов
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '========================================================================================================================
    'функция по локальному ID ищет объект
    Public Shared Function FuncFindObjectByLocalID(ByVal IDLocal As String) As Autodesk.AEC.DatabaseServices.Entity
        FuncFindObjectByLocalID = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim count As Integer = 0
        Dim count1 As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                'находим габаритный контейнер исходной полилиниии
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(ent)
                    If setIds.Count > 0 Then
                        Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(ent.Id, OpenMode.ForRead, False, False)
                        Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                        If psetDataColl.Count > 0 Then
                            Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                            For i As Integer = 0 To psetDataColl.Count - 1
                                Try
                                    pTempData = psetDataColl.Item(i)
                                    Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                    Dim data As String = pTempData.GetData() 'значение
                                    If dataField Like "LOCALID" Then
                                        If data.Trim.Length > 0 Then
                                            If data.Trim Like IDLocal Then
                                                FuncFindObjectByLocalID = ent
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                Catch ex As System.Exception
                                End Try
                            Next i
                        End If
                    End If
                Next
            Catch ex1 As System.Exception
                Exit Function
            End Try
        End Using
    End Function
    '========================================================================================================================
    'функция по idRNG участка ищет многострочный текст
    Public Shared Function FuncFindMTextByLocalId(ByVal UserIDObject As String) As MText
        FuncFindMTextByLocalId = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is MText Then 'если объект имеет тип атрибут
                        Dim acMtext As MText = ent
                        Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(ent)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                If psetDataColl.Count > 0 Then
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            If dataField Like "IDRNG" Then
                                                If data.Trim.Length > 0 Then
                                                    If data.Trim Like UserIDObject Then
                                                        FuncFindMTextByLocalId = acMtext
                                                        Exit Function
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Next
                        End If
                    End If
                Next
            Catch
                Exit Function
            End Try
line1:
            acTrans.Commit()
        End Using
    End Function
    '========================================================================================================================
    'функция по локальному idRNG участка ищет мультивыноску
    Public Shared Function FuncFindMLeaderByLocalId(ByVal LocalIDObject As String) As MLeader
        FuncFindMLeaderByLocalId = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is MLeader Then 'если объект имеет тип атрибут
                        Dim acMLeader As MLeader = ent
                        Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(ent)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                If psetDataColl.Count > 0 Then
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            If dataField Like "IDRNG" Then
                                                If data.Trim.Length > 0 Then
                                                    If data.Trim Like LocalIDObject Then
                                                        FuncFindMLeaderByLocalId = acMLeader
                                                        Exit Function
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Next
                        End If
                    End If
                Next
            Catch ex1 As System.Exception
                Exit Function
            End Try
        End Using
    End Function
    '========================================================================================================================
    'функция ищет кадастровый участок
    Public Shared Function FuncFindСadastralParcel(ByVal pointInnerParcel As Point2d, ByVal tablePS As String, Optional cadNumber As String = "") As Autodesk.AutoCAD.DatabaseServices.Polyline
        FuncFindСadastralParcel = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                    Dim acPoly As Polyline = ent
                    Dim vertCadastralBound As Extents3d = acPoly.Bounds
                    Dim maxX As Double = vertCadastralBound.MaxPoint.X + 10
                    Dim maxY As Double = vertCadastralBound.MaxPoint.Y + 10
                    Dim minX As Double = vertCadastralBound.MinPoint.X - 10
                    Dim minY As Double = vertCadastralBound.MinPoint.Y - 10
                    If maxX > pointInnerParcel.X And maxY > pointInnerParcel.Y Then
                        If minX < pointInnerParcel.X And minY < pointInnerParcel.Y Then
                            Dim boolCrossPoint As Boolean = MathFunction.FuncCrossPointPolyline(acPoly, pointInnerParcel)
                            If boolCrossPoint = True Then
                                Dim boolFindTablePS As Boolean = FuncPropertySet.FuncFindPSetTable(acPoly, tablePS)
                                If boolFindTablePS = True Then
                                    If cadNumber.Trim.Length = 0 Then
                                        Return acPoly
                                    Else
                                        Dim cadnum As String = FuncPropertySet.FuncReadValueSetProperties(acPoly, "Cad_number")
                                        Dim posFindCadnum As Integer = cadnum.IndexOf(cadNumber)
                                        If posFindCadnum > -1 Then
                                            Return acPoly
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
    End Function

    '========================================================================================================================
    'функция ищет кадастровые и реестровые номера всех объектов
    Public Shared Function getNumberCadastralObject() As List(Of String)
        Dim result As New List(Of String)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                Dim cadnumber As String = FuncPropertySet.FuncReadValueSetProperties(ent, "Cad_number")
                If cadnumber.Trim.Length = 0 Then
                    cadnumber = FuncPropertySet.FuncReadValueSetProperties(ent, "Rec_numb_border")
                End If
                If cadnumber.Trim.Length = 0 Then
                    result.Add(cadnumber)
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function
    '========================================================================================================================
    'функция по локальному id участка ищет штриховку
    Public Shared Function FuncFindHatchByLocalId(ByVal UserIDObject As String) As Hatch
        FuncFindHatchByLocalId = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is MLeader Then 'если объект имеет тип атрибут
                        Dim acHatch As Hatch = ent
                        Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(ent)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                If psetDataColl.Count > 0 Then
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    For i As Integer = 0 To psetDataColl.Count - 1
                                        Try
                                            pTempData = psetDataColl.Item(i)
                                            Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                            Dim data As String = pTempData.GetData() 'значение
                                            If dataField Like "IDRNG" Then
                                                If data.Trim.Length > 0 Then
                                                    If data.Trim Like UserIDObject Then
                                                        FuncFindHatchByLocalId = acHatch
                                                        Exit Function
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                        Catch ex As System.Exception
                                        End Try
                                    Next i
                                End If
                            Next
                        End If
                    End If
                Next
            Catch
                Exit Function
            End Try
line1:
            acTrans.Commit()
        End Using
    End Function
    '========================================================================================================================
    'функция по глобальному id ищет все контура которые входят в участок (Handle, локальный ID)
    Public Shared Function FuncFindCountersParcel(ByVal UserIDGlobal As String, ByRef ArrayCounterParcel As String(,), ByVal NameTablePS As String) As Boolean
        FuncFindCountersParcel = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Count As Integer = 0
        If IsArray(ArrayCounterParcel) = True Then Count = ArrayCounterParcel.GetUpperBound(1)
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                    Dim acPoly As Polyline = ent
                    Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(acPoly)
                    If setIds.Count > 0 Then
                        For Each psId As ObjectId In setIds
                            Try
                                Dim GLOBID As String = "0"
                                Dim LocalID As String = ""
                                Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                Dim NameTempTBL As String = "NoN"
                                NameTempTBL = pset.PropertySetDefinitionName
                                If NameTablePS Like NameTempTBL Then
                                    If psetDataColl.Count > 0 Then
                                        For j As Integer = 0 To psetDataColl.Count - 1
                                            Try
                                                pTempData = psetDataColl.Item(j)
                                                Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                Dim data As String = pTempData.GetData() 'значение
                                                'считываем значение поля
                                                If dataField Like "GLOBALID" Then
                                                    GLOBID = data
                                                ElseIf dataField Like "LOCALID" Then
                                                    LocalID = data
                                                End If
                                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            Catch ex As System.Exception
                                            End Try
                                        Next j
                                    End If
                                End If
                                'если глобальный id существует то записываем в массив контура
                                If Not (GLOBID Like "0") Then
                                    If GLOBID Like UserIDGlobal Then
                                        ReDim Preserve ArrayCounterParcel(1, Count)
                                        ArrayCounterParcel(0, Count) = acPoly.Handle.ToString
                                        ArrayCounterParcel(1, Count) = LocalID
                                        Count += 1
                                        GoTo line1
                                    End If
                                End If
                            Catch ex1 As System.InvalidCastException
                            Catch ex As System.Exception
                            End Try
                        Next
                    End If
                End If
line1:
            Next
            acTrans.Commit()
        End Using
        FuncFindCountersParcel = True
    End Function
    '========================================================================================================================
    'функция ищет все островки в пределах заданного контура которые ему принадлежат (массив номер контура, Handle) по PSData
    Public Shared Function FuncFindOstrov(ByVal UserIDLocal As String, ByRef ArrayOstrov As String(,)) As Boolean
        FuncFindOstrov = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim count As Integer = 0
        Dim count1 As Integer = 0
        'если массив не пустой
        If IsArray(ArrayOstrov) = True Then
            count = ArrayOstrov.GetUpperBound(1) + 1
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                'находим габаритный контейнер исходной полилиниии
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead) 'открываем объект на чтение
                    If TypeOf ent Is Polyline Then 'если объект имеет тип атрибут
                        Dim acPoly As Polyline = ent
                        Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(acPoly)
                        If setIds.Count > 0 Then
                            For Each psId As ObjectId In setIds
                                Try
                                    Dim UserIDCounter As String = "0"
                                    Dim ConditionNumber As String = ""
                                    Dim LocalID As String = ""
                                    Dim Note As String = ""
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(psId, OpenMode.ForRead, False, False)
                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                    Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                    Dim NameTempTBL As String = "NoN"
                                    NameTempTBL = pset.PropertySetDefinitionName
                                    If NameTempTBL Like "CADASTRAL_OSTROV" Then
                                        If psetDataColl.Count > 0 Then
                                            For j As Integer = 0 To psetDataColl.Count - 1
                                                Try
                                                    pTempData = psetDataColl.Item(j)
                                                    Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                    Dim data As String = pTempData.GetData() 'значение
                                                    'считываем значение поля
                                                    If dataField Like "IDKOUNTER" Then
                                                        UserIDCounter = data
                                                    ElseIf dataField Like "ConditionalNumber" Then
                                                        ConditionNumber = data
                                                    ElseIf dataField Like "LOCALID" Then
                                                        LocalID = data
                                                    ElseIf dataField Like "NOTE" Then
                                                        Note = data
                                                    End If
                                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                                Catch ex As System.Exception
                                                End Try
                                            Next j
                                        End If
                                    End If
                                    'если глобальный id существует то записываем в массив контура
                                    If Not (UserIDCounter Like "0") Then
                                        If UserIDCounter Like UserIDLocal Then
                                            ReDim Preserve ArrayOstrov(4, count)
                                            ArrayOstrov(0, count) = ConditionNumber
                                            ArrayOstrov(1, count) = acPoly.Handle.ToString
                                            ArrayOstrov(2, count) = UserIDCounter
                                            ArrayOstrov(3, count) = LocalID
                                            ArrayOstrov(4, count) = Note
                                            count += 1
                                            GoTo line1
                                        End If
                                    End If
                                Catch ex1 As System.InvalidCastException
                                Catch ex As System.Exception
                                End Try
                            Next
                        End If
                    End If
line1:
                Next
            Catch
                Exit Function
            End Try
            acTrans.Commit()
            FuncFindOstrov = True
        End Using
    End Function
    '========================================================================================================================
    'функция ищет все островки в пределах заданного контура которые ему принадлежат (массив номер контура, Handle) по PSData
    Public Shared Function FuncFindOstrovByDbCollection(ByVal dbObjCollection As DBObjectCollection, ByVal UserIDLocal As String, ByRef ArrayOstrov As String(,)) As Boolean
        FuncFindOstrovByDbCollection = False
        If dbObjCollection.Count > 0 Then
            Dim count As Integer = 0
            For Each acPoly As Polyline In dbObjCollection
                Dim arrayDataPS As String(,) = Nothing
                Dim boolFindDataPs As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(acPoly, arrayDataPS, "CADASTRAL_OSTROV")
                If IsArray(arrayDataPS) = True Then
                    Dim idRng As String = MathFunction.FuncFindValueToArray2d(arrayDataPS, "IDKOUNTER", 1, 2)
                    If idRng.Trim Like UserIDLocal Then
                        ReDim Preserve ArrayOstrov(4, count)
                        ArrayOstrov(0, count) = MathFunction.FuncFindValueToArray2d(arrayDataPS, "ConditionalNumber", 1, 2)
                        ArrayOstrov(1, count) = acPoly.Handle.ToString
                        ArrayOstrov(2, count) = MathFunction.FuncFindValueToArray2d(arrayDataPS, "IDKOUNTER", 1, 2)
                        ArrayOstrov(3, count) = MathFunction.FuncFindValueToArray2d(arrayDataPS, "LOCALID", 1, 2)
                        ArrayOstrov(4, count) = MathFunction.FuncFindValueToArray2d(arrayDataPS, "NOTE", 1, 2)
                        count += 1
                    End If
                End If
            Next
        End If
        If IsArray(ArrayOstrov) = True Then
            If ArrayOstrov.GetUpperBound(1) > 0 Then
                Dim boolSortArray As Boolean = MathFunction.FuncSortStrArray(ArrayOstrov, 0)
            End If
            Return True
        End If
    End Function
    '================================================================================================================================
    'функция заносит в массив все мультивыноски с номерами участков и координатами (массивArrayMLeader: содержимое, X,Y,Handle)
    Public Shared Function FuncFindAllMLeader(ByRef ArrayMLeader As String(,), Optional ByVal ArrayStyleName As String() = Nothing) As Integer
        FuncFindAllMLeader = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If Not IsDBNull(acEnt) Then
                        If TypeOf acEnt Is MLeader Then
                            Dim acMleader As MLeader = acEnt
                            If IsArray(ArrayStyleName) = True Then
                                Dim IDStyleFind As ObjectId = acMleader.MLeaderStyle
                                Dim StyleMleader As MLeaderStyle = IDStyleFind.GetObject(OpenMode.ForRead)
                                Dim boolFindStyle As Boolean = False
                                For i As Integer = 0 To ArrayStyleName.Length - 1
                                    If StyleMleader.Name Like ArrayStyleName(i).Trim Then
                                        boolFindStyle = True
                                        Exit For
                                    End If
                                Next
                                If boolFindStyle = False Then Continue For
                            End If
                            Dim acText As String = acMleader.MText.Contents
                            Dim PosFirstVertex As Point3d = New Point3d(0, 0, 0)
                            If acMleader.LeaderLineCount = 0 Then
                                PosFirstVertex = acMleader.TextLocation
                                ReDim Preserve ArrayMLeader(3, FuncFindAllMLeader)
                                ArrayMLeader(0, FuncFindAllMLeader) = acText
                                ArrayMLeader(1, FuncFindAllMLeader) = PosFirstVertex.X
                                ArrayMLeader(2, FuncFindAllMLeader) = PosFirstVertex.Y
                                ArrayMLeader(3, FuncFindAllMLeader) = acMleader.Handle.ToString
                                FuncFindAllMLeader += 1
                            Else
                                For i As Integer = 0 To acMleader.LeaderLineCount - 1
                                    PosFirstVertex = acMleader.GetFirstVertex(i)
                                    ReDim Preserve ArrayMLeader(3, FuncFindAllMLeader)
                                    ArrayMLeader(0, FuncFindAllMLeader) = acText
                                    ArrayMLeader(1, FuncFindAllMLeader) = PosFirstVertex.X
                                    ArrayMLeader(2, FuncFindAllMLeader) = PosFirstVertex.Y
                                    ArrayMLeader(3, FuncFindAllMLeader) = acMleader.Handle.ToString
                                    FuncFindAllMLeader += 1
                                Next i
                            End If
                        End If
                    End If
                Next
            Catch
                Exit Function
            End Try
            acTrans.Commit()
        End Using
    End Function
    '==============================================================================================================
    'функция составляет поле с площадью полигона с учетом внутренних островков
    Public Shared Function FuncCreateFieldSQRTPolyline(ByVal IDPLine As ObjectId, Optional ByRef ArrayIDOstrov As ObjectId() = Nothing) As String
        FuncCreateFieldSQRTPolyline = ""
        Dim ParcelIDStr As String = ""
        Dim RezStr As String = ""
        If IsNothing(IDPLine) = False Then
            ParcelIDStr = IDPLine.ToString
            If ParcelIDStr.Length < 4 Then
                Exit Function
            End If
        Else
            Exit Function
        End If
        'отделяем скобки
        Dim pos1 As String = ParcelIDStr.ToString.IndexOf("(")
        Dim pos2 As String = ParcelIDStr.ToString.IndexOf(")")
        If pos1 > -1 And pos2 > 1 Then
            ParcelIDStr = Mid(ParcelIDStr, pos1 + 2, pos2 - (pos1 + 1))
            pos1 = -1
            pos2 = -1
        Else
            Exit Function
        End If
        FuncCreateFieldSQRTPolyline = "%<\AcExpr (%<\AcObjProp Object(%<\_ObjId " & ParcelIDStr & ">%).Area>%"
        If IsArray(ArrayIDOstrov) = True Then
            For i As Integer = 0 To ArrayIDOstrov.Length - 1
                Dim strIDOstrov As String = ArrayIDOstrov(i).ToString
                pos1 = strIDOstrov.ToString.IndexOf("(")
                pos2 = strIDOstrov.ToString.IndexOf(")")
                If pos1 > -1 And pos2 > 1 Then
                    strIDOstrov = Mid(strIDOstrov, pos1 + 2, pos2 - (pos1 + 1))
                    FuncCreateFieldSQRTPolyline = FuncCreateFieldSQRTPolyline & "-%<\AcObjProp Object(%<\_ObjId " & strIDOstrov & ">%).Area>%"
                End If
            Next i
            FuncCreateFieldSQRTPolyline = FuncCreateFieldSQRTPolyline & ") \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
            Dim str As String = Chr(34) & Chr(34) & "%lu2%pr0" & Chr(34) & Chr(34)
        Else
            FuncCreateFieldSQRTPolyline = "%<\AcObjProp Object(%<\_ObjId " & ParcelIDStr & ">%).Area \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
        End If
    End Function
    '============================================================================================================================
    'функция отделят полное обозначение участка от его метки на мультивыноске
    Public Shared Function FuncFindNumberCadastralObject(ByVal ConditionalNumber As String) As String
        FuncFindNumberCadastralObject = ConditionalNumber
        Dim TempNumberStr As String = ""
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                ConditionalNumber = ConditionalNumber.Trim
                Dim pos As Integer = ConditionalNumber.IndexOf(Chr(10))
                If pos > -1 Then
                    TempNumberStr = Mid(ConditionalNumber, 1, pos)
                ElseIf ConditionalNumber.IndexOf(Chr(13)) > -1 Then
                    pos = ConditionalNumber.IndexOf(Chr(13))
                    TempNumberStr = Mid(ConditionalNumber, 1, pos)
                ElseIf ConditionalNumber.IndexOf("\pxqc;") > -1 Then
                    Dim pos1 As Integer = ConditionalNumber.IndexOf("\pxqc;")
                    If pos1 > -1 Then
                        TempNumberStr = Mid(ConditionalNumber, pos1 + 6)
                    Else
                        TempNumberStr = ConditionalNumber
                    End If
                Else
                    'Dim pos1 As Integer = ConditionalNumber.IndexOf("/P")
                    Dim pos1 As Integer = ConditionalNumber.IndexOf("\P")
                    If pos1 > -1 Then
                        TempNumberStr = Mid(ConditionalNumber, 1, pos1)
                    Else
                        pos1 = ConditionalNumber.IndexOf("/P")
                        If pos1 > -1 Then
                            TempNumberStr = Mid(ConditionalNumber, 1, pos1)
                        Else
                            TempNumberStr = ConditionalNumber
                        End If
                    End If
                End If
                FuncFindNumberCadastralObject = TempNumberStr
            End If
        End If
    End Function
    '============================================================================================================================
    'функция отделят короткий номер участка от его полного обозначения
    Public Shared Function FuncFindShortNumberCadastralObject(ByVal ConditionalNumber As String) As Integer
        FuncFindShortNumberCadastralObject = 0
        Dim TempNumberStr As String = ""
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                ConditionalNumber = ConditionalNumber.Trim
                Dim pos As Integer = ConditionalNumber.IndexOf(Chr(10))
                If pos > -1 Then
                    TempNumberStr = Mid(ConditionalNumber, 1, pos)
                ElseIf ConditionalNumber.IndexOf(Chr(13)) > -1 Then
                    pos = ConditionalNumber.IndexOf(Chr(13))
                    TempNumberStr = Mid(ConditionalNumber, 1, pos)
                Else
                    Dim pos1 As Integer = ConditionalNumber.IndexOf("\P")
                    pos1 = ConditionalNumber.IndexOf("/P")
                    If pos1 > -1 Then
                        TempNumberStr = Mid(ConditionalNumber, 1, pos1)
                    Else
                        TempNumberStr = ConditionalNumber
                    End If
                End If
                'отделяем кадастровый номер
                Dim pos2 As Integer = TempNumberStr.LastIndexOf(":")
                If pos2 > -1 Then
                    TempNumberStr = Mid(TempNumberStr, pos2 + 2)
                End If

                Dim pos3 As Integer = TempNumberStr.LastIndexOf("ЗУ")
                If pos3 > -1 Then
                    TempNumberStr = Mid(TempNumberStr, pos3 + 3)
                End If

                Dim pos4 As Integer = TempNumberStr.LastIndexOf("(")
                If pos4 > -1 Then
                    TempNumberStr = Mid(TempNumberStr, 1, TempNumberStr.Length - (pos4 + 1))
                End If

                pos4 = TempNumberStr.LastIndexOf("/")
                If pos4 > -1 Then
                    TempNumberStr = Mid(TempNumberStr, 1, TempNumberStr.Length - (pos4 + 1))
                End If

                TempNumberStr = TempNumberStr.Replace("\pxqc;", "")

                pos4 = TempNumberStr.LastIndexOf("\")
                If pos4 > -1 Then
                    TempNumberStr = Mid(TempNumberStr, 1, TempNumberStr.Length - (pos4 + 1))
                End If

                TempNumberStr = TempNumberStr.Replace("Участок", "")
                TempNumberStr = TempNumberStr.Replace("участок", "")
                TempNumberStr = TempNumberStr.Replace("№", "")


                Dim strNumb As String = ""
                For i As Integer = 0 To TempNumberStr.Length - 1
                    Dim str As String = Mid(TempNumberStr, i + 1, 1)
                    If IsNumeric(str) = True Then
                        strNumb = strNumb & str
                    End If
                Next
                FuncFindShortNumberCadastralObject = Val(strNumb)
            End If
        End If
    End Function
    '============================================================================================================================
    'функция возвращает номер контура
    Public Shared Function FuncFindNumberCounterCadastralObject(ByVal ConditionalNumber As String) As Integer
        FuncFindNumberCounterCadastralObject = 0
        Dim TempNumberStr As String = ""
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                ConditionalNumber = ConditionalNumber.Trim
                Dim Pos1 As Integer = ConditionalNumber.LastIndexOf("(")
                Dim Pos2 As Integer = ConditionalNumber.LastIndexOf(")")
                If Pos1 > -1 And Pos2 > -1 Then
                    Dim StrCounter As String = Mid(ConditionalNumber, Pos1 + 2, Pos2 - (Pos1 + 1))
                    Dim TempNumber As String = ""
                    For i As Integer = 1 To StrCounter.Length
                        Dim sumb As String = Mid(StrCounter, i, 1)
                        If IsNumeric(sumb) Then
                            TempNumber = TempNumber & sumb
                        End If
                    Next
                    FuncFindNumberCounterCadastralObject = Val(TempNumber)
                End If
            End If
        End If
    End Function
    '============================================================================================================================
    'функция отделят от обозначения ЗУ его номер контура
    Public Shared Function FuncFindFullConditionalNumberCadastralObject(ByVal ConditionalNumber As String) As String
        FuncFindFullConditionalNumberCadastralObject = ConditionalNumber
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                Try
                    ConditionalNumber = ConditionalNumber.Trim
                    Dim pos1 As Integer = ConditionalNumber.LastIndexOf("(")
                    If pos1 > -1 Then
                        FuncFindFullConditionalNumberCadastralObject = Mid(ConditionalNumber, 1, pos1)
                    End If
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'РАБОТА С ГЕОТОЧКАМИ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '===================================================================================================
    'функция отфильтровывает все геоточки в чертеже и заносит их в коллекцию
    Public Shared Function FuncFilterAllGeoPoint(ByVal arrayNameBlk As String()) As ObjectIdCollection
        Dim GeoPointCoolection As ObjectIdCollection = New ObjectIdCollection
        If IsArray(arrayNameBlk) = False Then
            Return GeoPointCoolection
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim Db As Database = acDoc.Database 'доступ к базе данных
        Dim ed As Editor = acDoc.Editor
        Using Tr As Transaction = Db.TransactionManager.StartTransaction() 'открываем транзакцию
            Dim acBlkTbl As BlockTable = Tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = Tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = Tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                Dim idLr As ObjectId = ent.LayerId
                Dim lyrTemp As LayerTableRecord = idLr.GetObject(OpenMode.ForRead)
                If lyrTemp.IsFrozen = True Then
                    Continue For
                End If
                If lyrTemp.IsOff = True Then
                    Continue For
                End If
                'блок
                If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                    Dim blkPt As BlockReference = ent
                    Dim nameBlk As String = AcadFunction.FuncReadNameBlk(blkPt)
                    Dim boolFindNameBlk As Boolean = MathFunction.FuncFindValueToArray(arrayNameBlk, nameBlk)
                    If boolFindNameBlk = True Then
                        GeoPointCoolection.Add(blkPt.ObjectId)
                    End If

                End If
            Next
        End Using
        Return GeoPointCoolection
    End Function
    '===================================================================================================
    'функция отфильтровывает геоточки в пределах габаритного контейнера полилинии  и заносит их в коллекцию
    Public Shared Function FuncFilterGeoPointByBoundObject(ByVal entObject As Autodesk.AutoCAD.DatabaseServices.Entity, ByVal arrayNameBlk As String()) As ObjectIdCollection
        Dim GeoPointCoolection As ObjectIdCollection = New ObjectIdCollection
        If IsNothing(entObject) = True Then
            Return GeoPointCoolection
        End If
        If IsArray(arrayNameBlk) = False Then
            Return GeoPointCoolection
        End If
        Dim userExtent As Extents3d = entObject.Bounds
        Dim minPoint As Point3d = New Point3d(userExtent.MinPoint.X - 1, userExtent.MinPoint.Y - 1, 0)
        Dim maxPoint As Point3d = New Point3d(userExtent.MaxPoint.X + 1, userExtent.MaxPoint.Y + 1, 0)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim Db As Database = acDoc.Database 'доступ к базе данных
        Dim ed As Editor = acDoc.Editor
        Dim acSSPrompt As PromptSelectionResult = ed.SelectCrossingWindow(minPoint, maxPoint)
        If acSSPrompt.Status = PromptStatus.OK Then
            Using Tr As Transaction = Db.TransactionManager.StartTransaction() 'открываем транзакцию
                Dim acBlkTbl As BlockTable = Tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = Tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = Tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    Dim idLr As ObjectId = ent.LayerId
                    Dim lyrTemp As LayerTableRecord = idLr.GetObject(OpenMode.ForRead)
                    If lyrTemp.IsFrozen = True Then
                        Continue For
                    End If
                    If lyrTemp.IsOff = True Then
                        Continue For
                    End If
                    'блок
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                        Dim blkPt As BlockReference = ent
                        Dim posBlk As Point3d = blkPt.Position
                        If posBlk.X > minPoint.X And posBlk.Y > minPoint.Y And posBlk.X < maxPoint.X And posBlk.Y < maxPoint.Y Then
                            Dim nameUserBlk As String = AcadFunction.FuncReadNameBlk(blkPt)
                            For i As Integer = 0 To arrayNameBlk.Length - 1
                                If arrayNameBlk(i) Like nameUserBlk Then
                                    GeoPointCoolection.Add(blkPt.ObjectId)
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
            End Using
        End If
        Return GeoPointCoolection
    End Function
    '===================================================================================================
    'функция отфильтровывает геоточки в пределах габаритного контейнера полилинии  и заносит их в коллекцию
    Public Shared Function FuncFilterGeoPoint(ByRef ssetObj As SelectionSet, ByVal EntObject As Autodesk.AutoCAD.DatabaseServices.Entity, Optional idLocal As String = "") As ObjectIdCollection
        FuncFilterGeoPoint = New ObjectIdCollection
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim acCurDb As Database = acDoc.Database 'доступ к базе данных
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() 'открываем транзакцию
            Try
                'находим габаритный контейнер
                Dim CoordMin As Point3d = EntObject.GeometricExtents.MinPoint
                Dim CoordMax As Point3d = EntObject.GeometricExtents.MaxPoint
                'перебираем набор
                For Each IDObj As ObjectId In ssetObj.GetObjectIds
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(IDObj, OpenMode.ForRead)
                    'проверяем слой
                    If TypeOf ent Is BlockReference Then
                        Dim blk As BlockReference = ent
                        If blk.Position.X >= CoordMin.X - 1 And blk.Position.Y >= CoordMin.Y - 1 And blk.Position.X <= CoordMax.X + 1 And blk.Position.Y <= CoordMax.Y + 1 Then
                            If idLocal.Trim.Length > 2 Then
                                Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(ent)
                                If setIds.Count > 0 Then
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = acTrans.GetObject(ent.Id, OpenMode.ForRead, False, False)
                                    Dim psetDataColl As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetDataCollection = pset.PropertySetData
                                    If psetDataColl.Count > 0 Then
                                        Dim pTempData As Autodesk.AEC.PropertyData.DatabaseServices.PropertySetData = Nothing
                                        For i As Integer = 0 To psetDataColl.Count - 1
                                            Try
                                                pTempData = psetDataColl.Item(i)
                                                Dim dataField As String = pset.PropertyIdToName(pTempData.Id) 'имя поля
                                                Dim data As String = pTempData.GetData() 'значение
                                                If dataField Like "IDRNG" Then
                                                    If data.Trim Like idLocal Then
                                                        FuncFilterGeoPoint.Add(IDObj)
                                                    End If
                                                End If
                                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            Catch ex As System.Exception
                                            End Try
                                        Next i
                                    End If
                                End If
                            Else
                                FuncFilterGeoPoint.Add(IDObj)
                            End If
                        End If
                    End If
                Next
            Catch
                Exit Function
            End Try
        End Using
    End Function
    '-===================================================================================
    'Поиск наибольшего номера в базе
    Public Shared Function FuncFindLargNumber(ByRef ssetObj As ObjectIdCollection, Optional ByVal namePoint As String = "PROJECT_POINT") As Integer
        Dim largNumber As Integer = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                For Each blkId As ObjectId In ssetObj
                    If Not IsDBNull(blkId) Then
                        'забирам блок ищем атрибут
                        Dim blkRef As BlockReference = acTrans.GetObject(blkId, OpenMode.ForRead)
                        If blkRef.Name Like "namePoint" Then
                            Dim attributeColl As AttributeCollection = blkRef.AttributeCollection
                            If attributeColl.Count > 0 Then
                                For Each attId As ObjectId In attributeColl
                                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                                    If attRef.Tag Like "Имя_точки" Or attRef.Tag Like "NUM_GEOPOINT" Or attRef.Tag Like "NUMBER" Then
                                        Dim numberPointStr As String = attRef.TextString
                                        If IsNothing(numberPointStr) = False Then
                                            numberPointStr = numberPointStr.Trim()
                                        End If
                                        Dim destNumber As Integer = 0
                                        If IsNumeric(numberPointStr) = False Then
                                            destNumber = MathFunction.TryConvertToNumber(numberPointStr)
                                        Else
                                            destNumber = Val(numberPointStr)
                                        End If
                                        If destNumber > largNumber Then
                                            largNumber = destNumber
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            Catch
            End Try
        End Using
        Return largNumber + 1
    End Function
    '-===================================================================================
    'Поиск наибольшего номера в базе
    Public Shared Function FuncReturnBlkToLargNumber(ByVal arrayBlkTag As String(), ByVal arrayNamePoint As String()) As BlockReference
        FuncReturnBlkToLargNumber = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim Db As Database = acDoc.Database
        Dim largNumber As Integer = 0
        Dim blkRez As BlockReference = Nothing
        Using Tr As Transaction = Db.TransactionManager.StartTransaction()
            Try
                Dim acBlkTbl As BlockTable = Tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = Tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = Tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.AutoCAD.DatabaseServices.BlockReference Then
                        Dim blkRef As BlockReference = ent
                        Dim nameBlk As String = AcadFunction.FuncReadNameBlk(blkRef)
                        Dim boolFindValue As Boolean = MathFunction.FuncFindValueToArray(arrayNamePoint, nameBlk)
                        If boolFindValue = True Then
                            Dim attCol As AttributeCollection = blkRef.AttributeCollection
                            For Each attId As ObjectId In attCol
                                Dim attRef As AttributeReference = Tr.GetObject(attId, OpenMode.ForRead)
                                Dim nameField As String = attRef.Tag
                                Dim boolTag As Boolean = MathFunction.FuncFindValueToArray(arrayBlkTag, nameField)
                                If boolTag = True Then
                                    Dim NumPoint As String = attRef.TextString
                                    Dim result As System.Text.RegularExpressions.MatchCollection = New System.Text.RegularExpressions.Regex("[0-9]+").Matches(NumPoint)
                                    If result.Count > 0 Then
                                        Dim value As Integer = Val(result.Item(0).Value)
                                        If value > largNumber Then
                                            blkRez = blkRef
                                            largNumber = value
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Catch
            End Try
        End Using
        Return blkRez
    End Function
    '===========================================================================================
    'функция проверяет есть ли в заданном узле полилинии геоточка и возвращает ее ObjectID
    Public Shared Function FuncFindGeoPoint(ByVal ObjectIDColl As ObjectIdCollection, ByVal pointGeo As Point2d, Optional dopusk As Double = 0) As ObjectId
        FuncFindGeoPoint = Nothing
        Dim Dist As Double = 100000 'начальное значение
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() 'открываем транзакцию
            Try
                For Each IDObj As ObjectId In ObjectIDColl
                    Try
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(IDObj, OpenMode.ForRead)
                        Dim PointPos As Point2d = New Point2d(-1, -1)
                        If TypeOf ent Is BlockReference Then
                            Dim blk As BlockReference = ent
                            Dim attCol As AttributeCollection = blk.AttributeCollection
                            If attCol.Count = 0 Then
                                Continue For
                            End If
                            PointPos = blk.Position.Convert2d(New Plane)
                            ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.CogoPoint Then
                                Dim CGPoint As Autodesk.Civil.DatabaseServices.CogoPoint = ent
                                PointPos = New Point2d(CGPoint.Easting, CGPoint.Northing)
                            ElseIf TypeOf ent Is MLeader Then
                                Dim acMleader As MLeader = ent
                                PointPos = acMleader.GetFirstVertex(0).Convert2d(New Plane)
                            ElseIf TypeOf ent Is Autodesk.Civil.DatabaseServices.NoteLabel Then
                                Dim CGPoint As Autodesk.Civil.DatabaseServices.NoteLabel = ent
                            PointPos = CGPoint.LabelLocation.Convert2d(New Plane)
                        End If
                        Dim LenDist As Double = PointPos.GetVectorTo(pointGeo).Length
                        If LenDist = 0 Then
                            FuncFindGeoPoint = ent.ObjectId
                            Exit Function
                        Else
                            If LenDist < Dist Then
                                Dist = LenDist
                                FuncFindGeoPoint = ent.ObjectId
                            End If
                        End If
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                Next
                If Dist > dopusk Then
                    FuncFindGeoPoint = Nothing
                End If
            Catch ex As System.Exception
                acTrans.Commit()
                FuncFindGeoPoint = Nothing
                Exit Function
            End Try
            acTrans.Commit()
        End Using
    End Function
    '===========================================================================================
    'функция возвращает номера поворотных точек узлов объекта
    Public Shared Function FuncReturnGeoPointInPolyline(ByVal GeoPointCollection As ObjectIdCollection, ByVal ObjPline As Polyline, Optional dopusk As Double = 0) As String
        FuncReturnGeoPointInPolyline = ""
        Dim StrNumber As String = ""
        Dim Dist As Double = 100000 'начальное значение
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim acCurDb As Database = acDoc.Database
        If GeoPointCollection.Count = 0 Then
            Exit Function
        End If
        If Not (TypeOf ObjPline Is Polyline) Then
            Exit Function
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() 'открываем транзакцию
            Try
                For i As Integer = 0 To ObjPline.NumberOfVertices - 1
                    Dim VertPlineXY As Point2d = ObjPline.GetPoint2dAt(i)
                    Dim blk As BlockReference = Nothing
                    For Each IDObj As ObjectId In GeoPointCollection
                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(IDObj, OpenMode.ForRead)
                        If TypeOf ent Is BlockReference Then
                            Dim Tempblk As BlockReference = ent
                            If (Tempblk.Position.X = VertPlineXY.X And Tempblk.Position.Y = VertPlineXY.Y) Then
                                blk = Tempblk
                                Exit For
                            Else
                                Dim Lenght As Double = Math.Sqrt((Tempblk.Position.X - VertPlineXY.X) ^ 2 + (Tempblk.Position.Y - VertPlineXY.Y) ^ 2)
                                If Lenght < Dist Then
                                    Dist = Lenght
                                    If Dist <= dopusk Then
                                        blk = Tempblk
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If IsNothing(blk) = False Then
                        If blk.Id.IsNull = False Then
                            Dim attCol As AttributeCollection = blk.AttributeCollection
                            For Each attId As ObjectId In attCol
                                Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                                If attRef.Tag Like "NUM_GEOPOINT" Or attRef.Tag Like "Имя_точки" Or attRef.Tag Like "NUMBER" Then
                                    If StrNumber.Trim.Length = 0 Then
                                        StrNumber = attRef.TextString
                                    Else
                                        StrNumber = StrNumber & "," & attRef.TextString
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next i
            Catch
                acTrans.Commit()
            End Try
            acTrans.Commit()
        End Using
        FuncReturnGeoPointInPolyline = StrNumber
    End Function
    '-===================================================================================
    'Функция возвращает номер, погрешность и тип закрепления точки
    Public Shared Function FuncReturnBlkArtibuteGeoPoint(ByVal blkPoint As BlockReference, ByRef ArrayGeoPoint As String()) As Boolean
        FuncReturnBlkArtibuteGeoPoint = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                If attCol.Count > 0 Then
                    ReDim ArrayGeoPoint(3)
                    ArrayGeoPoint(0) = ""
                    ArrayGeoPoint(1) = ""
                    ArrayGeoPoint(2) = ""
                    ArrayGeoPoint(3) = ""
                    For Each attId As ObjectId In attCol
                        Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                        If attRef.Tag Like "Имя_точки" Then
                            ArrayGeoPoint(0) = attRef.TextString
                        ElseIf attRef.Tag Like "NUM_GEOPOINT" Then
                            ArrayGeoPoint(0) = attRef.TextString
                        ElseIf attRef.Tag Like "NUMBER" Then
                            ArrayGeoPoint(0) = attRef.TextString
                        ElseIf attRef.Tag Like "Точность_положения" Then
                            ArrayGeoPoint(1) = attRef.TextString
                        ElseIf attRef.Tag Like "DELTA" Then
                            ArrayGeoPoint(1) = attRef.TextString
                        ElseIf attRef.Tag Like "Описание_закрепления" Then
                            ArrayGeoPoint(2) = attRef.TextString
                        ElseIf attRef.Tag Like "FASTDESK" Then
                            ArrayGeoPoint(2) = attRef.TextString
                        ElseIf attRef.Tag Like "Метод_определения_координат" Then
                            ArrayGeoPoint(3) = attRef.TextString
                        ElseIf attRef.Tag Like "METHODCOORDINATES" Then
                            ArrayGeoPoint(3) = attRef.TextString
                        End If
                    Next
                End If
            Catch ex As System.Exception
                FuncReturnBlkArtibuteGeoPoint = False
                Exit Function
            End Try
            FuncReturnBlkArtibuteGeoPoint = True
        End Using
    End Function
    '=============================================================================================
    'функция удаляет коллекцию точек (0-все,
    Public Shared Function FuncDeleteGeoPoint(ByRef pointIDCollection As ObjectIdCollection, ByVal parametr As Integer) As Integer
        FuncDeleteGeoPoint = 0
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument 'доступ к активному чертежу
        Dim acCurDb As Database = acDoc.Database 'доступ к базе данных
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() 'открываем транзакцию
            Try
                For Each idObjBlk As ObjectId In pointIDCollection
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(idObjBlk, OpenMode.ForWrite)
                    If TypeOf ent Is BlockReference Then
                        If parametr = 0 Then
                            ent.Erase()
                            FuncDeleteGeoPoint += 1
                        Else
                            Dim blk As BlockReference = ent
                            Dim attCol As AttributeCollection = blk.AttributeCollection
                            For Each attId As ObjectId In attCol
                                Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                                If attRef.Tag Like "Имя_точки" Then
                                    Dim NumPoint2 As String = attRef.TextString
                                    If parametr = 1 Then
                                        Dim pos As Integer = NumPoint2.IndexOf("н")
                                        If pos = -1 Then
                                            pos = NumPoint2.IndexOf("Н")
                                        End If
                                        If pos <> -1 Then
                                            ent.Erase()
                                            FuncDeleteGeoPoint += 1
                                        End If
                                    ElseIf parametr = 2 Then
                                        If IsNumeric(NumPoint2) = True Then
                                            ent.Erase()
                                            FuncDeleteGeoPoint += 1
                                        End If
                                    ElseIf parametr = 2 Then
                                        Dim pos As Integer = NumPoint2.IndexOf("%")
                                        If pos <> -1 Then
                                            ent.Erase()
                                            FuncDeleteGeoPoint += 1
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Catch
                acTrans.Commit()
                Exit Function
            End Try
            acTrans.Commit()
        End Using
    End Function

    '
End Class
