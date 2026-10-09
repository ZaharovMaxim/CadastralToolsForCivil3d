Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Buffer
Imports System.Xml

Public Class BufferZoneCreator
    Private ReadOnly _geometryFactory As GeometryFactory
    Public Sub New()
        _geometryFactory = New GeometryFactory()
    End Sub
    ' Создает буферную зону вокруг точки
    ' </summary>
    ' <param name="x">Координата X</param>
    ' <param name="y">Координата Y</param>
    ' <param name="bufferDistance">Расстояние буфера</param>
    ' <returns>Буферная зона как полигон</returns>
    Public Function CreatePointBuffer(x As Double, y As Double, bufferDistance As Double) As Polygon
        Dim point = _geometryFactory.CreatePoint(New Coordinate(x, y))
        Return CType(point.Buffer(bufferDistance), Polygon)
    End Function

    ' Создает буферную зону вокруг линии
    ' </summary>
    ' <param name="coordinates">Координаты линии</param>
    ' <param name="bufferDistance">Расстояние буфера</param>
    ' <returns>Буферная зона как полигон</returns>
    Public Function CreateLineBuffer(coordinates As List(Of Coordinate), bufferDistance As Double) As Polygon
        Dim lineString = _geometryFactory.CreateLineString(coordinates.ToArray())
        Return CType(lineString.Buffer(bufferDistance), Polygon)
    End Function

    ' Создает буферную зону вокруг полигона
    ' </summary>
    ' <param name="coordinates">Координаты полигона</param>
    ' <param name="bufferDistance">Расстояние буфера</param>
    ' <returns>Буферная зона как полигон</returns>
    Public Function CreatePolygonBuffer(coordinates As List(Of Coordinate), bufferDistance As Double) As Polygon
        ' Создаем линейное кольцо (первое и последнее координаты должны совпадать)
        If coordinates.Count > 0 AndAlso Not coordinates(0).Equals(coordinates(coordinates.Count - 1)) Then
            coordinates.Add(New Coordinate(coordinates(0)))
        End If
        Dim linearRing = _geometryFactory.CreateLinearRing(coordinates.ToArray())
        Dim polygon = _geometryFactory.CreatePolygon(linearRing)
        Return CType(polygon.Buffer(bufferDistance), Polygon)
    End Function

    ' Создает буферную зону с дополнительными параметрами
    ' </summary>
    ' <param name="geometry">Исходная геометрия</param>
    ' <param name="bufferDistance">Расстояние буфера</param>
    ' <param name="quadrantSegments">Количество сегментов для аппроксимации дуг</param>
    ' <param name="endCapStyle">Стиль окончаний</param>
    ' <returns>Буферная зона</returns>
    Public Function CreateCustomBuffer(geometry As Geometry, bufferDistance As Double,
                                     Optional quadrantSegments As Integer = 8,
                                     Optional endCapStyle As EndCapStyle = EndCapStyle.Round) As Geometry

        Dim bufferParameters = New BufferParameters With {
            .QuadrantSegments = quadrantSegments,
            .EndCapStyle = endCapStyle
        }

        Dim bufferOp = New BufferOp(geometry, bufferParameters)
        Return bufferOp.GetResultGeometry(bufferDistance)
    End Function

    ' Экспортирует геометрию в формат WKT
    ' </summary>
    ' <param name="geometry">Геометрия для экспорта</param>
    ' <returns>Строка в формате WKT</returns>
    Public Function ExportToWKT(geometry As Geometry) As String
        Dim writer = New WKTWriter()
        Return writer.Write(geometry)
    End Function

    ' Импортирует геометрию из формата WKT
    ' </summary>
    ' <param name="wkt">Строка в формате WKT</param>
    ' <returns>Геометрия</returns>
    Public Function ImportFromWKT(wkt As String) As Geometry
        Dim reader = New WKTReader()
        Return reader.Read(wkt)
    End Function

    ' Вычисляет площадь буферной зоны
    ' </summary>
    ' <param name="buffer">Буферная зона</param>
    ' <returns>Площадь</returns>
    Public Function CalculateBufferArea(buffer As Geometry) As Double
        Return buffer.Area
    End Function

    'вычитание полигонов
    Public Function SubtractPolygons(ByVal polygon1 As Polygon, ByVal polygon2 As Polygon) As Geometry
        ' Вычитание polygon2 из polygon1
        Return polygon1.Difference(polygon2)
    End Function
End Class
Public Class LandCadastral
    Private _netPolygon As Polygon
    Private _arrayData As String(,)
    ' Конструктор
    Public Sub New(ByVal netPolygon As Polygon, ByVal arrayData As String(,))
        _netPolygon = netPolygon
        _arrayData = arrayData
    End Sub
    ' Свойство thicknessLayer
    Public Property netPolygon() As Polygon
        Get
            Return _netPolygon
        End Get
        Set(ByVal value As Polygon)
            _netPolygon = value
        End Set
    End Property

    ' Свойство badThicknessLayer
    Public Property arrayData() As String(,)
        Get
            Return _arrayData
        End Get
        Set(ByVal value As String(,))
            _arrayData = value
        End Set
    End Property
End Class
Public Class LandReclamation
    Private _condinionalNumber As String
    Private _area As Double
    Private _thicknessLayer As Double
    Private _badThicknessLayer As Double
    Private _listMineralFertilizers As List(Of Dictionary(Of String, Double)) 'наименование, норма внесения
    Private _listOrganicalFertilizers As List(Of Dictionary(Of String, Double)) 'наименование, норма внесения
    Private _listAmeliorant As List(Of Dictionary(Of String, Double)) 'наименование, норма внесения
    Private _greenSiderats As List(Of Dictionary(Of String, Double)) 'наименование, норма внесения

    ' Конструктор
    Public Sub New(ByVal Area As Double, ByVal thLayer As Double, ByVal badThLayer As Double)
        _condinionalNumber = 1
        _area = Area
        _thicknessLayer = thLayer
        _badThicknessLayer = badThLayer
        _listMineralFertilizers = New List(Of Dictionary(Of String, Double))
        _listOrganicalFertilizers = New List(Of Dictionary(Of String, Double))
        _listAmeliorant = New List(Of Dictionary(Of String, Double))
        _greenSiderats = New List(Of Dictionary(Of String, Double))
    End Sub

    ' Конструктор по умолчанию (для JSON десериализации)
    Public Sub New()
    End Sub

    ' Свойство Площади
    Public Property ConditionatNumber() As String
        Get
            Return _condinionalNumber
        End Get
        Set(ByVal value As String)
            _condinionalNumber = value
        End Set
    End Property

    ' Свойство Площади
    Public Property Area() As Double
        Get
            Return _area
        End Get
        Set(ByVal value As Double)
            _area = value
        End Set
    End Property
    ' Свойство thicknessLayer
    Public Property thicknessLayer() As Double
        Get
            Return _thicknessLayer
        End Get
        Set(ByVal value As Double)
            _thicknessLayer = value
        End Set
    End Property

    ' Свойство badThicknessLayer
    Public Property badThicknessLayer() As Double
        Get
            Return _badThicknessLayer
        End Get
        Set(ByVal value As Double)
            _badThicknessLayer = value
        End Set
    End Property

    ' Свойство мелиорантов
    Public Property Ameliorant() As List(Of Dictionary(Of String, Double))
        Get
            Return _listAmeliorant
        End Get
        Set(ByVal value As List(Of Dictionary(Of String, Double)))
            _listAmeliorant = value
        End Set
    End Property

    ' Свойство минеральные удобрения
    Public Property MineralFertilizers() As List(Of Dictionary(Of String, Double))
        Get
            Return _listMineralFertilizers
        End Get
        Set(ByVal value As List(Of Dictionary(Of String, Double)))
            _listMineralFertilizers = value
        End Set
    End Property

    ' Свойство органические удобрения
    Public Property OrganicalFertilizers() As List(Of Dictionary(Of String, Double))
        Get
            Return _listOrganicalFertilizers
        End Get
        Set(ByVal value As List(Of Dictionary(Of String, Double)))
            _listOrganicalFertilizers = value
        End Set
    End Property

    ' Свойство газонные травы
    Public Property GreenSiderats() As List(Of Dictionary(Of String, Double))
        Get
            Return _greenSiderats
        End Get
        Set(ByVal value As List(Of Dictionary(Of String, Double)))
            _greenSiderats = value
        End Set
    End Property
    'функция читает схемы проведения работ по рекультивации
    Public Function ParseYearReclamationXml(ByVal xmlFilePath As String, ByVal stageName As String, ByVal schemeId As String) As Integer
        Try
            ' Проверка существования файла
            If Not File.Exists(xmlFilePath) Then
                Return 0
            End If
            ' Загрузка XML документа
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(xmlFilePath)
            ' Построение XPath запроса для поиска схемы
            Dim xpath As String = $"/LandReclamation/{stageName}/scheme[@id='{schemeId}']"
            ' Поиск узла схемы
            Dim schemeNode As XmlNode = xmlDoc.SelectSingleNode(xpath)
            If schemeNode Is Nothing Then
                ' Схема не найдена
                Return 0
            End If
            ' Проверяем наличие атрибута NumberOfYears
            If schemeNode.Attributes IsNot Nothing AndAlso
           schemeNode.Attributes("NumberOfYears") IsNot Nothing Then
                ' Пытаемся преобразовать значение в Integer
                Dim yearsValue As String = schemeNode.Attributes("NumberOfYears").Value
                Dim result As Integer = 0
                If Integer.TryParse(yearsValue, result) Then
                    Return result
                Else
                    ' Если не удалось преобразовать, возвращаем 0
                    Return 0
                End If
            Else
                ' Атрибут не найден
                Return 0
            End If
        Catch ex As XmlException
            Return 0
        Catch ex As Exception
            Return 0
        End Try
    End Function
    'функция читает схемы проведения работ по рекультивации
    Public Function ParseReclamationXml(ByVal xmlFilePath As String, ByVal stageName As String, ByVal schemeId As String, ByVal numberYear As Integer) As List(Of Dictionary(Of String, String))
        Dim result As New List(Of Dictionary(Of String, String))()
        Try
            ' Проверка существования файла
            If Not File.Exists(xmlFilePath) Then
                Return result
            End If
            ' Загрузка XML документа
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(xmlFilePath)
            ' Построение XPath запроса
            Dim xpath As String = ""
            If numberYear <= 0 Then
                ' Поиск всех work без учета года (для схем без вложенных year)
                xpath = $"/LandReclamation/{stageName}/scheme[@id='{schemeId}']/work"
            Else
                ' Поиск work внутри конкретного года
                xpath = $"/LandReclamation/{stageName}/scheme[@id='{schemeId}']/year[@number='{numberYear}']/work"
            End If

            ' Выборка всех узлов work
            Dim workNodes As XmlNodeList = xmlDoc.SelectNodes(xpath)
            If workNodes Is Nothing OrElse workNodes.Count = 0 Then
                ' Возвращаем пустой список, если ничего не найдено
                Return result
            End If
            ' Обход каждого узла work
            For Each workNode As XmlNode In workNodes
                Dim workDict As New Dictionary(Of String, String)()
                ' Добавляем атрибуты узла
                If workNode.Attributes IsNot Nothing Then
                    For Each attr As XmlAttribute In workNode.Attributes
                        workDict(attr.Name) = attr.Value
                    Next
                End If
                ' Добавляем текст узла (если есть)
                If Not String.IsNullOrWhiteSpace(workNode.InnerText) Then
                    workDict("InnerText") = workNode.InnerText.Trim()
                End If
                ' Добавляем имя узла для идентификации
                workDict("NodeName") = workNode.Name
                result.Add(workDict)
            Next
        Catch ex As XmlException
            Return result
        Catch ex As Exception
            Return result
        End Try
        Return result
    End Function
    'функция считает значения
    Public Function calculateValue(ByVal formula As String, ByVal arrayParcelReclamation As LandReclamation(), ByRef parameters As Dictionary(Of String, String)) As String(,)
        Dim result As String(,) = {}
        Dim countResult As Integer = 0
        Dim strFormula As String = ""
        Dim nameWork As String = ""
        Dim value As Double = 0
        If IsNothing(formula) = True Then Return result
        If IsNothing(arrayParcelReclamation) = True Then Return result
        If IsArray(arrayParcelReclamation) = False Then Return result
        If arrayParcelReclamation.Length = 0 Then Return result
        If formula Like "BadThLayer*Area" Then
            ReDim result(3, 0)
            result(0, 0) = ""
            result(1, 0) = "0"
            result(2, 0) = "0"
        ElseIf formula Like "ThLayer*Area" Then
            ReDim result(3, 0)
            result(0, 0) = ""
            result(1, 0) = "0"
            result(2, 0) = "0"
        ElseIf formula Like "Area" Then
            ReDim result(3, 0)
            result(0, 0) = ""
            result(1, 0) = "0"
            result(2, 0) = "0"
        ElseIf formula Like "Area*ApplicationRate" Then
            If parameters.Count > 0 Then
                If parameters.ContainsKey("MineralFertilizers") Then
                    Dim nameMinerale As String = parameters.Item("MineralFertilizers")
                    Dim arrayStr As String() = nameMinerale.Split(";")
                    If arrayStr.Length > 0 Then
                        ReDim result(3, arrayStr.Length - 1)
                        For j As Integer = 0 To result.GetUpperBound(1)
                            result(0, j) = arrayStr(j)
                        Next j
                    End If
                ElseIf parameters.ContainsKey("OrganicalFertilizers") Then
                    Dim nameMinerale As String = parameters.Item("OrganicalFertilizers")
                    Dim arrayStr As String() = nameMinerale.Split(";")
                    If arrayStr.Length > 0 Then
                        ReDim result(3, arrayStr.Length - 1)
                        For j As Integer = 0 To result.GetUpperBound(1)
                            result(0, j) = arrayStr(j)
                        Next j
                    End If
                ElseIf parameters.ContainsKey("GreenSiderat") Then
                    Dim nameMinerale As String = parameters.Item("GreenSiderat")
                    Dim arrayStr As String() = nameMinerale.Split(";")
                    If arrayStr.Length > 0 Then
                        ReDim result(3, arrayStr.Length - 1)
                        For j As Integer = 0 To result.GetUpperBound(1)
                            result(0, j) = arrayStr(j)
                        Next j
                    End If
                ElseIf parameters.ContainsKey("Ameliorant") Then
                    Dim nameMinerale As String = parameters.Item("Ameliorant")
                    Dim arrayStr As String() = nameMinerale.Split(";")
                    If arrayStr.Length > 0 Then
                        ReDim result(3, arrayStr.Length - 1)
                        For j As Integer = 0 To result.GetUpperBound(1)
                            result(0, j) = arrayStr(j)
                        Next j
                    End If
                End If
                If result.Length > 0 Then
                    If parameters.ContainsKey("ApplicationRate") Then
                        Dim rates As String = parameters.Item("ApplicationRate")
                        Dim arrayStr As String() = rates.Split(";")
                        If arrayStr.Length > 0 And result.GetUpperBound(1) = arrayStr.Length - 1 Then
                            For j As Integer = 0 To result.GetUpperBound(1)
                                result(1, j) = arrayStr(j)
                            Next j
                        End If
                    End If
                    result(2, 0) = "0"
                    result(3, 0) = ""
                End If
            End If
        End If
        For i As Integer = 0 To arrayParcelReclamation.Length - 1
            Dim parcelReclamation As LandReclamation = arrayParcelReclamation(i)
            Dim areaParcel As Double = Math.Round(parcelReclamation.Area, 0)
            If formula Like "BadThLayer*Area" Then
                value += areaParcel * parcelReclamation.badThicknessLayer
                If i = 0 Then
                    strFormula += areaParcel & "*" & parcelReclamation.badThicknessLayer
                ElseIf i = arrayParcelReclamation.Length - 1 Then
                    strFormula += "+" & areaParcel & "*" & parcelReclamation.badThicknessLayer
                Else
                    strFormula += "+" & areaParcel & "*" & parcelReclamation.badThicknessLayer
                End If
                result(1, 0) = Math.Round(parcelReclamation.badThicknessLayer, 2)
                result(2, 0) = Math.Round(value, 2)
                result(3, 0) = strFormula & "=" & value
            ElseIf formula Like "ThLayer*Area" Then
                value += areaParcel * parcelReclamation.thicknessLayer
                If i = 0 Then
                    strFormula += areaParcel & "*" & parcelReclamation.thicknessLayer
                ElseIf i = arrayParcelReclamation.Length - 1 Then
                    strFormula += "+" & areaParcel & "*" & parcelReclamation.thicknessLayer
                Else
                    strFormula += "+" & areaParcel & "*" & parcelReclamation.thicknessLayer
                End If
                result(1, 0) = Math.Round(parcelReclamation.thicknessLayer, 2)
                result(2, 0) = Math.Round(value, 2)
                result(3, 0) = strFormula & "=" & value
            ElseIf formula Like "Area" Then
                value += areaParcel
                If i = 0 Then
                    strFormula += CStr(areaParcel)
                ElseIf i = arrayParcelReclamation.Length - 1 Then
                    strFormula += "+" & areaParcel
                Else
                    strFormula += "+" & areaParcel
                End If
                result(2, 0) = Math.Round(value, 2)
                result(3, 0) = strFormula & "=" & value
            ElseIf formula Like "Area*ApplicationRate" Then
                areaParcel /= 10000
                If result.Length > 0 Then
                    For j As Integer = 0 To result.GetUpperBound(1)
                        Dim rez As Double = Val(result(1, j)) * areaParcel + Val(result(2, j))
                        Dim tempstrFormula As String = ""
                        If i = 0 Then
                            tempstrFormula = areaParcel & "*" & result(1, j)
                        Else
                            tempstrFormula = result(3, j) & "+" & areaParcel & "*" & result(1, j)
                        End If
                        result(2, j) = Math.Round(rez, 2)
                        result(3, j) = tempstrFormula
                    Next j

                End If
            End If
        Next i
        If formula Like "Area*ApplicationRate" Then
            For j As Integer = 0 To result.GetUpperBound(1)
                result(3, j) = result(3, j) & "=" & result(2, j)
            Next j
        End If

        Return result
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '1. автоматическое создание участков рекультивации
    <CommandMethod("CreateReclamationParcels")>
    Public Sub CreateReclamationParcels()
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rng As Random = New Random
        Dim CategoryObject As String = "Благоустройство и озеленение территории"
        Dim userCategoryObject As String = "Рекультивация земельных участков"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к временному катологу
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '3. Путь к файлам шаблонам с условными знами
        Dim putchTemplate As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim arrayTemplates As String() = {}
        Dim countArrayPTemplates As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(putchTemplate)
                Try
                    Dim File1 As String = IO.Path.GetFileName(fileDwg)
                    Dim ext As String = IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve arrayTemplates(countArrayPTemplates)
                        arrayTemplates(countArrayPTemplates) = Path.GetFileNameWithoutExtension(File1)
                        countArrayPTemplates += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If

        Dim processor As New PolygonIntersectionInnerProcessor()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'считываем установки
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim templateXML As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            putchTemplate = parentPatch2 & "\FileResources\Sample\"
            templateXML = putchTemplate & "Стандарт_ADSK.xml"
            If File.Exists(templateXML) = False Then
                MsgBox("Шаблон оформления чертежей не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If
        Dim DeskLine As String = "Участок подлежащий рекультивации"
        'свойства линии
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim TypeUserLine As String = "Continuous"
            Dim LayerLine As String = "0"
            Dim WidthUserLine As Double = 30
            Dim ColorUserLine As Integer = 10
            Dim CondUserLile As Double = 0
            Dim NameTableLinePS As String = ""
            Dim ArrayParcelProperties As String(,) = Nothing
            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            Dim acLineTypTbl As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IO.File.Exists(templateXML) = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'загружаем типы линий и таблицу PS
                Dim boolFindPropertiesObject As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(templateXML, CategoryObject, userCategoryObject, "Полигон", DeskLine, ArrayParcelProperties)
                If IsArray(ArrayParcelProperties) = True Then
                    For i As Integer = 0 To ArrayParcelProperties.GetUpperBound(1)
                        If ArrayParcelProperties(0, i) Like "Name" Then
                            Dim NameTypeLine As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameTypeLine) = False Then
                                If acLineTypTbl.Has(NameTypeLine) = False Then
                                    Dim recTypeline As LinetypeTableRecord = TOPOSymbols.FuncCreateLineTypeToXML(templateXML, NameTypeLine)
                                    'Dim boolCopyTypeLineStyle As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, NameTypeLine)
                                    TypeUserLine = NameTypeLine
                                Else
                                    TypeUserLine = NameTypeLine
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameLayerTemp) = False Then
                                If acLyrTbl.Has(NameLayerTemp) = False Then
                                    Dim idLayer As ObjectId = TOPOSymbols.FuncCreateLayerToXML(templateXML, NameLayerTemp)
                                    'Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, NameLayerTemp)
                                    LayerLine = NameLayerTemp
                                Else
                                    LayerLine = NameLayerTemp
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "PS" Then
                            Dim NameTablePSTemp As String = ArrayParcelProperties(1, i)
                            If IsNothing(NameTablePSTemp) = False Then
                                If dictPropSetDefTemp.Has(NameTablePSTemp, tr) = False Then
                                    Dim idPs As ObjectId = FuncPropertySet.FuncCreateTablePSByXML(templateXML, NameTablePSTemp)
                                    'Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NameTablePSTemp)
                                    NameTableLinePS = NameTablePSTemp
                                Else
                                    NameTableLinePS = NameTablePSTemp
                                End If
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayParcelProperties(1, i)
                            If IndexColor Like "ПоСлою" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                ColorUserLine = 0
                            ElseIf IndexColor Like "ByLayer" Then
                                ColorUserLine = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) < 1 Then
                                ColorUserLine = 0
                            ElseIf Val(IndexColor) > 255 Then
                                ColorUserLine = 256
                            Else
                                ColorUserLine = Val(IndexColor)
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineWidth" Then
                            Dim WidthLine As String = ArrayParcelProperties(1, i)
                            If WidthLine Like "ПоСлою" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ПоБлоку" Then
                                WidthUserLine = LineWeight.ByBlock
                            ElseIf WidthLine Like "ByLayer" Then
                                WidthUserLine = LineWeight.ByLayer
                            ElseIf WidthLine Like "ByBlock" Then
                                WidthUserLine = LineWeight.ByBlock
                            Else
                                Try
                                    WidthUserLine = Val(WidthLine)
                                    WidthUserLine = Math.Abs(WidthUserLine)
                                Catch ex As ArgumentException
                                End Try
                            End If
                        ElseIf ArrayParcelProperties(0, i) Like "LineGWidth" Then
                            CondUserLile = Val(ArrayParcelProperties(1, i))
                        End If
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'задаем номер участка
            Dim pStrOptsNumb As PromptIntegerOptions = New PromptIntegerOptions(vbLf & "Введите начальный номер участка: ")
            Dim pStrResNumb As PromptIntegerResult = acDoc.Editor.GetInteger(pStrOptsNumb)
            Dim numbParcel As Integer = pStrResNumb.Value

            Dim NameCadastralTable As String = "LAND_RECLAMATION"
            Dim ArrayDataObj As String(,) = Nothing
            Dim boolFindFieldsArray As Boolean = FuncPropertySet.FuncReadAllValueFieldSetTables(NameTableLinePS, ArrayDataObj)
            Dim idTablePs As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(NameCadastralTable)

            Dim listCadastralParcels As List(Of LandCadastral) = New List(Of LandCadastral)
            Dim listCadastralKvartal As List(Of LandCadastral) = New List(Of LandCadastral)
            Dim NameCadastralTableParcels As String = "LAND_CADASTRAL_PARCEL"
            Dim NameCadastralTableKvartals As String = "LAND_CADASTRAL_DIVISION"
            Dim putchNameKPT As String = ""
            Dim netTopo As FuncNetTopologySuite = New FuncNetTopologySuite
            Dim rs As PromptNestedEntityResult = ed.GetNestedEntity(vbLf & "Выберите кадастровый план: ")
            If Not (rs.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim nameLayer As String = "0"
            If rs.Status = PromptStatus.OK Then
                Dim ids As List(Of ObjectId) = New List(Of ObjectId)(rs.GetContainers())
                ids.Reverse()
                ids.Add(rs.ObjectId)
                For Each id As ObjectId In ids
                    Dim idStr As String = id.ObjectClass.Name
                    If idStr Like "AcDbPolyline" Then
                        Dim nameFileKPT As String = id.Database.Filename
                        Using selectObject As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(id.Open(OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                            nameLayer = selectObject.Layer
                            Dim strLayerArr As String() = nameLayer.Split("|")
                            If strLayerArr.Length > 1 Then
                                nameLayer = strLayerArr(1)
                            End If
                            If IO.File.Exists(nameFileKPT) = True Then
                                putchNameKPT = nameFileKPT
                                Exit For
                            End If
                        End Using
                    End If
                Next
            End If
            Dim sourceDb As Database = New Database(False, True)
            sourceDb.ReadDwgFile(putchNameKPT, FileShare.Read, True, String.Empty)
            Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
            Using sourceTr
                Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                For Each idObject As ObjectId In sourceBtr
                    If idObject.IsNull = False Then
                        Dim idStr As String = idObject.ObjectClass.Name
                        If idStr Like "AcDbPolyline" Then
                            Dim poly As Polyline = sourceTr.GetObject(idObject, OpenMode.ForRead)
                            Dim entObject As Autodesk.AutoCAD.DatabaseServices.Entity = sourceTr.GetObject(idObject, OpenMode.ForRead) 'открываем объект на чтение
                            Dim setIds As ObjectIdCollection = Autodesk.AEC.PropertyData.DatabaseServices.PropertyDataServices.GetPropertySets(entObject)
                            If setIds.Count > 0 Then
                                For Each psId As ObjectId In setIds
                                    Dim pset As Autodesk.AEC.PropertyData.DatabaseServices.PropertySet = sourceTr.GetObject(psId, OpenMode.ForRead, False, False)
                                    Dim nameTable As String = pset.PropertySetDefinitionName
                                    If nameTable Like "LAND_CADASTRAL_PARCEL" Then
                                        Dim secondPolygon As Polygon = netTopo.CreatePolygon(poly)
                                        'перебираем список с полигонами
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
                                            Dim userLand As LandCadastral = New LandCadastral(secondPolygon, arrayData)
                                            listCadastralParcels.Add(userLand)
                                        End If
                                    ElseIf nameTable Like "LAND_CADASTRAL_DIVISION" Then
                                        Dim secondPolygon As Polygon = netTopo.CreatePolygon(poly)
                                        'перебираем список с полигонами
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
                                            Dim userLand As LandCadastral = New LandCadastral(secondPolygon, arrayData)
                                            listCadastralKvartal.Add(userLand)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End Using

            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            'выбираем объекты
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim userPoly As Polyline = acEnt
                                    If userPoly.Area > 0 Then
                                        Dim userPolygon As Polygon = netTopo.CreatePolygon(userPoly)
                                        Dim boolFindParcels As Boolean = False
                                        If listCadastralParcels.Count > 0 Then
                                            For i As Integer = 0 To listCadastralParcels.Count - 1
                                                Dim cadastralPolygon As Polygon = listCadastralParcels.Item(i).netPolygon
                                                Dim boolIntersectPolygon As Boolean = netTopo.DoPolygonsIntersect(userPolygon, cadastralPolygon)
                                                If boolIntersectPolygon = True Then
                                                    Dim geom As Geometry = netTopo.GetIntersection(userPolygon, cadastralPolygon)
                                                    'если полигон имеет площадь более 1 кв.м, тогда соддаем участок рекультивации
                                                    If geom.Area >= 1 Then
                                                        Dim arrayData As String(,) = listCadastralParcels.Item(i).arrayData
                                                        Dim Category As String = MathFunction.FuncFindValueToArray2d(arrayData, "Category", 0, 1)
                                                        Dim cad_number As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cad_number", 0, 1)
                                                        Dim Utilization As String = MathFunction.FuncFindValueToArray2d(arrayData, "Utilization", 0, 1)
                                                        Dim UtilizationDoc As String = MathFunction.FuncFindValueToArray2d(arrayData, "UtilizationByDoc", 0, 1)
                                                        Dim Address As String = MathFunction.FuncFindValueToArray2d(arrayData, "Adress", 0, 1)
                                                        Dim Patronymic As String = MathFunction.FuncFindValueToArray2d(arrayData, "Patronymic", 0, 1)
                                                        For k2 As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                                            Dim NameField As String = ArrayDataObj(1, k2).Trim
                                                            If NameField Like "ConditionalNumber" Then
                                                                ArrayDataObj(2, k2) = ""
                                                            ElseIf NameField Like "Patronymic" Then
                                                                ArrayDataObj(2, k2) = Patronymic
                                                            ElseIf NameField Like "Cad_number" Then
                                                                ArrayDataObj(2, k2) = cad_number
                                                            ElseIf NameField Like "Category" Then
                                                                ArrayDataObj(2, k2) = Category
                                                            ElseIf NameField Like "Utilization" Then
                                                                If Utilization.Trim.Length > 0 Then
                                                                    ArrayDataObj(2, k2) = Utilization
                                                                Else
                                                                    ArrayDataObj(2, k2) = UtilizationDoc
                                                                End If
                                                            ElseIf NameField Like "Area" Then
                                                                ArrayDataObj(2, k2) = Math.Round(userPoly.Area)
                                                            ElseIf NameField Like "Address" Then
                                                                ArrayDataObj(2, k2) = Address
                                                            ElseIf NameField Like "GLOBALID" Then
                                                                ArrayDataObj(2, k2) = ""
                                                            ElseIf NameField Like "LOCALID" Then
                                                                ArrayDataObj(2, k2) = rng.Next.ToString
                                                            End If
                                                        Next k2
                                                        Dim arrayPoly As Polyline() = AcadFunction.FuncDrawPolylineByPolygon(geom)
                                                        If arrayPoly.Length > 0 Then
                                                            For j As Integer = 0 To arrayPoly.Length - 1
                                                                Dim poly As Polyline = arrayPoly(j)
                                                                Dim numberP As String = numbParcel
                                                                If arrayPoly.Length > 1 Then
                                                                    numberP = numberP & "(" & j + 1 & ")"
                                                                End If
                                                                For k2 As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                                                    Dim NameField As String = ArrayDataObj(1, k2).Trim
                                                                    If NameField Like "ConditionalNumber" Then
                                                                        ArrayDataObj(2, k2) = numberP
                                                                    End If
                                                                Next k2
                                                                'присваиваем таблицу PS
                                                                If IsNothing(idTablePs) = False Then
                                                                    If idTablePs.IsNull = False Then
                                                                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(poly, idTablePs)
                                                                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(poly, ArrayDataObj)
                                                                    End If
                                                                End If
                                                                'присваиваем харакиеричтики линии
                                                                'заносим характеристики
                                                                poly.Layer = LayerLine
                                                                poly.ColorIndex = ColorUserLine
                                                                poly.LineWeight = WidthUserLine
                                                                poly.Linetype = TypeUserLine
                                                                poly.ConstantWidth = CondUserLile
                                                            Next j
                                                        End If
                                                        boolFindParcels = True
                                                        numbParcel += 1
                                                    End If
                                                End If
                                            Next i
                                            'If listCadastralKvartal.Count > 0 Then
                                            '    For i As Integer = 0 To listCadastralKvartal.Count - 1
                                            '        Dim cadastralPolygon As Polygon = listCadastralKvartal.Item(i).netPolygon
                                            '        Dim boolIntersectPolygon As Boolean = netTopo.DoPolygonsIntersect(userPolygon, cadastralPolygon)
                                            '        If boolIntersectPolygon = True Then
                                            '            Dim geom As Geometry = netTopo.GetIntersection(userPolygon, cadastralPolygon)
                                            '            'усли полигон большей частью находится в кадастровом участке
                                            '            If geom.Area >= userPoly.Area / 2 Then
                                            '                Dim arrayData As String(,) = listCadastralKvartal.Item(i).arrayData
                                            '                Dim cad_number As String = MathFunction.FuncFindValueToArray2d(arrayData, "Cadastral_number", 0, 1)
                                            '                For k2 As Integer = 0 To ArrayDataObj.GetUpperBound(1)
                                            '                    Dim NameField As String = ArrayDataObj(1, k2).Trim
                                            '                    If NameField Like "ConditionalNumber" Then
                                            '                        ArrayDataObj(2, k2) = ""
                                            '                    ElseIf NameField Like "Cad_number" Then
                                            '                        ArrayDataObj(2, k2) = cad_number
                                            '                    ElseIf NameField Like "Area" Then
                                            '                        ArrayDataObj(2, k2) = Math.Round(userPoly.Area)
                                            '                    ElseIf NameField Like "KeyParameter" Then
                                            '                        ArrayDataObj(2, k2) = valTh
                                            '                    ElseIf NameField Like "GLOBALID" Then
                                            '                        ArrayDataObj(2, k2) = ""
                                            '                    ElseIf NameField Like "LOCALID" Then
                                            '                        ArrayDataObj(2, k2) = rng.Next.ToString
                                            '                    ElseIf NameField Like "Category" Then
                                            '                        ArrayDataObj(2, k2) = "Категория не установлена"
                                            '                    ElseIf NameField Like "Utilization" Then
                                            '                        ArrayDataObj(2, k2) = "-"
                                            '                    ElseIf NameField Like "Patronymic" Then
                                            '                        ArrayDataObj(2, k2) = "Не разграниченная государственная собственность"
                                            '                    End If
                                            '                Next k2
                                            '                'присваиваем таблицу PS
                                            '                If IsNothing(idTablePs) = False Then
                                            '                    If idTablePs.IsNull = False Then
                                            '                        Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(userPoly, idTablePs)
                                            '                        Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(userPoly, ArrayDataObj)
                                            '                    End If
                                            '                End If
                                            '                'присваиваем харакиеричтики линии
                                            '                'заносим характеристики
                                            '                userPoly.Layer = LayerLine
                                            '                userPoly.ColorIndex = ColorUserLine
                                            '                userPoly.LineWeight = WidthUserLine
                                            '                userPoly.Linetype = TypeUserLine
                                            '                userPoly.ConstantWidth = CondUserLile
                                            '                boolFindParcels = True
                                            '                Exit For
                                            '            End If
                                            '        End If
                                            '    Next i
                                            'End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                Next
            End If
            tr.Commit()
        End Using
    End Sub

    '2. назначить толщину растительного грунта участкам
    <CommandMethod("LandReclamationCreateThLayer")>
    Public Sub LandReclamationCreateThLayer()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        ArrayTablesXML(0, 0) = "LAND_RECLAMATION"
        ArrayTablesXML(1, 0) = "Рекультивация земель"

        Dim IDArrayPLine As String(,) = Nothing 'массив для хранения глобальных ID участков и их Hangle (0-Handle , 1-ID)
        Dim CountPline As Integer = 0 'счетчик участков
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                '=========================================================================================================
                'выбираем участки
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
                                    Dim boolTbl As Boolean = FuncPropertySet.FuncFindPSetTable(acEnt, ArrayTablesXML(0, 0))
                                    If boolTbl = True Then
                                        ReDim Preserve ArrayAllPoligon(1, countArrayAllPoligon)
                                        ArrayAllPoligon(0, countArrayAllPoligon) = acEnt.Handle.ToString
                                        ArrayAllPoligon(1, countArrayAllPoligon) = ArrayTablesXML(0, 0)
                                        countArrayAllPoligon += 1
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
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'спрашиваем что назначаем плохой или хороший грунт
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Грунт пригоден для целей рекультивации?"
                pIntOpts.Keywords.Add("Да")
                pIntOpts.Keywords.Add("Нет")
                pIntOpts.Keywords.Add("Игнорировать")
                pIntOpts.Keywords.Default = "Да"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'вводим толщину
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите толщину горизонта в см: ")
                Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
                Dim valTh As Double = Math.Abs(pStrRes.Value / 100)
                Dim ArrayData As String(,) = Nothing
                If pIntRes.StringResult Like "Да" Then
                    ReDim Preserve ArrayData(2, 0)
                    ArrayData(0, 0) = "LAND_RECLAMATION"
                    ArrayData(1, 0) = "ThLayer"
                    ArrayData(2, 0) = valTh
                ElseIf pIntRes.StringResult Like "Нет" Then
                    ReDim Preserve ArrayData(2, 0)
                    ArrayData(0, 0) = "LAND_RECLAMATION"
                    ArrayData(1, 0) = "BadThLayer"
                    ArrayData(2, 0) = valTh
                ElseIf pIntRes.StringResult Like "Игнорировать" Then
                    ReDim Preserve ArrayData(2, 1)
                    ArrayData(0, 0) = "LAND_RECLAMATION"
                    ArrayData(1, 0) = "BadThLayer"
                    ArrayData(2, 0) = valTh
                    ArrayData(0, 1) = "LAND_RECLAMATION"
                    ArrayData(1, 1) = "ThLayer"
                    ArrayData(2, 1) = valTh
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем обозначения участков
                For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                    Dim NameOD As String = ArrayAllPoligon(1, i)
                    Dim PlineHandle As String = ArrayAllPoligon(0, i)
                    Dim ln As Long = Convert.ToInt64(PlineHandle, 16)
                    Dim hn As Handle = New Handle(ln)
                    Dim id As ObjectId = db.GetObjectId(False, hn, 0)
                    Dim PlineObj As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(id, OpenMode.ForWrite)
                    'записываем данные в полилинию
                    '=======================================================================================
                    'записываем в OD
                    If IsArray(ArrayData) = True Then
                        'записываем в PropertySet
                        Dim boolWr2 As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(PlineObj, ArrayData)
                    End If
                Next i
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub

    '3. Экспликация земельных участков
    <CommandMethod("LandReclamationListParcels")>
    Public Sub LandReclamationListParcels()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayProjectParcels As String(,) = Nothing 'массив для хранения всех выбранных проектируемых участков (handle, nameOD)
        Dim countProjectParcels As Integer = 0

        Dim ArrayCadastralParcels As String(,) = Nothing 'массив для хранения всех выбранных существующих кадастровых участков (handle, nameOD)
        Dim countCadastralParcels As Integer = 0

        Dim ArrayPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        'только участки с временным отводом
        ArrayTablesXML(0, 0) = "LAND_RECLAMATION"
        ArrayTablesXML(1, 0) = "Рекультивация земель"
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам
        Dim PutchTemplate As String = "k:\Civil3DToolsUtility\TemplateTables\LandReclamation\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\LandReclamation\") Then
            PutchTemplate = parentPatch2 & "\TemplateTables\LandReclamation\"
        End If
        If IO.Directory.Exists(PutchTemplate) = False Then
            MsgBox("Шаблон таблицы не найден!!!")
            Exit Sub
        End If

        Dim IDArrayPLine As String(,) = Nothing 'массив для хранения глобальных ID участков и их Hangle (0-Handle , 1-ID)
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayOstrov As String() = Nothing ' массив для хранения островков на выбранном контуре
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim boolAlign As Boolean = False
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                Dim pKeyOpts2 As PromptKeywordOptions = New PromptKeywordOptions("")
                pKeyOpts2.Message = vbLf & "Сделать сортировку участков по номеру или по пикетажу?: "
                pKeyOpts2.Keywords.Add("Номеру")
                pKeyOpts2.Keywords.Add("Пикету")
                pKeyOpts2.Keywords.Default = "Номеру"
                pKeyOpts2.AllowNone = False
                Dim pKeyRes2 As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts2)
                Dim opt As PromptSelectionOptions = New PromptSelectionOptions()
                opt.MessageForAdding = vbLf & "Выберите проектные участки: "
                'выбираем участки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(opt)
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
                                    Dim boolTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)

                                    If IsArray(ArrayNameTableOD) = True Then
                                        For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                            ReDim Preserve ArrayProjectParcels(1, countProjectParcels)
                                            ArrayProjectParcels(0, countProjectParcels) = acEnt.Handle.ToString
                                            ArrayProjectParcels(1, countProjectParcels) = ArrayNameTableOD(i)
                                            countProjectParcels += 1
                                        Next i
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayProjectParcels) = False Then
                    MsgBox("Не удалось найти участки!!! Макро прервано.")
                    Exit Sub
                End If
                '========================================================================================================
                'выбираем трассу
                Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
                If ColAlign.Count = 1 Then
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                    boolAlign = True
                Else
                    '=============================================================
                    'выбираем трассу
                    Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу: ")
                    If ppr1.Status = PromptStatus.OK Then
                        Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                        If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                            Align = ent1
                            boolAlign = True
                        End If
                    End If
                End If

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем обозначения участков
                For i As Integer = 0 To ArrayProjectParcels.GetUpperBound(1)
                    Dim tableProjectPS As String = ArrayProjectParcels(1, i)
                    Dim PlineHandle As String = ArrayProjectParcels(0, i)
                    Dim idProjectParcels As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                    If IsNothing(idProjectParcels) = True Then Continue For
                    If idProjectParcels.IsNull = True Then Continue For
                    Dim projectParcel As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(idProjectParcels, OpenMode.ForRead)
                    'читаем расширенные свойства
                    Dim arrayData As String(,) = Nothing
                    Dim NameTBLPS As String = ""
                    Dim GlobalID As String = ""
                    Dim LocalID As String = ""
                    Dim ConditionParcel As String = "1"
                    Dim CadastrtalNumber As String = "00:00:0000000:0000"
                    Dim Purpose As String = "Для переустройства инженерных сетей"
                    Dim Patronymic As String = ""
                    Dim Adress As String = "-"
                    Dim Category As String = "Не установлена"
                    Dim Utilization As String = "-"
                    Dim note As String = "-"
                    Dim pkStr As String = "-"
                    Dim direction As String = ""
                    Dim AreaParcel As Double = 0
                    Dim countReadFlag As Integer = 0
                    Dim countCounters As Integer = 1
                    'читаем свойства
                    Dim boolXRecord As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(projectParcel, arrayData, tableProjectPS)
                    If IsArray(arrayData) = True Then
                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim NField As String = arrayData(1, j)
                            Dim Zn As String = arrayData(2, j)
                            If NField Like "GLOBALID" Then
                                GlobalID = Zn
                            ElseIf NField Like "LOCALID" Then
                                LocalID = Zn
                                NameTBLPS = arrayData(0, j)
                            ElseIf NField Like "ConditionalNumber" Then
                                ConditionParcel = Zn
                            ElseIf NField Like "Cad_number" Then
                                CadastrtalNumber = Zn
                            ElseIf NField Like "Purpose" Then
                                If Zn.Trim.Length > 0 Then
                                    Purpose = Zn
                                End If
                            ElseIf NField Like "Address" Then
                                If Zn.Trim.Length > 0 Then
                                    Adress = Zn
                                End If
                            ElseIf NField Like "Category" Then
                                If Zn.Trim.Length > 0 Then
                                    Category = Zn
                                End If
                            ElseIf NField Like "Utilization" Then
                                If Zn.Trim.Length > 0 Then
                                    Utilization = Zn
                                End If
                            ElseIf NField Like "Patronymic" Then
                                If Zn.Trim.Length > 0 Then
                                    Patronymic = Zn
                                End If
                            ElseIf NField Like "Direction" Then
                                If Zn.Trim.Length > 0 Then
                                    direction = Zn
                                End If
                            ElseIf NField Like "NOTE" Then
                                If Zn.Trim.Length > 0 Then
                                    note = Zn
                                End If
                            End If
                        Next j
                    End If
                    '==========================================================================================
                    'проверяем участок на многоконтурность
                    If GlobalID.Trim.Length > 1 Then
                        'ищем глобальный ид в массиве дабы исключить дубли
                        If IsArray(ArrayPoligon) = True Then
                            For j As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                                If ArrayPoligon(16, j) Like GlobalID Then
                                    GoTo line1
                                End If
                            Next j
                        End If
                    End If
                    '========================================================================================================
                    'определяем внутреннюю точку полигона
                    Dim posProjectParcel As Point3d = New Point3d(-1, -1, -1)
                    Dim pkProjectParcel As Double = -1
                    Dim boolInsPoint As Boolean = False
                    Dim acMleader As MLeader = FuncCadastr.FuncFindMLeaderByLocalId(LocalID)
                    If IsNothing(acMleader) = False Then
                        If acMleader.Id.IsNull = False Then
                            posProjectParcel = acMleader.GetFirstVertex(0)
                            boolInsPoint = MathFunction.FuncCrossPointPolyline(projectParcel, posProjectParcel.Convert2d(New Plane))
                        End If
                    End If
                    If boolInsPoint = False Then
                        posProjectParcel = MathFunction.FuncFindCrossCenterPointPolyline(projectParcel)
                        boolInsPoint = MathFunction.FuncCrossPointPolyline(projectParcel, posProjectParcel.Convert2d(New Plane))
                    End If
                    '==============================================================================================================
                    'определяем пикет
                    If IsNothing(Align) = False Then
                        Dim off As Double = -1
                        'определяем пикет относительно трассы
                        If boolInsPoint = True Then
                            Try
                                Align.StationOffset(posProjectParcel.X, posProjectParcel.Y, pkProjectParcel, off)
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                        End If
                        'если пикет не удалос определить, то делаем более сложный поиск пикетажа о вершинам
                        If pkProjectParcel = -1 And off = -1 Then
                            For j As Integer = 0 To projectParcel.NumberOfVertices - 1
                                Try
                                    Align.StationOffset(projectParcel.GetPoint2dAt(j).X, projectParcel.GetPoint2dAt(j).Y, pkProjectParcel, off)
                                    Exit For
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            Next j
                        End If
                        If pkProjectParcel <> -1 Then
                            Try
                                pkStr = Align.GetStationStringWithEquations(pkProjectParcel)
                            Catch ex As System.ArgumentException
                            End Try
                        End If
                    End If
                    '=========================================================================================================
                    'ищем контура проектного участка
                    If LocalID.Trim.Length > 0 Then
                        Dim ArrayCounterTemp As String(,) = Nothing
                        If GlobalID.Trim.Length > 2 Then
                            Dim boolKounter As Boolean = FuncCadastr.FuncFindCountersParcel(GlobalID, ArrayCounterTemp, NameTBLPS)
                        End If
                        If IsArray(ArrayCounterTemp) = True Then
                            For j As Integer = 0 To ArrayCounterTemp.GetUpperBound(1)
                                Dim HandlePLine As String = ArrayCounterTemp(0, j) 'нангл контура
                                Dim idLoc As String = ArrayCounterTemp(1, j) 'локальный ид
                                'получаем доступ к полилинии для извлечения координат
                                Dim idCounter As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                If IsNothing(idCounter) = False Then
                                    If idCounter.IsNull = False Then
                                        Dim counterProjectParcel As Polyline = tr.GetObject(idCounter, OpenMode.ForRead)
                                        AreaParcel = AreaParcel + counterProjectParcel.Area
                                        'ищем островки
                                        Dim ArrayOstrov1 As String(,) = Nothing
                                        Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(idLoc, ArrayOstrov1)
                                        If IsArray(ArrayOstrov1) = True Then
                                            For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                                'получаем доступ к полилинии для извлечения координат
                                                HandlePLine = ArrayOstrov1(1, k)
                                                Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine)
                                                If IsNothing(idOstrov) = False Then
                                                    If idOstrov.IsNull = False Then
                                                        Dim ostrovProjectParcel As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                                        AreaParcel = AreaParcel - ostrovProjectParcel.Area
                                                    End If
                                                End If
                                            Next k
                                        End If
                                    End If
                                End If
                            Next j
                            If ArrayCounterTemp.GetUpperBound(1) > 0 Then
                                countCounters = ArrayCounterTemp.GetUpperBound(1) + 1
                            End If
                        Else
                            AreaParcel = projectParcel.Area
                            'ищем островки
                            Dim ArrayOstrov1 As String(,) = Nothing
                            Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov1)
                            If IsArray(ArrayOstrov1) = True Then
                                For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                                    'получаем доступ к полилинии для извлечения координат
                                    Dim HandlePLine1 As String = ArrayOstrov1(1, k)
                                    Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine1)
                                    If IsNothing(idOstrov) = False Then
                                        If idOstrov.IsNull = False Then
                                            Dim ostrovProjectParcel As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                            AreaParcel = AreaParcel - ostrovProjectParcel.Area
                                        End If
                                    End If
                                Next k
                            End If
                        End If
                        '==============================================================================================
                        'отделяем номер от обозначения участка
                        Dim numbParcel As Integer = 1
                        If pKeyRes2.StringResult Like "Номеру" Then
                            numbParcel = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionParcel)
                        End If
                        'обозначение проектного участка
                        Dim oboznZU As String = ConditionParcel
                        If countCounters > 1 Then
                            oboznZU = FuncCadastr.FuncFindFullConditionalNumberCadastralObject(ConditionParcel) & ", состоящий из " & countCounters & " контуров."
                        End If
                        '==============================================================================================
                        'записываем результат в массив
                        ReDim Preserve ArrayPoligon(15, countArrayPoligon)
                        If pKeyRes2.StringResult Like "Номеру" Then
                            ArrayPoligon(0, countArrayPoligon) = numbParcel  'номер участка (число)
                        Else
                            ArrayPoligon(0, countArrayPoligon) = pkProjectParcel  'пикетаж
                        End If
                        ArrayPoligon(1, countArrayPoligon) = oboznZU 'полное обозначение участка
                        ArrayPoligon(2, countArrayPoligon) = pkStr
                        ArrayPoligon(3, countArrayPoligon) = CadastrtalNumber 'кадастровый номер
                        ArrayPoligon(4, countArrayPoligon) = Adress 'адрес
                        ArrayPoligon(5, countArrayPoligon) = Category 'категория земель
                        ArrayPoligon(6, countArrayPoligon) = Utilization 'существующее разрешенное использование
                        ArrayPoligon(7, countArrayPoligon) = Purpose 'планируемое разрешенное использование
                        Dim totalArea As Double = Math.Round(AreaParcel, 0) 'площадь участка
                        If totalArea = 0 Then
                            totalArea = Math.Round(projectParcel.Area, 2) 'площадь участка
                        End If
                        ArrayPoligon(8, countArrayPoligon) = totalArea
                        ArrayPoligon(9, countArrayPoligon) = direction
                        ArrayPoligon(10, countArrayPoligon) = Patronymic 'фамилия правообладателя
                        ArrayPoligon(12, countArrayPoligon) = note 'примечание
                        ArrayPoligon(13, countArrayPoligon) = tableProjectPS 'тип участка
                        ArrayPoligon(14, countArrayPoligon) = GlobalID 'тип участка
                        countArrayPoligon += 1
                    End If
line1:
                Next i
                '==================================================================================================
                If IsArray(ArrayPoligon) = False Then
                    MsgBox("Массив с участками пуст!!!")
                    Exit Sub
                End If
                '===================================================================================================
                If ArrayPoligon.GetUpperBound(1) > 0 Then
                    'сортируем участки
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(ArrayPoligon, 0)
                End If
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записываем отчет
                If IsArray(ArrayPoligon) = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'Word документ
                    Dim WApp As Word.Application = Nothing
                    Dim WDoc As Word.Document = Nothing
                    Dim rng As Random = New Random
                    Dim boolWDVisibl As Boolean = False
                    Try
                        WApp = CreateObject("Word.Application")
                        WDoc = WApp.Documents.Open(PutchTemplate & "Сведения о площади месторасположении и характере нарушения земель.docx")
                        Dim Numbrng As Integer = rng.Next
                        Dim TempDir As String = Path.GetTempPath()
                        WDoc.SaveAs2(TempDir & "Сведения о площади месторасположении и характере нарушения земель_" & Numbrng & ".docx")
                    Catch ex As Exception
                        MsgBox("Файл-шаблон не найден.")
                        WDoc = Nothing
                        WApp.Quit()
                        WApp = Nothing
                        Exit Sub
                    End Try
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    WApp.Visible = True
                    For i As Integer = 0 To ArrayPoligon.GetUpperBound(1)
                        For Each tabl As Word.Table In WDoc.Tables
                            Dim countTableword As Integer = tabl.Rows.Count
                            tabl.Rows.Add()
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'записываем номер участка
                            tabl.Cell(countTableword, 1).Range.Text = ArrayPoligon(1, i) 'номер участка
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'записываем пикет
                            Dim pk As String = ArrayPoligon(2, i)
                            If IsNothing(pk) = True Then pk = "-"
                            If pk.Trim.Length = 0 Then pk = "-"
                            tabl.Cell(countTableword, 2).Range.Text = pk  'пикет
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'кадастровый номер
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim KN As String = ArrayPoligon(3, i)
                            If IsNothing(KN) = True Then KN = "-"
                            If KN.Trim.Length = 0 Then KN = "-"
                            tabl.Cell(countTableword, 3).Range.Text = KN 'кадастровый номер
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'адрес
                            Dim Adress As String = ArrayPoligon(4, i)
                            If IsNothing(Adress) = True Then Adress = "-"
                            If Adress.Trim.Length = 0 Then Adress = "-"
                            tabl.Cell(countTableword, 4).Range.Text = Adress 'адрес
                            'площадь
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim AreaP As Double = Val(ArrayPoligon(8, i))
                            tabl.Cell(countTableword, 5).Range.Text = AreaP 'площадь
                            'цель предоставления
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim Udp As String = ArrayPoligon(7, i)
                            If IsNothing(Udp) = True Then Udp = "-"
                            If Udp.Trim.Length = 0 Then Udp = "-"
                            tabl.Cell(countTableword, 6).Range.Text = Udp 'цель предоставления
                            'категория земель
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim Kategory As String = ArrayPoligon(5, i)
                            If Kategory.Trim.Length > 0 Then
                                tabl.Cell(countTableword, 7).Range.Text = Kategory 'обозначение
                            Else
                                tabl.Cell(countTableword, 7).Range.Text = "-" 'обозначение
                            End If
                            'разрешенное использование
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            Dim Utiliza As String = ArrayPoligon(6, i)
                            If Utiliza.Trim.Length > 0 Then
                                tabl.Cell(countTableword, 8).Range.Text = Utiliza 'обозначение
                            Else
                                tabl.Cell(countTableword, 8).Range.Text = "-" 'обозначение
                            End If
                            'примечание
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'направление проведения рекультивации
                            Dim direction As String = ArrayPoligon(9, i)
                            If IsNothing(direction) = True Then direction = "-"
                            If direction.Trim.Length = 0 Then direction = "-"
                            tabl.Cell(countTableword, 9).Range.Text = direction  'пикет
                            'Dim Rg As String = ArrayPoligon(10, i)
                            'Dim Patron As String = ArrayPoligon(10, i)
                            'Dim enc As String = ArrayPoligon(11, i)
                            'Dim note As String = ArrayPoligon(12, i)
                            'If Rg.IndexOf("еразгр") > -1 Then
                            '    tabl.Cell(countTableword, 9).Range.Text = "Неразграниченная государственная собственность"
                            'ElseIf Rg.IndexOf("РФ") > -1 Then
                            '    Dim str As String = "Собственность Российской Федерации"
                            '    If enc.Trim.Length > 0 Then
                            '        str = str & vbLf & enc
                            '    End If
                            '    If Patron.Trim.Length > 0 Then
                            '        str = str & vbLf & Patron
                            '    End If
                            '    If note.Trim.Length > 0 Then
                            '        str = str & vbLf & note
                            '    End If
                            '    tabl.Cell(countTableword, 9).Range.Text = str
                            'ElseIf Rg.IndexOf("униципальн") > -1 Then
                            '    Dim str As String = "Муниципальная собственность"
                            '    If enc.Trim.Length > 0 Then
                            '        str = str & vbLf & enc
                            '    End If
                            '    If Patron.Trim.Length > 0 Then
                            '        str = str & vbLf & Patron
                            '    End If
                            '    If note.Trim.Length > 0 Then
                            '        str = str & vbLf & note
                            '    End If
                            '    tabl.Cell(countTableword, 9).Range.Text = str
                            'Else
                            '    Dim StrGg As String = ""
                            '    If Rg.Trim.Length > 0 Then
                            '        StrGg = Rg
                            '    End If
                            '    If Patron.Trim.Length > 0 And StrGg.Trim.Length > 0 Then
                            '        StrGg = StrGg & vbLf & Patron
                            '    ElseIf Patron.Trim.Length > 0 And StrGg.Trim.Length = 0 Then
                            '        StrGg = Patron
                            '    End If
                            '    If note.Trim.Length > 0 Then
                            '        StrGg = StrGg & vbLf & note
                            '    End If
                            '    tabl.Cell(countTableword, 9).Range.Text = StrGg
                            'End If
                            Exit For
                        Next
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                End If
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub

    '4. Ведомость объёмов работ для ПЗ
    <CommandMethod("LandReclamationVolumeParcels")>
    Public Sub LandReclamationVolumeParcels()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        ArrayTablesXML(0, 0) = "LAND_RECLAMATION"
        ArrayTablesXML(1, 0) = "Рекультивация земель"
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '1. путь к шаблонам
        Dim PutchTemplate As String = "k:\Civil3DToolsUtility\TemplateTables\LandReclamation\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\LandReclamation\") Then
            PutchTemplate = parentPatch2 & "\TemplateTables\LandReclamation\"
        End If
        If IO.Directory.Exists(PutchTemplate) = False Then
            MsgBox("Шаблон таблицы не найден!!!")
            Exit Sub
        End If
        'технологическая схема
        Dim templateShemeXML As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\Semantic\") Then
            templateShemeXML = parentPatch2 & "\FileResources\Sample\Semantic\"
            templateShemeXML = templateShemeXML & "LandReclamation.xml"
            If File.Exists(templateShemeXML) = False Then
                MsgBox("Шаблон оформления чертежей не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim dictReclamationParcel As Dictionary(Of Integer, LandReclamation()) = New Dictionary(Of Integer, LandReclamation())
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                '=========================================================================================================
                'выбираем участки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                Dim ArrayNameTableOD As String() = Nothing
                                Dim boolTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                If IsArray(ArrayNameTableOD) = True Then
                                    For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                        ReDim Preserve ArrayAllPoligon(1, countArrayAllPoligon)
                                        ArrayAllPoligon(0, countArrayAllPoligon) = acEnt.Handle.ToString
                                        ArrayAllPoligon(1, countArrayAllPoligon) = ArrayNameTableOD(i)
                                        countArrayAllPoligon += 1
                                    Next i
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось найти участки!!!")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем характеристики участков
                For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                    Dim NameOD As String = ArrayAllPoligon(1, i)
                    Dim PlineHandle As String = ArrayAllPoligon(0, i)
                    Dim idProjectParcel As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                    If IsNothing(idProjectParcel) = True Then Continue For
                    If idProjectParcel.IsNull = True Then Continue For
                    Dim projectParcel As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(idProjectParcel, OpenMode.ForRead)
                    'ищем глобальный и локальный id
                    Dim arrayData As String(,) = {}
                    Dim GlobalID As String = ""
                    Dim LocalID As String = ""
                    Dim ConditionParcel As String = ""
                    Dim biologicalSheme As Integer = 0
                    Dim BadThLayer As Double = 0
                    Dim ThLayer As Double = 0
                    Dim boolXRecord As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(projectParcel, arrayData, ArrayTablesXML)
                    If IsArray(arrayData) = True Then
                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim NField As String = arrayData(1, j)
                            Dim Zn As String = arrayData(2, j)
                            If Zn.Trim.Length > 0 Then
                                If NField Like "GLOBALID" Then
                                    GlobalID = Zn
                                ElseIf NField Like "LOCALID" Then
                                    LocalID = Zn
                                ElseIf NField Like "ConditionalNumber" Then
                                    ConditionParcel = Zn
                                ElseIf NField Like "BiologicalScheme" Then
                                    biologicalSheme = Val(Zn)
                                ElseIf NField Like "BadThLayer" Then
                                    BadThLayer = Val(Zn)
                                ElseIf NField Like "ThLayer" Then
                                    ThLayer = Val(Zn)
                                End If
                            End If
                        Next j
                    End If
                    '===============================================================================
                    'отделяем номер участка от его префикса
                    Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionParcel)
                    'определяем площадь
                    Dim AreaParcel As Double = projectParcel.Area
                    'ищем островки
                    Dim ArrayOstrov As String(,) = Nothing
                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov)
                    If IsArray(ArrayOstrov) = True Then
                        For k As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                            'получаем доступ к полилинии для извлечения координат
                            Dim HandlePLine1 As String = ArrayOstrov(1, k)
                            Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine1)
                            If IsNothing(idOstrov) = False Then
                                If idOstrov.IsNull = False Then
                                    Dim ostrovPline As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                    AreaParcel = AreaParcel - ostrovPline.Area
                                End If
                            End If
                        Next k
                    End If
                    'находим технологическую схему проведения биологической рекультивации
                    '==============================================================================================
                    'считываем схему биологической рекультивации
                    Dim numberYear As Integer = ParseYearReclamationXml(templateShemeXML, "biologicalStage", biologicalSheme)
                    If numberYear = 0 Then numberYear = 1
                    Dim rectamationParcel As LandReclamation = New LandReclamation(AreaParcel, ThLayer, BadThLayer)
                    For numYaear As Integer = 1 To numberYear
                        Dim dictBiologicalSheme As List(Of Dictionary(Of String, String)) = ParseReclamationXml(templateShemeXML, "biologicalStage", biologicalSheme, numYaear)
                        If dictBiologicalSheme.Count > 0 Then
                            For j As Integer = 0 To dictBiologicalSheme.Count - 1
                                Dim dictBiological As Dictionary(Of String, String) = dictBiologicalSheme.Item(j)
                                '======================================================================================================
                                'мелиоранты
                                If dictBiological.ContainsKey("Ameliorant") = True Then
                                    Dim tempAmeliorant As String = dictBiological.Item("Ameliorant")
                                    Dim dictAmeliorant As New Dictionary(Of String, Double)
                                    If tempAmeliorant.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempAmeliorant.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictAmeliorant.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictAmeliorant.Count >= k + 1 Then
                                                        Dim name = dictAmeliorant.ElementAt(k).Key
                                                        dictAmeliorant.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictAmeliorant.Count > 0 Then
                                        rectamationParcel.MineralFertilizers.Add(dictAmeliorant)
                                    End If
                                End If
                                '======================================================================================================
                                'минеральные удобрения
                                If dictBiological.ContainsKey("MineralFertilizers") = True Then
                                    Dim tempMineralFertilizers As String = dictBiological.Item("MineralFertilizers")
                                    Dim dictMineralFertilizers As New Dictionary(Of String, Double)
                                    If tempMineralFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempMineralFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictMineralFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictMineralFertilizers.Count >= k + 1 Then
                                                        Dim name = dictMineralFertilizers.ElementAt(k).Key
                                                        dictMineralFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictMineralFertilizers.Count > 0 Then
                                        rectamationParcel.MineralFertilizers.Add(dictMineralFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'органические удобрения
                                If dictBiological.ContainsKey("OrganicalFertilizers") = True Then
                                    Dim tempOrganicalFertilizers As String = dictBiological.Item("OrganicalFertilizers")
                                    Dim dictOrganicalFertilizers As New Dictionary(Of String, Double)
                                    If tempOrganicalFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempOrganicalFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictOrganicalFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictOrganicalFertilizers.Count >= k + 1 Then
                                                        Dim name = dictOrganicalFertilizers.ElementAt(k).Key
                                                        dictOrganicalFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictOrganicalFertilizers.Count > 0 Then
                                        rectamationParcel.OrganicalFertilizers.Add(dictOrganicalFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'сидераты
                                If dictBiological.ContainsKey("GreenSiderat") = True Then
                                    Dim tempGreenSiderat As String = dictBiological.Item("GreenSiderat")
                                    Dim dictGreenSiderat As New Dictionary(Of String, Double)
                                    If tempGreenSiderat.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempGreenSiderat.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictGreenSiderat.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictGreenSiderat.Count >= k + 1 Then
                                                        Dim name = dictGreenSiderat.ElementAt(k).Key
                                                        dictGreenSiderat.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictGreenSiderat.Count > 0 Then
                                        rectamationParcel.GreenSiderats.Add(dictGreenSiderat)
                                    End If
                                End If
                            Next j
                        End If
                    Next numYaear
                    If dictReclamationParcel.ContainsKey(numbParcel) = False Then
                        dictReclamationParcel.Add(numbParcel, {rectamationParcel})
                    Else
                        Dim arrayReclamationParcel As LandReclamation() = dictReclamationParcel.Item(numbParcel)
                        Dim count As Integer = arrayReclamationParcel.Length
                        ReDim Preserve arrayReclamationParcel(count)
                        arrayReclamationParcel(count) = rectamationParcel
                        dictReclamationParcel.Item(numbParcel) = arrayReclamationParcel
                    End If
                Next i
                tr.Commit()
            End Using
            '==================================================================================================
            If dictReclamationParcel.Count = 0 Then
                MsgBox("Массив с участками пуст!!!")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'сортируем участки сначала по номеру
            If dictReclamationParcel.Count > 1 Then
                dictReclamationParcel = dictReclamationParcel.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '========================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Word документ
            Dim WApp As Word.Application = Nothing
            Dim WDoc As Word.Document = Nothing
            Dim boolWDVisibl As Boolean = False
            Try
                WApp = CreateObject("Word.Application")
                WDoc = WApp.Documents.Open(PutchTemplate & "Ведомость объемов работ.docx")
                Dim rng As Random = New Random
                Dim Numbrng As Integer = rng.Next
                Dim TempDir As String = Path.GetTempPath()
                WDoc.SaveAs2(TempDir & "Ведомость объемов работ_" & Numbrng & ".docx")
                WApp.Visible = True
            Catch ex As Exception
                MsgBox("Файл-шаблон не найден.")
                WDoc = Nothing
                WApp.Quit()
                WApp = Nothing
                Exit Sub
            End Try
            For Each tabl As Word.Table In WDoc.Tables
                If tabl.Title Like "Ведомость объемов работ" Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'начинаем перебирать участки
                    For i As Integer = 0 To dictReclamationParcel.Count - 1
                        Dim arrayPropertiesCounter As LandReclamation() = dictReclamationParcel.ElementAt(i).Value
                        If IsArray(arrayPropertiesCounter) = True Then
                            Dim totalAreaPercel As Double = 0 'общая площадь
                            Dim totalThLayer As Double = 0
                            Dim totalBadThLayer As Double = 0
                            Dim middleThLayer As Double = 0
                            'минеральные удобрения
                            Dim arrayMineralFert As String(,) = {} '0-наименование 1-количество
                            Dim countArrayMineralFert As Integer = 0
                            'органические удобрения
                            Dim arrayOrganicalFert As String(,) = {} '0-наименование 1-количество
                            Dim countArrayOrganicalFert As Integer = 0
                            'сидераты
                            Dim arrayGreenSiderat As String(,) = {} '0-наименование 1-количество
                            Dim countGreenSiderat As Integer = 0
                            'мелиоранты
                            Dim arrayMeliorant As String(,) = {} '0-наименование 1-количество
                            Dim countMeliorant As Integer = 0
                            'читаем характеристи
                            For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                Dim tempArea As Double = landParcelReclamation.Area 'площадь
                                totalAreaPercel += tempArea
                                Dim tempThLayer As Double = landParcelReclamation.thicknessLayer
                                totalThLayer += tempThLayer * tempArea
                                Dim tempBadThLayer As Double = landParcelReclamation.badThicknessLayer
                                totalBadThLayer += tempBadThLayer * tempArea
                                middleThLayer += tempThLayer + tempBadThLayer
                                'делаем подсчет органических удобрений
                                Dim listOrganicalFert As List(Of Dictionary(Of String, Double)) = landParcelReclamation.OrganicalFertilizers 'органические удобрения
                                'список с органическими удобрениями не пустой, продолжаем
                                If listOrganicalFert.Count > 0 Then
                                    'читаем все удобрения и их нормы
                                    For k As Integer = 0 To listOrganicalFert.Count - 1
                                        Dim dictValueFert As Dictionary(Of String, Double) = listOrganicalFert(k)
                                        For k1 As Integer = 0 To dictValueFert.Count - 1
                                            Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                            Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                            'делаем расчет количества
                                            Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                            'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                            Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayOrganicalFert, nameFert, 0)
                                            If indexNameFert = -1 Then
                                                ReDim Preserve arrayOrganicalFert(1, countArrayOrganicalFert)
                                                arrayOrganicalFert(0, countArrayOrganicalFert) = nameFert
                                                arrayOrganicalFert(1, countArrayOrganicalFert) = rezultFert
                                                countArrayOrganicalFert += 1
                                            Else
                                                Dim tempRez As Double = Val(arrayOrganicalFert(1, indexNameFert)) + rezultFert
                                                arrayOrganicalFert(1, indexNameFert) = tempRez
                                            End If
                                        Next k1
                                    Next k
                                End If
                                'делаем подсчет минеральных удобрений
                                Dim listMineralFert As List(Of Dictionary(Of String, Double)) = landParcelReclamation.MineralFertilizers 'минеральные удобрения
                                'список с органическими удобрениями не пустой, продолжаем
                                If listMineralFert.Count > 0 Then
                                    'читаем все удобрения и их нормы
                                    For k As Integer = 0 To listMineralFert.Count - 1
                                        Dim dictValueFert As Dictionary(Of String, Double) = listMineralFert(k)
                                        For k1 As Integer = 0 To dictValueFert.Count - 1
                                            Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                            Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                            'делаем расчет количества
                                            Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                            'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                            Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayMineralFert, nameFert, 0)
                                            If indexNameFert = -1 Then
                                                ReDim Preserve arrayMineralFert(1, countArrayMineralFert)
                                                arrayMineralFert(0, countArrayMineralFert) = nameFert
                                                arrayMineralFert(1, countArrayMineralFert) = rezultFert
                                                countArrayMineralFert += 1
                                            Else
                                                Dim tempRez As Double = Val(arrayMineralFert(1, indexNameFert)) + rezultFert
                                                arrayMineralFert(1, indexNameFert) = tempRez
                                            End If
                                        Next k1
                                    Next k
                                End If
                                'делаем подсчет сидератов
                                Dim listGreenSiderat As List(Of Dictionary(Of String, Double)) = landParcelReclamation.GreenSiderats 'сидераты
                                'список с органическими удобрениями не пустой, продолжаем
                                If listGreenSiderat.Count > 0 Then
                                    'читаем все удобрения и их нормы
                                    For k As Integer = 0 To listGreenSiderat.Count - 1
                                        Dim dictValueFert As Dictionary(Of String, Double) = listGreenSiderat(k)
                                        For k1 As Integer = 0 To dictValueFert.Count - 1
                                            Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                            Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                            'делаем расчет количества
                                            Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                            'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                            Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayGreenSiderat, nameFert, 0)
                                            If indexNameFert = -1 Then
                                                ReDim Preserve arrayGreenSiderat(1, countGreenSiderat)
                                                arrayGreenSiderat(0, countGreenSiderat) = nameFert
                                                arrayGreenSiderat(1, countGreenSiderat) = rezultFert
                                                countGreenSiderat += 1
                                            Else
                                                Dim tempRez As Double = Val(arrayGreenSiderat(1, indexNameFert)) + rezultFert
                                                arrayGreenSiderat(1, indexNameFert) = tempRez
                                            End If
                                        Next k1
                                    Next k
                                End If
                                'делаем подсчет мелиорантов
                                Dim listMeliorant As List(Of Dictionary(Of String, Double)) = landParcelReclamation.Ameliorant 'Мелиоранты
                                'список с органическими удобрениями не пустой, продолжаем
                                If listMeliorant.Count > 0 Then
                                    'читаем все удобрения и их нормы
                                    For k As Integer = 0 To listMeliorant.Count - 1
                                        Dim dictValueFert As Dictionary(Of String, Double) = listMeliorant(k)
                                        For k1 As Integer = 0 To dictValueFert.Count - 1
                                            Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                            Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                            'делаем расчет количества
                                            Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                            'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                            Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayMeliorant, nameFert, 0)
                                            If indexNameFert = -1 Then
                                                ReDim Preserve arrayMeliorant(1, countMeliorant)
                                                arrayMeliorant(0, countMeliorant) = nameFert
                                                arrayMeliorant(1, countMeliorant) = rezultFert
                                                countMeliorant += 1
                                            Else
                                                Dim tempRez As Double = Val(arrayMeliorant(1, indexNameFert)) + rezultFert
                                                arrayMeliorant(1, indexNameFert) = tempRez
                                            End If
                                        Next k1
                                    Next k
                                End If
                            Next j
                            'заполняем таблицу
                            tabl.Rows.Add()
                            Dim countTablRow As Integer = tabl.Rows.Count
                            Dim boolFormatRow As Boolean = FuncFormatZN.FormatTableRow(tabl, countTablRow)
                            'номер участка
                            tabl.Cell(countTablRow, 1).Range.Text = dictReclamationParcel.ElementAt(i).Key
                            'мошность растительного грунта
                            Dim middleLayer As Double = Math.Round(middleThLayer / arrayPropertiesCounter.Length, 2)
                            tabl.Cell(countTablRow, 2).Range.Text = middleLayer
                            'объем надвигаемого грунты
                            tabl.Cell(countTablRow, 3).Range.Text = Math.Round(totalThLayer + totalBadThLayer, 2)
                            'объем дополнительно приобретаемого растительного грунта
                            tabl.Cell(countTablRow, 4).Range.Text = Math.Round(totalBadThLayer, 2)
                            'мелиоранты
                            If arrayMeliorant.Length > 0 Then
                                Dim str As String = ""
                                For j As Integer = 0 To arrayMeliorant.GetUpperBound(1)
                                    Dim name As String = arrayMeliorant(0, j)
                                    Dim value As Double = Val(arrayMeliorant(1, j))
                                    str = str & name & "-" & Math.Round(value, 2) & vbLet
                                Next j
                                tabl.Cell(countTablRow, 5).Range.Text = str
                            Else
                                tabl.Cell(countTablRow, 5).Range.Text = "-"
                            End If
                            'органика
                            If arrayOrganicalFert.Length > 0 Then
                                Dim str As String = ""
                                For j As Integer = 0 To arrayOrganicalFert.GetUpperBound(1)
                                    Dim name As String = arrayOrganicalFert(0, j)
                                    Dim value As Double = Val(arrayOrganicalFert(1, j))
                                    str = str & name & "-" & Math.Round(value, 2) & vbLet
                                Next j
                                tabl.Cell(countTablRow, 6).Range.Text = str
                            Else
                                tabl.Cell(countTablRow, 6).Range.Text = "-"
                            End If
                            'менеральные удобрения
                            If arrayMineralFert.Length > 0 Then
                                Dim str As String = ""
                                For j As Integer = 0 To arrayMineralFert.GetUpperBound(1)
                                    Dim name As String = arrayMineralFert(0, j)
                                    Dim value As Double = Val(arrayMineralFert(1, j))
                                    str = str & name & "-" & Math.Round(value, 2) & vbLet
                                Next j
                                tabl.Cell(countTablRow, 7).Range.Text = str
                            Else
                                tabl.Cell(countTablRow, 7).Range.Text = "-"
                            End If
                            'сидераты
                            If arrayGreenSiderat.Length > 0 Then
                                Dim str As String = ""
                                For j As Integer = 0 To arrayGreenSiderat.GetUpperBound(1)
                                    Dim name As String = arrayGreenSiderat(0, j)
                                    Dim value As Double = Val(arrayGreenSiderat(1, j))
                                    str = str & name & "-" & Math.Round(value, 2) & vbLet
                                Next j
                                tabl.Cell(countTablRow, 8).Range.Text = str
                            Else
                                tabl.Cell(countTablRow, 8).Range.Text = "-"
                            End If
                        End If
                    Next i
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                End If
            Next
        Catch
        End Try
    End Sub

    '5. Пояснительная записка
    <CommandMethod("LandReclamationExplanatoryNote")>
    Public Sub LandReclamationExplanatoryNote()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayProjectParcels As String(,) = Nothing 'массив для хранения всех выбранных проектируемых участков (handle, nameOD)
        Dim countProjectParcels As Integer = 0
        Dim ArrayCadastralParcels As String(,) = Nothing 'массив для хранения всех выбранных существующих кадастровых участков (handle, nameOD)
        Dim countCadastralParcels As Integer = 0
        Dim countArrayPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        'только участки с временным отводом
        ArrayTablesXML(0, 0) = "LAND_RECLAMATION"
        ArrayTablesXML(1, 0) = "Рекультивация земель"
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к шаблонам
        Dim PutchTemplate As String = "k:\Civil3DToolsUtility\TemplateTables\LandReclamation\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\LandReclamation\") Then
            PutchTemplate = parentPatch2 & "\TemplateTables\LandReclamation\"
        End If
        If IO.Directory.Exists(PutchTemplate) = False Then
            MsgBox("Шаблон таблицы не найден!!!")
            Exit Sub
        End If
        Dim templateShemeXML As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\Semantic\") Then
            templateShemeXML = parentPatch2 & "\FileResources\Sample\Semantic\"
            templateShemeXML = templateShemeXML & "LandReclamation.xml"
            If File.Exists(templateShemeXML) = False Then
                MsgBox("Шаблон оформления чертежей не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If
        Dim IDArrayPLine As String(,) = Nothing 'массив для хранения глобальных ID участков и их Hangle (0-Handle , 1-ID)
        Dim CountPline As Integer = 0 'счетчик участков
        Dim ArrayOstrov As String() = Nothing ' массив для хранения островков на выбранном контуре
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
        Dim alignColl As ObjectIdCollection = CivilDoc.GetAlignmentIds

        Dim boolAlign As Boolean = False
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                If alignColl.Count = 1 Then
                    Align = tr.GetObject(alignColl.Item(0), OpenMode.ForRead)
                End If
                Dim opt As PromptSelectionOptions = New PromptSelectionOptions()
                opt.MessageForAdding = vbLf & "Выберите участки подлежащие рекультивации: "
                'выбираем участки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(opt)
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
                                    'фмльтр по участкам рекультивации
                                    Dim boolTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                    If IsArray(ArrayNameTableOD) = True Then
                                        For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                            ReDim Preserve ArrayProjectParcels(1, countProjectParcels)
                                            ArrayProjectParcels(0, countProjectParcels) = acEnt.Handle.ToString
                                            ArrayProjectParcels(1, countProjectParcels) = ArrayNameTableOD(i)
                                            countProjectParcels += 1
                                        Next i
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayProjectParcels) = False Then
                    MsgBox("Не удалось найти участки рекультивации!!! Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем обозначения участков
                Dim dictionaryParcel As Dictionary(Of String, LandReclamation()) = New Dictionary(Of String, LandReclamation())
                For i As Integer = 0 To ArrayProjectParcels.GetUpperBound(1)
                    Dim PlineHandle As String = ArrayProjectParcels(0, i) 'Handle Контура
                    Dim tableProjectPS As String = ArrayProjectParcels(1, i) 'таблица свойств
                    'получаем ID
                    Dim idProjectParcels As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                    If IsNothing(idProjectParcels) = True Then Continue For
                    If idProjectParcels.IsNull = True Then Continue For
                    'получаем полилинию
                    Dim projectParcel As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(idProjectParcels, OpenMode.ForRead)
                    Dim AreaParcel As Double = projectParcel.Area
                    'читаем расширенные свойства
                    Dim LocalID As String = ""
                    Dim ConditionParcel As String = "1"
                    Dim CadastrtalNumber As String = "00:00:0000000:0000"
                    Dim Purpose As String = "Для переустройства инженерных сетей"
                    Dim Patronymic As String = ""
                    Dim Adress As String = "-"
                    Dim Category As String = "Не установлена"
                    Dim Utilization As String = "-"
                    Dim note As String = "-"
                    Dim direction As String = "Сельскохозяйственное"
                    Dim thLayer As Double = 0
                    Dim badLayer As Double = 0
                    Dim Biological As Integer = 0
                    Dim PKParcel As Double = 0
                    Dim technicalScheme As Integer = 0
                    Dim biologicalScheme As Integer = 0
                    'читаем свойства
                    Dim NameTBLPS As String = ""
                    Dim arrayData As String(,) = {} 'массив со свойствами участка
                    Dim boolXRecord As Boolean = FuncPropertySet.FuncReadDataSetPropertiesTbl(projectParcel, arrayData, tableProjectPS)
                    If IsArray(arrayData) = True Then
                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim NField As String = arrayData(1, j)
                            Dim Zn As String = arrayData(2, j)
                            If IsNothing(NField) = True Then Continue For
                            If NField.Trim.Length = 0 Then Continue For
                            If IsNothing(Zn) = True Then Zn = ""
                            If NField Like "LOCALID" Then
                                LocalID = Zn
                                NameTBLPS = arrayData(0, j)
                            ElseIf NField Like "ConditionalNumber" Then
                                ConditionParcel = Zn
                            ElseIf NField Like "Cad_number" Then
                                CadastrtalNumber = Zn
                            ElseIf NField Like "Purpose" Then
                                If Zn.Trim.Length > 0 Then
                                    Purpose = Zn
                                End If
                            ElseIf NField Like "Address" Then
                                If Zn.Trim.Length > 0 Then
                                    Adress = Zn
                                End If
                            ElseIf NField Like "Category" Then
                                If Zn.Trim.Length > 0 Then
                                    Category = Zn
                                End If
                            ElseIf NField Like "Utilization" Then
                                If Zn.Trim.Length > 0 Then
                                    Utilization = Zn
                                End If
                            ElseIf NField Like "Patronymic" Then
                                If Zn.Trim.Length > 0 Then
                                    Patronymic = Zn
                                End If
                            ElseIf NField Like "Direction" Then
                                If Zn.Trim.Length > 0 Then
                                    direction = Zn
                                End If
                            ElseIf NField Like "ThLayer" Then
                                If Zn.Trim.Length > 0 Then
                                    thLayer = Val(Zn)
                                End If
                            ElseIf NField Like "BadThLayer" Then
                                If Zn.Trim.Length > 0 Then
                                    badLayer = Val(Zn)
                                End If
                            ElseIf NField Like "TechnicalScheme" Then
                                If Zn.Trim.Length > 0 Then
                                    technicalScheme = Val(Zn)
                                End If
                            ElseIf NField Like "BiologicalScheme" Then
                                If Zn.Trim.Length > 0 Then
                                    biologicalScheme = Val(Zn)
                                End If
                            ElseIf NField Like "NOTE" Then
                                If Zn.Trim.Length > 0 Then
                                    note = Zn
                                End If
                            End If
                        Next j
                    End If
                    '========================================================================================================
                    'определяем внутреннюю точку полигона
                    Dim posProjectParcel As Point3d = New Point3d(-1, -1, -1)
                    Dim pkProjectParcel As Double = -1
                    Dim boolInsPoint As Boolean = False
                    Dim acMleader As MLeader = FuncCadastr.FuncFindMLeaderByLocalId(LocalID)
                    If IsNothing(acMleader) = False Then
                        If acMleader.Id.IsNull = False Then
                            posProjectParcel = acMleader.GetFirstVertex(0)
                            boolInsPoint = MathFunction.FuncCrossPointPolyline(projectParcel, posProjectParcel.Convert2d(New Plane))
                        End If
                    End If
                    If boolInsPoint = False Then
                        posProjectParcel = MathFunction.FuncFindCrossCenterPointPolyline(projectParcel)
                        boolInsPoint = MathFunction.FuncCrossPointPolyline(projectParcel, posProjectParcel.Convert2d(New Plane))
                    End If
                    '==============================================================================================================
                    'определяем пикет
                    If IsNothing(Align) = False Then
                        Dim off As Double = -1
                        'определяем пикет относительно трассы
                        If boolInsPoint = True Then
                            Try
                                Align.StationOffset(posProjectParcel.X, posProjectParcel.Y, pkProjectParcel, off)
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                        End If
                        'если пикет не удалос определить, то делаем более сложный поиск пикетажа о вершинам
                        If pkProjectParcel = -1 And off = -1 Then
                            For j As Integer = 0 To projectParcel.NumberOfVertices - 1
                                Try
                                    Align.StationOffset(projectParcel.GetPoint2dAt(j).X, projectParcel.GetPoint2dAt(j).Y, pkProjectParcel, off)
                                    Exit For
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            Next j
                        End If
                    End If
                    'ищем островки
                    Dim ArrayOstrov1 As String(,) = Nothing
                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov1)
                    If IsArray(ArrayOstrov1) = True Then
                        For k As Integer = 0 To ArrayOstrov1.GetUpperBound(1)
                            'получаем доступ к полилинии для извлечения координат
                            Dim HandlePLine1 As String = ArrayOstrov1(1, k)
                            Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine1)
                            If IsNothing(idOstrov) = False Then
                                If idOstrov.IsNull = False Then
                                    Dim ostrovProjectParcel As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                    AreaParcel = AreaParcel - ostrovProjectParcel.Area
                                End If
                            End If
                        Next k
                    End If
                    '==============================================================================================
                    'считываем схему биологической рекультивации
                    Dim numberYear As Integer = ParseYearReclamationXml(templateShemeXML, "biologicalStage", biologicalScheme)
                    If numberYear = 0 Then numberYear = 1
                    Dim rectamationParcel As LandReclamation = New LandReclamation(AreaParcel, thLayer, badLayer)
                    For numYaear As Integer = 1 To numberYear
                        Dim dictBiologicalSheme As List(Of Dictionary(Of String, String)) = ParseReclamationXml(templateShemeXML, "biologicalStage", biologicalScheme, numYaear)
                        If dictBiologicalSheme.Count > 0 Then
                            For j As Integer = 0 To dictBiologicalSheme.Count - 1
                                Dim dictBiological As Dictionary(Of String, String) = dictBiologicalSheme.Item(j)
                                '======================================================================================================
                                'минеральные удобрения
                                If dictBiological.ContainsKey("MineralFertilizers") = True Then
                                    Dim tempMineralFertilizers As String = dictBiological.Item("MineralFertilizers")
                                    Dim dictMineralFertilizers As New Dictionary(Of String, Double)
                                    If tempMineralFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempMineralFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictMineralFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictMineralFertilizers.Count >= k + 1 Then
                                                        Dim name = dictMineralFertilizers.ElementAt(k).Key
                                                        dictMineralFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictMineralFertilizers.Count > 0 Then
                                        rectamationParcel.MineralFertilizers.Add(dictMineralFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'органические удобрения
                                If dictBiological.ContainsKey("OrganicalFertilizers") = True Then
                                    Dim tempOrganicalFertilizers As String = dictBiological.Item("OrganicalFertilizers")
                                    Dim dictOrganicalFertilizers As New Dictionary(Of String, Double)
                                    If tempOrganicalFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempOrganicalFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictOrganicalFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictOrganicalFertilizers.Count >= k + 1 Then
                                                        Dim name = dictOrganicalFertilizers.ElementAt(k).Key
                                                        dictOrganicalFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictOrganicalFertilizers.Count > 0 Then
                                        rectamationParcel.OrganicalFertilizers.Add(dictOrganicalFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'сидераты
                                If dictBiological.ContainsKey("GreenSiderat") = True Then
                                    Dim tempGreenSiderat As String = dictBiological.Item("GreenSiderat")
                                    Dim dictGreenSiderat As New Dictionary(Of String, Double)
                                    If tempGreenSiderat.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempGreenSiderat.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictGreenSiderat.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictGreenSiderat.Count >= k + 1 Then
                                                        Dim name = dictGreenSiderat.ElementAt(k).Key
                                                        dictGreenSiderat.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictGreenSiderat.Count > 0 Then
                                        rectamationParcel.GreenSiderats.Add(dictGreenSiderat)
                                    End If
                                End If
                            Next j
                        End If
                    Next numYaear

                    '==============================================================================================
                    'записываем результат в словарь
                    If Category.Trim.Length = 0 Then Category = "Не установлена"
                    If dictionaryParcel.ContainsKey(Category) = False Then
                        dictionaryParcel.Add(Category, {rectamationParcel})
                    Else
                        Dim propertiesCounter As LandReclamation() = dictionaryParcel.Item(Category)
                        Dim count As Integer = propertiesCounter.Length
                        ReDim Preserve propertiesCounter(count)
                        propertiesCounter(count) = rectamationParcel
                        dictionaryParcel.Item(Category) = propertiesCounter
                    End If
                Next i
                '==================================================================================================
                If dictionaryParcel.Count = 0 Then
                    MsgBox("Фильтр не выявил участков рекультивации!!!")
                    Exit Sub
                End If
                '
                '========================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'записываем отчет
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Word документ
                Dim WApp As Word.Application = Nothing
                Dim WDoc As Word.Document = Nothing
                Dim rng As Random = New Random
                Dim boolWDVisibl As Boolean = False
                Try
                    WApp = CreateObject("Word.Application")
                    WDoc = WApp.Documents.Open(PutchTemplate & "Пояснительная записка.docx")
                    Dim Numbrng As Integer = rng.Next
                    Dim TempDir As String = Path.GetTempPath()
                    WDoc.SaveAs2(TempDir & "Пояснительная записка_" & Numbrng & ".docx")
                Catch ex As Exception
                    MsgBox("Файл-шаблон не найден.")
                    WDoc = Nothing
                    WApp.Quit()
                    WApp = Nothing
                    Exit Sub
                End Try
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'начинаем перебирать таблицы в документе Word
                WApp.Visible = True
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'график занимаемых земель
                Dim formatRow As Boolean = False
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "График занимаемых земель" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 2 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalBadThLayer As Double = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                tabl.Rows.Add()
                                Dim countTableword As Integer = tabl.Rows.Count
                                formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель
                                Dim maxThLayer As Double = 0
                                Dim minThLayer As Double = 100
                                Dim areaPercel As Double = 0
                                Dim badThLayer As Double = 0
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area
                                    areaPercel += tempArea
                                    Dim tempThLayer As Double = landParcelReclamation.thicknessLayer + landParcelReclamation.badThicknessLayer
                                    If tempThLayer > maxThLayer Then
                                        maxThLayer = tempThLayer
                                    End If
                                    If tempThLayer < minThLayer Then
                                        minThLayer = tempThLayer
                                    End If
                                    badThLayer += landParcelReclamation.badThicknessLayer * tempArea
                                Next j
                                'площадь
                                tabl.Cell(countTableword - 1, 2).Range.Text = Math.Round(areaPercel)
                                'мощность слоя
                                Dim thLayerStr As String = Math.Round(minThLayer, 2)
                                If maxThLayer > minThLayer Then
                                    thLayerStr = thLayerStr & "-" & Math.Round(maxThLayer, 2)
                                End If
                                tabl.Cell(countTableword - 1, 3).Range.Text = thLayerStr
                                'плохой грунт
                                If badThLayer > 0 Then
                                    Dim strBadLayer As String = "Объем плодородного грунта " & Math.Round(badThLayer, 2) & "куб.м., не пригоден для целей рекультивации."
                                    tabl.Cell(countTableword - 1, 4).Range.Text = strBadLayer
                                Else
                                    tabl.Cell(countTableword - 1, 4).Range.Text = "-"
                                End If
                                totalArea += areaPercel
                                totalBadThLayer += badThLayer
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword2)
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 2).Range.Text = Math.Round(totalArea, 0)
                        tabl.Cell(countTableword2, 4).Range.Text = Math.Round(totalBadThLayer, 2)
                        Exit For
                    End If
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'объем ПРС
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "Объемы ПРС" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 2 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalThLayer As Double = 0
                        Dim totalBadThLayer As Double = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                'добавляем строку
                                tabl.Rows.Add()
                                Dim countTableword As Integer = tabl.Rows.Count
                                formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель
                                Dim maxThLayer As Double = 0
                                Dim minThLayer As Double = 100
                                Dim areaPercel As Double = 0
                                Dim badThLayer As Double = 0
                                Dim thLayer As Double = 0
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area
                                    areaPercel += tempArea
                                    Dim tempThLayer As Double = landParcelReclamation.thicknessLayer + landParcelReclamation.badThicknessLayer
                                    If tempThLayer > maxThLayer Then
                                        maxThLayer = tempThLayer
                                    End If
                                    If tempThLayer < minThLayer Then
                                        minThLayer = tempThLayer
                                    End If
                                    thLayer += landParcelReclamation.thicknessLayer * tempArea
                                    badThLayer += landParcelReclamation.badThicknessLayer * tempArea
                                Next j
                                'площадь
                                tabl.Cell(countTableword - 1, 2).Range.Text = Math.Round(areaPercel / 10000, 4)
                                'мощность слоя
                                Dim thLayerStr As String = Math.Round(minThLayer, 2)
                                If maxThLayer > minThLayer Then
                                    thLayerStr = thLayerStr & "-" & Math.Round(maxThLayer, 2)
                                End If
                                tabl.Cell(countTableword - 1, 3).Range.Text = thLayerStr
                                'хороший грунт
                                tabl.Cell(countTableword - 1, 4).Range.Text = Math.Round(thLayer, 2)
                                'плохой грунт
                                tabl.Cell(countTableword - 1, 5).Range.Text = Math.Round(badThLayer, 2)
                                tabl.Cell(countTableword - 1, 6).Range.Text = Math.Round(badThLayer, 2)
                                totalArea += areaPercel
                                totalBadThLayer += badThLayer
                                totalThLayer += thLayer
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword2)
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 2).Range.Text = Math.Round(totalArea / 10000, 4)
                        tabl.Cell(countTableword2, 4).Range.Text = Math.Round(totalThLayer, 2)
                        tabl.Cell(countTableword2, 5).Range.Text = Math.Round(totalBadThLayer, 2)
                        tabl.Cell(countTableword2, 6).Range.Text = Math.Round(totalBadThLayer, 2)
                        Exit For
                    End If
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'объем органических удобрений OrganicalFert
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "Органические удобрения" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 2 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalArrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                        Dim countTotalArrayFert As Integer = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            'берем участки с определенной категории земель
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                Dim arrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                                Dim countArrayFert As Integer = 0
                                Dim areaPercel As Double = 0 'общая площадь
                                'читаем характеристи
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area / 10000 'площадь
                                    Dim listOrganicalFert As List(Of Dictionary(Of String, Double)) = landParcelReclamation.OrganicalFertilizers 'органические удобрения
                                    'список с органическими удобрениями не пустой, продолжаем
                                    If listOrganicalFert.Count > 0 Then
                                        'читаем все удобрения и их нормы
                                        For k As Integer = 0 To listOrganicalFert.Count - 1
                                            Dim dictValueFert As Dictionary(Of String, Double) = listOrganicalFert(k)
                                            For k1 As Integer = 0 To dictValueFert.Count - 1
                                                Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                                Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                                'делаем расчет количества
                                                Dim rezultFert As Double = normaFert * tempArea
                                                'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                                Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayFert, nameFert, 0)
                                                If indexNameFert = -1 Then
                                                    ReDim Preserve arrayFert(2, countArrayFert)
                                                    arrayFert(0, countArrayFert) = nameFert
                                                    arrayFert(1, countArrayFert) = normaFert
                                                    arrayFert(2, countArrayFert) = rezultFert
                                                    countArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(arrayFert(2, indexNameFert)) + rezultFert
                                                    arrayFert(2, indexNameFert) = tempRez
                                                End If
                                                'записываем результат в общий массив
                                                Dim totalIndexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(totalArrayFert, nameFert, 0)
                                                If totalIndexNameFert = -1 Then
                                                    ReDim Preserve totalArrayFert(2, countTotalArrayFert)
                                                    totalArrayFert(0, countTotalArrayFert) = nameFert
                                                    totalArrayFert(1, countTotalArrayFert) = normaFert
                                                    totalArrayFert(2, countTotalArrayFert) = rezultFert
                                                    countTotalArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(totalArrayFert(2, totalIndexNameFert)) + rezultFert
                                                    totalArrayFert(2, totalIndexNameFert) = tempRez
                                                End If
                                            Next k1
                                        Next k
                                    End If
                                    areaPercel += tempArea
                                Next j
                                'назначена органика
                                If arrayFert.Length > 0 Then
                                    tabl.Rows.Add()
                                    Dim countTableword As Integer = tabl.Rows.Count
                                    formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                    If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                    tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель
                                    'площадь
                                    areaPercel /= 10000
                                    tabl.Cell(countTableword - 1, 2).Range.Text = Math.Round(areaPercel, 4)
                                    Dim strNameY As String = ""
                                    Dim strNormaY As String = ""
                                    Dim summY As String = ""
                                    For k As Integer = 0 To arrayFert.GetUpperBound(1)
                                        If k = 0 Then
                                            strNameY = arrayFert(0, k)
                                            strNormaY = arrayFert(1, k)
                                            summY = arrayFert(2, k)
                                        Else
                                            strNameY = strNameY & "/" & arrayFert(0, k)
                                            strNormaY = strNormaY & "/" & arrayFert(1, k)
                                            summY = summY & "/" & arrayFert(2, k)
                                        End If
                                    Next k
                                    'наименование
                                    tabl.Cell(countTableword - 1, 2).Range.Text = strNameY
                                    'норма внесения
                                    tabl.Cell(countTableword - 1, 3).Range.Text = strNormaY
                                    'значение
                                    tabl.Cell(countTableword - 1, 4).Range.Text = summY
                                    totalArea += areaPercel
                                End If
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword2)
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 2).Range.Text = Math.Round(totalArea, 4)
                        Dim totalSummFert As String = ""
                        For k As Integer = 0 To totalArrayFert.GetUpperBound(1)
                            If k = 0 Then
                                totalSummFert = totalArrayFert(2, k)
                            Else
                                totalSummFert = totalSummFert & "/" & totalArrayFert(2, k)
                            End If
                        Next k
                        tabl.Cell(countTableword2, 4).Range.Text = totalSummFert
                        Exit For
                    End If
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'объем минеральных удобрений
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "Минеральные удобрения" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 3 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalArrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                        Dim countTotalArrayFert As Integer = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            'берем участки с определенной категории земель
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                Dim arrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                                Dim countArrayFert As Integer = 0
                                Dim areaPercel As Double = 0 'общая площадь
                                'читаем характеристи
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area  'площадь
                                    Dim listOrganicalFert As List(Of Dictionary(Of String, Double)) = landParcelReclamation.MineralFertilizers 'органические удобрения
                                    'список с органическими удобрениями не пустой, продолжаем
                                    If listOrganicalFert.Count > 0 Then
                                        'читаем все удобрения и их нормы
                                        For k As Integer = 0 To listOrganicalFert.Count - 1
                                            Dim dictValueFert As Dictionary(Of String, Double) = listOrganicalFert(k)
                                            For k1 As Integer = 0 To dictValueFert.Count - 1
                                                Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                                Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                                'делаем расчет количества
                                                Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                                'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                                Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayFert, nameFert, 0)
                                                If indexNameFert = -1 Then
                                                    ReDim Preserve arrayFert(2, countArrayFert)
                                                    arrayFert(0, countArrayFert) = nameFert
                                                    arrayFert(1, countArrayFert) = normaFert
                                                    arrayFert(2, countArrayFert) = rezultFert
                                                    countArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(arrayFert(2, indexNameFert)) + rezultFert
                                                    arrayFert(2, indexNameFert) = tempRez
                                                End If
                                                'записываем результат в общий массив
                                                Dim totalIndexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(totalArrayFert, nameFert, 0)
                                                If totalIndexNameFert = -1 Then
                                                    ReDim Preserve totalArrayFert(2, countTotalArrayFert)
                                                    totalArrayFert(0, countTotalArrayFert) = nameFert
                                                    totalArrayFert(1, countTotalArrayFert) = normaFert
                                                    totalArrayFert(2, countTotalArrayFert) = rezultFert
                                                    countTotalArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(totalArrayFert(2, totalIndexNameFert)) + rezultFert
                                                    totalArrayFert(2, totalIndexNameFert) = tempRez
                                                End If
                                            Next k1
                                        Next k
                                    End If
                                    areaPercel += tempArea
                                Next j
                                'назначена органика
                                If arrayFert.Length > 0 Then
                                    tabl.Rows.Add()
                                    Dim countTableword As Integer = tabl.Rows.Count
                                    formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                    If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                    tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель

                                    Dim strNameY As String = ""
                                    Dim strNormaY As String = ""
                                    Dim summY As String = ""
                                    For k As Integer = 0 To arrayFert.GetUpperBound(1)
                                        If k = 0 Then
                                            strNameY = arrayFert(0, k)
                                            strNormaY = arrayFert(1, k)
                                            Dim tempSumb As Double = Val(arrayFert(2, k))
                                            tempSumb = Math.Round(tempSumb, 2)
                                            summY = tempSumb
                                        Else
                                            strNameY = strNameY & "/" & arrayFert(0, k)
                                            strNormaY = strNormaY & "/" & arrayFert(1, k)
                                            Dim tempSumb As Double = Val(arrayFert(2, k))
                                            tempSumb = Math.Round(tempSumb, 2)
                                            summY = summY & "/" & tempSumb
                                        End If
                                    Next k
                                    'наименование
                                    tabl.Cell(countTableword - 1, 2).Range.Text = strNameY
                                    'площадь
                                    areaPercel /= 10000
                                    tabl.Cell(countTableword - 1, 3).Range.Text = Math.Round(areaPercel, 4)
                                    'норма внесения
                                    tabl.Cell(countTableword - 1, 4).Range.Text = strNormaY
                                    'значение
                                    tabl.Cell(countTableword - 1, 5).Range.Text = summY
                                    totalArea += areaPercel
                                End If
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword2)
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 3).Range.Text = Math.Round(totalArea, 4)
                        Dim totalSummFert As String = ""
                        For k As Integer = 0 To totalArrayFert.GetUpperBound(1)
                            If k = 0 Then
                                Dim tempSumb As Double = Val(totalArrayFert(2, k))
                                tempSumb = Math.Round(tempSumb, 2)
                                totalSummFert = tempSumb
                            Else
                                Dim tempSumb As Double = Val(totalArrayFert(2, k))
                                tempSumb = Math.Round(tempSumb, 2)
                                totalSummFert = totalSummFert & "/" & tempSumb
                            End If
                        Next k
                        tabl.Cell(countTableword2, 5).Range.Text = totalSummFert
                        Exit For
                    End If
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'Расчет потребности семян
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "Расчет потребности семян" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 2 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalArrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                        Dim countTotalArrayFert As Integer = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            'берем участки с определенной категории земель
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                Dim arrayFert As String(,) = {} '0-наименование 1-норма внесения 2-количество
                                Dim countArrayFert As Integer = 0
                                Dim areaPercel As Double = 0 'общая площадь
                                'читаем характеристи
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area  'площадь
                                    Dim listOrganicalFert As List(Of Dictionary(Of String, Double)) = landParcelReclamation.GreenSiderats 'органические удобрения
                                    'список с органическими удобрениями не пустой, продолжаем
                                    If listOrganicalFert.Count > 0 Then
                                        'читаем все удобрения и их нормы
                                        For k As Integer = 0 To listOrganicalFert.Count - 1
                                            Dim dictValueFert As Dictionary(Of String, Double) = listOrganicalFert(k)
                                            For k1 As Integer = 0 To dictValueFert.Count - 1
                                                Dim nameFert As String = dictValueFert.ElementAt(k1).Key
                                                Dim normaFert As Double = dictValueFert.ElementAt(k1).Value
                                                'делаем расчет количества
                                                Dim rezultFert As Double = normaFert * (tempArea / 10000)
                                                'находим уже это удобрение в массиве и если его нет то добавляем, иначе только ссумирем результат
                                                Dim indexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(arrayFert, nameFert, 0)
                                                If indexNameFert = -1 Then
                                                    ReDim Preserve arrayFert(2, countArrayFert)
                                                    arrayFert(0, countArrayFert) = nameFert
                                                    arrayFert(1, countArrayFert) = normaFert
                                                    arrayFert(2, countArrayFert) = rezultFert
                                                    countArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(arrayFert(2, indexNameFert)) + rezultFert
                                                    arrayFert(2, indexNameFert) = tempRez
                                                End If
                                                'записываем результат в общий массив
                                                Dim totalIndexNameFert As Integer = MathFunction.FuncFindIndexValueToArray2d(totalArrayFert, nameFert, 0)
                                                If totalIndexNameFert = -1 Then
                                                    ReDim Preserve totalArrayFert(2, countTotalArrayFert)
                                                    totalArrayFert(0, countTotalArrayFert) = nameFert
                                                    totalArrayFert(1, countTotalArrayFert) = normaFert
                                                    totalArrayFert(2, countTotalArrayFert) = rezultFert
                                                    countTotalArrayFert += 1
                                                Else
                                                    Dim tempRez As Double = Val(totalArrayFert(2, totalIndexNameFert)) + rezultFert
                                                    totalArrayFert(2, totalIndexNameFert) = tempRez
                                                End If
                                            Next k1
                                        Next k
                                    End If
                                    areaPercel += tempArea
                                Next j
                                'назначена органика
                                If arrayFert.Length > 0 Then
                                    tabl.Rows.Add()
                                    Dim countTableword As Integer = tabl.Rows.Count
                                    formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                    If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                    tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель
                                    Dim strNameY As String = ""
                                    Dim strNormaY As String = ""
                                    Dim summY As String = ""
                                    For k As Integer = 0 To arrayFert.GetUpperBound(1)
                                        If k = 0 Then
                                            strNameY = arrayFert(0, k)
                                            strNormaY = arrayFert(1, k)
                                            Dim tempSumb As Double = Val(arrayFert(2, k))
                                            tempSumb = Math.Round(tempSumb, 2)
                                            summY = tempSumb
                                        Else
                                            strNameY = strNameY & "/" & arrayFert(0, k)
                                            strNormaY = strNormaY & "/" & arrayFert(1, k)
                                            Dim tempSumb As Double = Val(arrayFert(2, k))
                                            tempSumb = Math.Round(tempSumb, 2)
                                            summY = summY & "/" & tempSumb
                                        End If
                                    Next k
                                    'наименование
                                    tabl.Cell(countTableword - 1, 2).Range.Text = strNameY
                                    'площадь
                                    areaPercel /= 10000
                                    tabl.Cell(countTableword - 1, 3).Range.Text = Math.Round(areaPercel, 4)
                                    'норма внесения
                                    tabl.Cell(countTableword - 1, 4).Range.Text = strNormaY
                                    'значение
                                    tabl.Cell(countTableword - 1, 5).Range.Text = summY
                                    totalArea += areaPercel
                                End If
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword2)
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 3).Range.Text = Math.Round(totalArea, 3)
                        Dim totalSummFert As String = ""
                        For k As Integer = 0 To totalArrayFert.GetUpperBound(1)
                            If k = 0 Then
                                Dim tempSumb As Double = Val(totalArrayFert(2, k))
                                tempSumb = Math.Round(tempSumb, 2)
                                totalSummFert = tempSumb
                            Else
                                Dim tempSumb As Double = Val(totalArrayFert(2, k))
                                tempSumb = Math.Round(tempSumb, 2)
                                totalSummFert = totalSummFert & "/" & tempSumb
                            End If
                        Next k
                        tabl.Cell(countTableword2, 5).Range.Text = totalSummFert
                        Exit For
                    End If
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'стоимость проведения рекультивации
                For Each tabl As Word.Table In WDoc.Tables
                    'читаем заголовок таблицы
                    If tabl.Title Like "Сметная стоимость" Then
                        Dim rowCount As Integer = tabl.Rows.Count
                        If rowCount > 2 Then
                            For i As Integer = rowCount To 2 Step -1
                                tabl.Rows(i - 1).Delete()
                            Next
                        End If
                        Dim totalArea As Double = 0
                        Dim totalCalculateLucerna As Double = 0
                        For i As Integer = 0 To dictionaryParcel.Count - 1
                            Dim arrayPropertiesCounter As LandReclamation() = dictionaryParcel.ElementAt(i).Value
                            If IsArray(arrayPropertiesCounter) = True Then
                                Dim areaPercel As Double = 0
                                For j As Integer = 0 To arrayPropertiesCounter.Length - 1
                                    Dim landParcelReclamation As LandReclamation = arrayPropertiesCounter(j)
                                    Dim tempArea As Double = landParcelReclamation.Area
                                    areaPercel += tempArea
                                Next j
                                If areaPercel > 0 Then
                                    tabl.Rows.Add()
                                    Dim countTableword As Integer = tabl.Rows.Count
                                    formatRow = FuncFormatZN.FormatTableRow(tabl, countTableword)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    Dim categoryName As String = dictionaryParcel.ElementAt(i).Key
                                    If categoryName Like "Не установлена" Then categoryName = "Земли неразграниченной государсвенной собственности"
                                    tabl.Cell(countTableword - 1, 1).Range.Text = categoryName 'категория земель
                                    'площадь
                                    tabl.Cell(countTableword - 1, 2).Range.Text = Math.Round(areaPercel, 4)
                                    totalArea += areaPercel
                                End If
                            End If
                        Next i
                        Dim countTableword2 As Integer = tabl.Rows.Count
                        tabl.Cell(countTableword2, 1).Range.Text = "Итого: "
                        tabl.Cell(countTableword2, 2).Range.Text = Math.Round(totalArea, 4)
                        Exit For
                    End If
                Next
                WDoc = Nothing
                WApp.Quit()
                WApp = Nothing
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub

    '6. Ведомость объёмов работ для смет
    <CommandMethod("LandReclamationVolumeSmeta")>
    Public Sub LandReclamationVolumeSmeta()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayAllPoligon As String(,) = Nothing 'массив для хранения всех участков, подучастков и островков на всем пространстве чертежа (handle, nameOD)
        Dim countArrayAllPoligon As Integer = 0
        Dim ArrayTablesXML As String(,) = Nothing
        ReDim ArrayTablesXML(1, 0)
        ArrayTablesXML(0, 0) = "LAND_RECLAMATION"
        ArrayTablesXML(1, 0) = "Рекультивация земель"
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '1. путь к шаблонам
        Dim PutchTemplate As String = "k:\Civil3DToolsUtility\TemplateTables\LandReclamation\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\LandReclamation\") Then
            PutchTemplate = parentPatch2 & "\TemplateTables\LandReclamation\"
        End If
        If IO.Directory.Exists(PutchTemplate) = False Then
            MsgBox("Шаблон таблицы не найден!!!")
            Exit Sub
        End If
        'технологическая схема
        Dim templateShemeXML As String = ""
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\Semantic\") Then
            templateShemeXML = parentPatch2 & "\FileResources\Sample\Semantic\"
            templateShemeXML = templateShemeXML & "LandReclamation.xml"
            If File.Exists(templateShemeXML) = False Then
                MsgBox("Шаблон оформления чертежей не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim dictPacelReclamation As Dictionary(Of String, LandReclamation()) = New Dictionary(Of String, LandReclamation())

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Try
            Using tr
                '=========================================================================================================
                'выбираем участки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                '===================================================================================================
                                'читаем расширенные данные
                                Dim ArrayNameTableOD As String() = Nothing
                                Dim boolTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acEnt, ArrayNameTableOD, ArrayTablesXML)
                                If IsArray(ArrayNameTableOD) = True Then
                                    For i As Integer = 0 To ArrayNameTableOD.Length - 1
                                        ReDim Preserve ArrayAllPoligon(1, countArrayAllPoligon)
                                        ArrayAllPoligon(0, countArrayAllPoligon) = acEnt.Handle.ToString
                                        ArrayAllPoligon(1, countArrayAllPoligon) = ArrayNameTableOD(i)
                                        countArrayAllPoligon += 1
                                    Next i
                                End If
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayAllPoligon) = False Then
                    MsgBox("Не удалось найти участки!!!")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем характеристики участков
                For i As Integer = 0 To ArrayAllPoligon.GetUpperBound(1)
                    Dim NameOD As String = ArrayAllPoligon(1, i)
                    Dim PlineHandle As String = ArrayAllPoligon(0, i)
                    Dim idProjectParcel As ObjectId = AcadFunction.FuncReturnObjectID(PlineHandle)
                    If IsNothing(idProjectParcel) = True Then Continue For
                    If idProjectParcel.IsNull = True Then Continue For
                    Dim projectParcel As Autodesk.AutoCAD.DatabaseServices.Polyline = tr.GetObject(idProjectParcel, OpenMode.ForRead)
                    'ищем глобальный и локальный id
                    Dim arrayData As String(,) = {}
                    Dim GlobalID As String = ""
                    Dim LocalID As String = ""
                    Dim ConditionParcel As String = ""
                    Dim biologicalSheme As Integer = 0
                    Dim technicalSheme As Integer = 0
                    Dim BadThLayer As Double = 0
                    Dim ThLayer As Double = 0
                    Dim boolXRecord As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(projectParcel, arrayData, ArrayTablesXML)
                    If IsArray(arrayData) = True Then
                        For j As Integer = 0 To arrayData.GetUpperBound(1)
                            Dim NField As String = arrayData(1, j)
                            Dim Zn As String = arrayData(2, j)
                            If Zn.Trim.Length > 0 Then
                                If NField Like "GLOBALID" Then
                                    GlobalID = Zn
                                ElseIf NField Like "LOCALID" Then
                                    LocalID = Zn
                                ElseIf NField Like "ConditionalNumber" Then
                                    ConditionParcel = Zn
                                ElseIf NField Like "TechnicalScheme" Then
                                    technicalSheme = Val(Zn)
                                ElseIf NField Like "BiologicalScheme" Then
                                    biologicalSheme = Val(Zn)
                                ElseIf NField Like "BadThLayer" Then
                                    BadThLayer = Val(Zn)
                                ElseIf NField Like "ThLayer" Then
                                    ThLayer = Val(Zn)
                                End If
                            End If
                        Next j
                    End If
                    '===============================================================================
                    'отделяем номер участка от его префикса
                    Dim numbParcel As Integer = FuncCadastr.FuncFindShortNumberCadastralObject(ConditionParcel)
                    'определяем площадь
                    Dim AreaParcel As Double = projectParcel.Area
                    'ищем островки
                    Dim ArrayOstrov As String(,) = Nothing
                    Dim boolOstrov As Boolean = FuncCadastr.FuncFindOstrov(LocalID, ArrayOstrov)
                    If IsArray(ArrayOstrov) = True Then
                        For k As Integer = 0 To ArrayOstrov.GetUpperBound(1)
                            'получаем доступ к полилинии для извлечения координат
                            Dim HandlePLine1 As String = ArrayOstrov(1, k)
                            Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(HandlePLine1)
                            If IsNothing(idOstrov) = False Then
                                If idOstrov.IsNull = False Then
                                    Dim ostrovPline As Polyline = tr.GetObject(idOstrov, OpenMode.ForRead)
                                    AreaParcel = AreaParcel - ostrovPline.Area
                                End If
                            End If
                        Next k
                    End If
                    'находим технологическую схему проведения биологической рекультивации
                    '==============================================================================================
                    'считываем схему биологической рекультивации
                    Dim numberYear As Integer = ParseYearReclamationXml(templateShemeXML, "biologicalStage", biologicalSheme)
                    If numberYear = 0 Then numberYear = 1
                    Dim rectamationParcel As LandReclamation = New LandReclamation(AreaParcel, ThLayer, BadThLayer)
                    For numYaear As Integer = 1 To numberYear
                        Dim dictBiologicalSheme As List(Of Dictionary(Of String, String)) = ParseReclamationXml(templateShemeXML, "biologicalStage", biologicalSheme, numYaear)
                        If dictBiologicalSheme.Count > 0 Then
                            For j As Integer = 0 To dictBiologicalSheme.Count - 1
                                Dim dictBiological As Dictionary(Of String, String) = dictBiologicalSheme.Item(j)
                                '======================================================================================================
                                'мелиоранты
                                If dictBiological.ContainsKey("Ameliorant") = True Then
                                    Dim tempAmeliorant As String = dictBiological.Item("Ameliorant")
                                    Dim dictAmeliorant As New Dictionary(Of String, Double)
                                    If tempAmeliorant.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempAmeliorant.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictAmeliorant.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictAmeliorant.Count >= k + 1 Then
                                                        Dim name = dictAmeliorant.ElementAt(k).Key
                                                        dictAmeliorant.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictAmeliorant.Count > 0 Then
                                        rectamationParcel.MineralFertilizers.Add(dictAmeliorant)
                                    End If
                                End If
                                '======================================================================================================
                                'минеральные удобрения
                                If dictBiological.ContainsKey("MineralFertilizers") = True Then
                                    Dim tempMineralFertilizers As String = dictBiological.Item("MineralFertilizers")
                                    Dim dictMineralFertilizers As New Dictionary(Of String, Double)
                                    If tempMineralFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempMineralFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictMineralFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictMineralFertilizers.Count >= k + 1 Then
                                                        Dim name = dictMineralFertilizers.ElementAt(k).Key
                                                        dictMineralFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictMineralFertilizers.Count > 0 Then
                                        rectamationParcel.MineralFertilizers.Add(dictMineralFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'органические удобрения
                                If dictBiological.ContainsKey("OrganicalFertilizers") = True Then
                                    Dim tempOrganicalFertilizers As String = dictBiological.Item("OrganicalFertilizers")
                                    Dim dictOrganicalFertilizers As New Dictionary(Of String, Double)
                                    If tempOrganicalFertilizers.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempOrganicalFertilizers.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictOrganicalFertilizers.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictOrganicalFertilizers.Count >= k + 1 Then
                                                        Dim name = dictOrganicalFertilizers.ElementAt(k).Key
                                                        dictOrganicalFertilizers.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictOrganicalFertilizers.Count > 0 Then
                                        rectamationParcel.OrganicalFertilizers.Add(dictOrganicalFertilizers)
                                    End If
                                End If
                                '======================================================================================================
                                'сидераты
                                If dictBiological.ContainsKey("GreenSiderat") = True Then
                                    Dim tempGreenSiderat As String = dictBiological.Item("GreenSiderat")
                                    Dim dictGreenSiderat As New Dictionary(Of String, Double)
                                    If tempGreenSiderat.Trim.Length > 0 Then
                                        Dim arrayFert As String() = tempGreenSiderat.Split(";")
                                        If IsArray(arrayFert) = True Then
                                            For k As Integer = 0 To arrayFert.Length - 1
                                                Dim nameFert As String = arrayFert(k)
                                                If nameFert.Trim.Length > 0 Then
                                                    dictGreenSiderat.Add(nameFert, 0)
                                                End If
                                            Next k
                                        End If

                                    End If
                                    'заполняем норму внесения
                                    If dictBiological.ContainsKey("ApplicationRate") = True Then
                                        Dim tempApplicationRate As String = dictBiological.Item("ApplicationRate")
                                        If tempApplicationRate.Trim.Length > 0 Then
                                            Dim arrayRate As String() = tempApplicationRate.Split(";")
                                            If IsArray(arrayRate) = True Then
                                                For k As Integer = 0 To arrayRate.Length - 1
                                                    Dim valueRate As Double = Val(arrayRate(k))
                                                    If dictGreenSiderat.Count >= k + 1 Then
                                                        Dim name = dictGreenSiderat.ElementAt(k).Key
                                                        dictGreenSiderat.Item(name) = valueRate
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    End If
                                    If dictGreenSiderat.Count > 0 Then
                                        rectamationParcel.GreenSiderats.Add(dictGreenSiderat)
                                    End If
                                End If
                            Next j
                        End If
                    Next numYaear
                    rectamationParcel.ConditionatNumber = numbParcel
                    Dim strClass As String = technicalSheme & "-" & biologicalSheme
                    If dictPacelReclamation.ContainsKey(strClass) = False Then
                        dictPacelReclamation.Add(strClass, {rectamationParcel})
                    Else
                        Dim arrayReclamationParcel As LandReclamation() = dictPacelReclamation.Item(strClass)
                        Dim count As Integer = arrayReclamationParcel.Length
                        ReDim Preserve arrayReclamationParcel(count)
                        arrayReclamationParcel(count) = rectamationParcel
                        dictPacelReclamation.Item(strClass) = arrayReclamationParcel
                    End If
                Next i
                tr.Commit()
            End Using

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '========================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Word документ
            'запускаем Excel
            Dim startRow As Integer = 17
            Dim countSheet As Integer = 1
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            If dictPacelReclamation.Count > 0 Then
                For i As Integer = 0 To dictPacelReclamation.Count - 1
                    Try
                        Dim strSheme As String = dictPacelReclamation.ElementAt(i).Key
                        Dim parcelReclamatrion As LandReclamation() = dictPacelReclamation.Item(strSheme)
                        Dim arraySheme As String() = strSheme.Split("-")
                        Dim numberShemeTechnicalReclamation As Integer = Val(arraySheme(0))
                        Dim numberShemeBiologicalReclamation As Integer = Val(arraySheme(1))
                        Dim countRow As Integer = 1
                        xlApp = CreateObject("Excel.Application")
                        xlBook = xlApp.Workbooks.Open(PutchTemplate & "Ведомость объемов работ.xlsx")
                        Dim rng As Random = New Random
                        Dim Numbrng As Integer = rng.Next
                        Dim TempDir As String = Path.GetTempPath()
                        xlBook.SaveAs(TempDir & "Ведомость объемов работ_" & Numbrng & ".xlsx")
                        xlApp.Visible = True
                        xlSheets = xlBook.Sheets.Item("ВР_")
                        Dim columnRange As Excel.Range = xlSheets.Columns(2)
                        columnRange.AutoFit()
                        Dim dictTechnicalSheme As List(Of Dictionary(Of String, String)) = ParseReclamationXml(templateShemeXML, "technicalStage", numberShemeTechnicalReclamation, 0)
                        If dictTechnicalSheme.Count > 0 Then
                            For j As Integer = 0 To dictTechnicalSheme.Count - 1
                                Dim listElement As Dictionary(Of String, String) = dictTechnicalSheme.Item(j)
                                Dim strWork As String = ""
                                Dim unit As String = ""
                                If listElement.ContainsKey("desk") = True Then
                                    strWork = listElement.Item("desk")
                                End If
                                Dim formula As String = ""
                                For k As Integer = 0 To listElement.Count - 1
                                    Dim namefield As String = listElement.ElementAt(k).Key
                                    Dim value As String = listElement.Item(namefield)
                                    If namefield Like "formula" Then
                                        formula = value
                                    ElseIf namefield Like "unit" Then
                                        unit = value
                                    Else
                                        strWork = strWork.Replace(namefield, value)
                                    End If
                                Next k
                                If formula.Trim.Length > 0 Then
                                    Dim valCalculate As Double = 0
                                    Dim arrayRezult As String(,) = calculateValue(formula, parcelReclamatrion, listElement)
                                    If arrayRezult.Length > 0 Then
                                        For k As Integer = 0 To arrayRezult.GetUpperBound(1)
                                            xlSheets.Rows(startRow + countRow).Insert()
                                            Dim rowRange As Excel.Range = xlSheets.Rows(startRow + countRow)
                                            rowRange.Font.Bold = False
                                            xlSheets.Rows(startRow + countRow).AutoFit() ' Автоподбор высоты строки
                                            xlSheets.Cells(startRow + countRow, 2).WrapText = True
                                            xlSheets.Cells(startRow + countRow, 5).WrapText = True
                                            xlSheets.Cells(startRow + countRow, 1) = countRow
                                            xlSheets.Cells(startRow + countRow, 2) = "Технический этап. " & strWork
                                            xlSheets.Cells(startRow + countRow, 3) = unit
                                            xlSheets.Cells(startRow + countRow, 4) = arrayRezult(2, 0)
                                            Dim strFormula As String = "S="
                                            If unit Like "м3" Or unit Like "ц" Or unit Like "кг" Then
                                                strFormula = "Vобщ=" & arrayRezult(3, k) & unit
                                            ElseIf unit Like "м2" Then
                                                strFormula = "Sобщ=" & arrayRezult(3, k) & unit
                                            End If
                                            xlSheets.Cells(startRow + countRow, 5) = strFormula
                                            countRow += 1
                                        Next k
                                    End If
                                End If
                            Next j
                        End If
                        'биологическая рекультивация
                        Dim numberYear As Integer = ParseYearReclamationXml(templateShemeXML, "biologicalStage", numberShemeBiologicalReclamation)
                        If numberYear = 0 Then numberYear = 1
                        For numYaear As Integer = 1 To numberYear
                            Dim dictBiologicalSheme As List(Of Dictionary(Of String, String)) = ParseReclamationXml(templateShemeXML, "biologicalStage", numberShemeBiologicalReclamation, numYaear)
                            If dictBiologicalSheme.Count > 0 Then
                                For j As Integer = 0 To dictBiologicalSheme.Count - 1
                                    Dim listElement As Dictionary(Of String, String) = dictBiologicalSheme.Item(j)
                                    Dim strWork As String = ""
                                    If listElement.ContainsKey("desk") = True Then
                                        strWork = listElement.Item("desk")
                                    End If
                                    Dim unit As String = ""
                                    Dim formula As String = ""
                                    For k As Integer = 0 To listElement.Count - 1
                                        Dim namefield As String = listElement.ElementAt(k).Key
                                        Dim value As String = listElement.Item(namefield)
                                        If namefield Like "formula" Then
                                            formula = value
                                        ElseIf namefield Like "unit" Then
                                            unit = value
                                        Else
                                            strWork = strWork.Replace(namefield, value)
                                        End If
                                    Next k
                                    If formula.Trim.Length > 0 Then
                                        Dim valCalculate As Double = 0
                                        Dim arrayRezult As String(,) = calculateValue(formula, parcelReclamatrion, listElement)
                                        If arrayRezult.Length > 0 Then
                                            For k As Integer = 0 To arrayRezult.GetUpperBound(1)
                                                xlSheets.Rows(startRow + countRow).Insert()
                                                Dim rowRange As Excel.Range = xlSheets.Rows(startRow + countRow)
                                                rowRange.Font.Bold = False
                                                xlSheets.Rows(startRow + countRow).AutoFit() ' Автоподбор высоты строки
                                                xlSheets.Cells(startRow + countRow, 2).WrapText = True
                                                xlSheets.Cells(startRow + countRow, 5).WrapText = True
                                                xlSheets.Cells(startRow + countRow, 1) = countRow
                                                If numberYear > 1 Then
                                                    xlSheets.Cells(startRow + countRow, 2) = "Билогический этап. Год " & numYaear & ". " & strWork & " " & arrayRezult(0, k)
                                                Else
                                                    xlSheets.Cells(startRow + countRow, 2) = "Билогический этап. " & strWork & " " & arrayRezult(0, k)
                                                End If
                                                xlSheets.Cells(startRow + countRow, 3) = unit
                                                xlSheets.Cells(startRow + countRow, 4) = arrayRezult(2, k)
                                                Dim strFormula As String = "S="
                                                If unit Like "м3" Or unit Like "ц" Or unit Like "кг" Then
                                                    strFormula = "Vобщ=" & arrayRezult(3, k) & unit
                                                ElseIf unit Like "м2" Then
                                                    strFormula = "Sобщ=" & arrayRezult(3, k) & unit
                                                End If
                                                xlSheets.Cells(startRow + countRow, 5) = strFormula
                                                countRow += 1
                                            Next k
                                        End If
                                    End If
                                Next j
                            End If
                        Next numYaear
                    Catch ex As Exception
                        MsgBox("Файл-шаблон не найден.")
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                        Exit Sub
                    Finally
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                    End Try
                Next i
            End If
        Catch
        End Try
    End Sub

End Class
